import { describe, expect, it, vi } from 'vitest'
import { graphUrl, MicrosoftGraphClient } from './microsoft-graph'

describe('Microsoft Graph mailbox adapter', () => {
  it('follows opaque pagination and requests immutable IDs on every page', async () => {
    const next = 'https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages?$skiptoken=opaque%2Bvalue'
    const fetcher = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      expect(new Headers(init?.headers).get('Prefer')).toBe('IdType="ImmutableId"')
      return Response.json(String(input) === next
        ? { value: [{ id: 'second', isRead: true }] }
        : { value: [{ id: 'first', subject: 'First' }], '@odata.nextLink': next })
    })
    const client = new MicrosoftGraphClient('access', false, fetcher)
    const messages = await client.listMessages('INBOX', 2)
    expect(messages.map(({ graphMessageId, uid }) => [graphMessageId, uid]))
      .toEqual([['first', null], ['second', null]])
    expect(fetcher.mock.calls[1][0]).toBe(next)
    expect(messages[1].isRead).toBe(true)
  })

  it('rejects off-origin continuation URLs before sending credentials', async () => {
    expect(() => graphUrl('https://evil.example/v1.0/me/messages')).toThrow('invalid_graph_url')
    expect(() => graphUrl('https://:userinfo@graph.microsoft.com/v1.0/me/messages')).toThrow('invalid_graph_url')
    expect(() => graphUrl('https://graph.microsoft.com/v1.0/users/another/messages')).toThrow('invalid_graph_url')
    const fetcher = vi.fn(async () => Response.json({ value: [], '@odata.nextLink': 'https://evil.example/' }))
    const client = new MicrosoftGraphClient('access', false, fetcher)
    await expect(client.listMessages('INBOX', 50)).rejects.toThrow('invalid_graph_url')
    expect(fetcher).toHaveBeenCalledTimes(1)
  })

  it('refreshes once on a rejected access token and respects throttling', async () => {
    const renew = vi.fn(async () => 'renewed')
    const fetcher = vi.fn(async (_input: RequestInfo | URL, init?: RequestInit) => (
      new Headers(init?.headers).get('Authorization') === 'Bearer access'
        ? new Response(null, { status: 401 })
        : new Response(null, { status: 429, headers: { 'Retry-After': '120' } })
    ))
    const client = new MicrosoftGraphClient('access', true, fetcher, renew)
    await expect(client.open()).rejects.toMatchObject({ code: 'graph_throttled', retryable: true, retryAfter: 120 })
    expect(renew).toHaveBeenCalledTimes(1)
    expect(fetcher).toHaveBeenCalledTimes(2)
  })

  it('reads MIME with encoded message identity and never marks read on download', async () => {
    const fetcher = vi.fn(async (input: RequestInfo | URL) => {
      expect(String(input)).toContain('/messages/id%2Fwith%2Bchars%3D/$value')
      return new Response('From: sender@example.com\r\nSubject: Hello\r\nContent-Type: text/plain\r\n\r\nBody')
    })
    const parsed = await new MicrosoftGraphClient('access', true, fetcher)
      .getMessage('INBOX', 'id/with+chars=')
    expect(parsed.message).toMatchObject({ subject: 'Hello', body: 'Body' })
    expect(fetcher).toHaveBeenCalledTimes(1)
  })

  it('blocks read-state writes when only Mail.Read was granted', async () => {
    const fetcher = vi.fn()
    await expect(new MicrosoftGraphClient('access', false, fetcher).markSeen('INBOX', 'message'))
      .rejects.toMatchObject({ code: 'graph_readonly' })
    expect(fetcher).not.toHaveBeenCalled()
  })

  it('traverses nested folders and keeps the localized inbox mapped to INBOX', async () => {
    const fetcher = vi.fn(async (input: RequestInfo | URL) => {
      const url = String(input)
      if (url.includes('/inbox?')) return Response.json({ id: 'inbox-id', displayName: '收件箱' })
      if (url.includes('/childFolders?')) return Response.json({ value: [{ id: 'child', displayName: 'Child', childFolderCount: 0 }] })
      return Response.json({ value: [{ id: 'inbox-id', displayName: '收件箱', childFolderCount: 1 }] })
    })
    const folders = await new MicrosoftGraphClient('access', true, fetcher).listFolders()
    expect(folders.map(({ path }) => path)).toEqual(['INBOX', 'graph:child'])
  })
})
