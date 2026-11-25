import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { ImapConnectionError } from '../../platform/imap/imap-errors'
import { validateMicrosoftImport } from './microsoft-account-api'
import type { ValidMicrosoftImport } from './microsoft-types'

const imap = vi.hoisted(() => ({ open: vi.fn(), close: vi.fn() }))
vi.mock('./microsoft-imap', () => ({
  MicrosoftImapClient: class {
    readonly protocol = 'imap'
    open = imap.open
    close = imap.close
  },
}))
const input: ValidMicrosoftImport = {
  email: 'mail@example.com', name: 'Mailbox', authMode: 'oauth2',
  password: null, refreshToken: 'original-refresh', authority: 'common',
  clientId: '00000000-0000-4000-8000-000000000000',
}

function remote(profileScope = false, profileMatches = true) {
  return vi.fn(async (url: RequestInfo | URL, init?: RequestInit) => {
    if (String(url).startsWith('https://login.microsoftonline.com/')) {
      const body = new URLSearchParams(String(init?.body))
      if (body.get('scope')?.includes('IMAP.AccessAsUser.All')) return Response.json({
        access_token: 'imap-access', refresh_token: 'imap-rotated', expires_in: 3600,
        scope: 'https://outlook.office.com/IMAP.AccessAsUser.All offline_access',
      })
      expect(imap.close).toHaveBeenCalledOnce()
      expect(body.get('refresh_token')).toBe('imap-rotated')
      return Response.json({ access_token: 'graph-access', refresh_token: 'graph-rotated', expires_in: 3600,
        scope: `https://graph.microsoft.com/Mail.Read${profileScope ? ' https://graph.microsoft.com/User.Read' : ''}` })
    }
    if (String(url).includes('/me/?')) return Response.json({ mail: profileMatches ? input.email : 'different@example.com' })
    if (String(url).includes('/inbox?')) return Response.json({ id: 'inbox', displayName: 'Inbox', totalItemCount: 0 })
    return Response.json({ value: [{ id: 'inbox', displayName: 'Inbox', childFolderCount: 0 }] })
  })
}

beforeEach(() => {
  imap.open.mockReset().mockRejectedValue(new ImapConnectionError(400, 'Authentication failed', true))
  imap.close.mockReset().mockResolvedValue(undefined)
})
afterEach(() => vi.unstubAllGlobals())

describe('Microsoft import transport validation', () => {
  it('accepts Mail.Read after IMAP login rejection without requiring User.Read', async () => {
    const fetcher = remote()
    vi.stubGlobal('fetch', fetcher)
    await expect(validateMicrosoftImport(input)).resolves.toMatchObject({
      protocol: 'graph', refreshToken: 'graph-rotated', canMarkRead: false,
      folders: [{ path: 'INBOX' }],
    })
    expect(fetcher.mock.calls.some(([url]) => String(url).includes('/me/?'))).toBe(false)
  })

  it('still checks address identity when User.Read was granted', async () => {
    const fetcher = remote(true, false)
    vi.stubGlobal('fetch', fetcher)
    await expect(validateMicrosoftImport(input)).rejects.toMatchObject({ code: 'graph_identity_mismatch' })
  })

  it('does not change the protocol of an existing IMAP account', async () => {
    const fetcher = remote()
    vi.stubGlobal('fetch', fetcher)
    await expect(validateMicrosoftImport(input, 'imap')).rejects.toBeInstanceOf(ImapConnectionError)
    expect(fetcher).toHaveBeenCalledTimes(1)
  })

  it('does not treat connection timeouts as protocol permission failures', async () => {
    imap.open.mockRejectedValue(new ImapConnectionError(504, 'Timed out'))
    const fetcher = remote()
    vi.stubGlobal('fetch', fetcher)
    await expect(validateMicrosoftImport(input)).rejects.toMatchObject({ status: 504 })
    expect(fetcher).toHaveBeenCalledTimes(1)
  })
})
