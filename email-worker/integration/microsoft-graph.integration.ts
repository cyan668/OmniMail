import { env } from 'cloudflare:workers'
import { applyD1Migrations } from 'cloudflare:test'
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest'
import type { SessionUser } from '../src/app/types'
import { importMicrosoftAccounts, listMicrosoftAccounts, verifyMicrosoftAccount } from '../src/features/microsoft/microsoft-account-api'
import { getMicrosoftAttachment, getMicrosoftMessage, listMicrosoftMessages } from '../src/features/microsoft/microsoft-message-api'
import { syncMicrosoftAccount, messageStatement } from '../src/features/microsoft/microsoft-sync'
import { graphMetadata } from '../src/features/microsoft/microsoft-graph'
import { microsoftAccountForSync } from '../src/features/microsoft/microsoft-store'
import { microsoftAccessToken } from '../src/features/microsoft/microsoft-token-manager'

const user = { id: 'graph-owner', email: 'graph@example.com', role: 'super_admin' } as SessionUser
const now = Math.floor(Date.now() / 1000)
const mime = ['From: Sender <sender@example.com>', 'To: graph@example.com', 'Subject: Graph test',
  'MIME-Version: 1.0', 'Content-Type: multipart/mixed; boundary="test"', '',
  '--test', 'Content-Type: text/plain', '', 'Graph body', '--test', 'Content-Type: text/plain',
  'Content-Disposition: attachment; filename="hello.txt"', 'Content-Transfer-Encoding: base64', '',
  'aGVsbG8=', '--test--', ''].join('\r\n')

beforeAll(async () => {
  await applyD1Migrations(env.DB, env.TEST_MIGRATIONS.filter(({ name }) => name < '0042'))
  await env.DB.prepare("INSERT INTO users(id,email,display_name,password_hash) VALUES('graph-owner','graph@example.com','Graph','test')").run()
  await env.DB.prepare(`INSERT INTO microsoft_imap_accounts(id,user_id,name,provided_email,normalized_email,auth_mode,client_id,refresh_token_cipher,created_at,updated_at)
    VALUES('legacy-ms','graph-owner','Legacy','legacy@example.com','legacy@example.com','oauth2','test','test',1,1)`).run()
  await env.DB.prepare("INSERT INTO microsoft_imap_folders(account_id,path,display_name,last_listed_at) VALUES('legacy-ms','INBOX','INBOX',1)").run()
  await env.DB.prepare(`INSERT INTO microsoft_imap_messages(id,account_id,folder_path,uid_validity,imap_uid,subject,received_at,created_at,updated_at)
    VALUES('legacy-message','legacy-ms','INBOX',42,7,'Keep me',1,1,1)`).run()
  await applyD1Migrations(env.DB, env.TEST_MIGRATIONS)
})
afterEach(() => vi.unstubAllGlobals())

describe('Microsoft Graph with real D1 and existing IMAP rows', () => {
  it('preserves legacy identity, triggers, and foreign keys when migrating', async () => {
    expect(await env.DB.prepare("SELECT uid_validity,imap_uid,subject,graph_message_id FROM microsoft_imap_messages WHERE id='legacy-message'").first())
      .toEqual({ uid_validity: 42, imap_uid: 7, subject: 'Keep me', graph_message_id: '' })
    expect(await env.DB.prepare("SELECT protocol FROM microsoft_imap_accounts WHERE id='legacy-ms'").first()).toEqual({ protocol: 'imap' })
    const triggers = await env.DB.prepare("SELECT name FROM sqlite_master WHERE type='trigger' AND tbl_name='microsoft_imap_messages'").all<{ name: string }>()
    expect(triggers.results.map(({ name }) => name).sort()).toEqual([
      'trg_microsoft_notification_delete', 'trg_microsoft_notification_insert',
      'trg_microsoft_notification_update', 'trg_telegram_microsoft_insert',
    ])
    expect((await env.DB.prepare('PRAGMA foreign_key_check').all()).results).toEqual([])
    const accountTriggers = await env.DB.prepare("SELECT name FROM sqlite_master WHERE type='trigger' AND tbl_name='microsoft_imap_accounts'").all<{ name: string }>()
    expect(accountTriggers.results.map(({ name }) => name).sort()).toEqual([
      'trg_microsoft_notification_account_delete', 'trg_microsoft_notification_account_insert',
      'trg_microsoft_notification_account_update',
    ])
  })

  it('imports Graph-only authorization, syncs without duplicates, reads MIME/attachment and rotates tokens', async () => {
    let marked = false
    const fetcher = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = String(input)
      if (url.startsWith('https://login.microsoftonline.com/')) {
        const form = new URLSearchParams(String(init?.body))
        if (form.get('scope')?.includes('IMAP')) return Response.json({ error: 'invalid_grant' }, { status: 400 })
        return Response.json({ access_token: 'graph-access', refresh_token: 'graph-rotated', expires_in: 3600,
          scope: 'https://graph.microsoft.com/Mail.ReadWrite https://graph.microsoft.com/User.Read' })
      }
      expect(new Headers(init?.headers).get('Prefer')).toBe('IdType="ImmutableId"')
      if (url.includes('/me/?')) return Response.json({ mail: 'graph@example.com' })
      if (url.includes('/$value')) return new Response(mime)
      if (init?.method === 'PATCH') { marked = true; expect(JSON.parse(String(init.body))).toEqual({ isRead: true }); return Response.json({ isRead: true }) }
      if (url.includes('/messages?')) return Response.json({ value: [{ id: 'immutable=message', subject: 'Graph test',
        receivedDateTime: new Date(now * 1000).toISOString(), isRead: marked, hasAttachments: true }] })
      if (url.includes('/inbox?')) return Response.json({ id: 'inbox-id', displayName: 'Inbox', totalItemCount: 1 })
      return Response.json({ value: [{ id: 'inbox-id', displayName: 'Inbox', childFolderCount: 0 }] })
    })
    vi.stubGlobal('fetch', fetcher)
    const runtime = { ...env, MAIL_QUEUE: { send: vi.fn(async () => undefined) } as unknown as Queue }
    const response = await importMicrosoftAccounts(runtime, user, new Request('https://mail.example.com/api/microsoft/accounts/import', {
      method: 'POST', body: JSON.stringify({ accounts: [{ email: user.email, authMode: 'oauth2', refreshToken: 'graph-grant',
        clientId: '00000000-0000-4000-8000-000000000000' }] }),
    }), '127.0.0.1')
    expect(response.status).toBe(201)
    const imported = await response.json<{ results: Array<{ account: { id: string; protocol: string } }> }>()
    const accountId = imported.results[0].account.id
    expect(imported.results[0].account.protocol).toBe('graph')
    expect(await syncMicrosoftAccount(runtime, accountId, now)).toMatchObject({ status: 'synced' })
    expect(await syncMicrosoftAccount(runtime, accountId, now + 1)).toMatchObject({ status: 'synced' })
    const rows = await env.DB.prepare('SELECT id,uid_validity,imap_uid,graph_message_id FROM microsoft_imap_messages WHERE account_id=?')
      .bind(accountId).all<{ id: string; uid_validity: number | null; imap_uid: number | null; graph_message_id: string }>()
    expect(rows.results).toHaveLength(1)
    expect(rows.results[0]).toMatchObject({ uid_validity: null, imap_uid: null, graph_message_id: 'immutable=message' })
    const messageId = rows.results[0].id
    const listed = await listMicrosoftMessages(runtime, user, new Request(`https://mail.example.com/api/microsoft/messages?accountId=${accountId}`))
    expect(await listed.json()).toMatchObject({ messages: [{ account: { protocol: 'graph' }, subject: 'Graph test' }] })
    expect((await verifyMicrosoftAccount(runtime, user, accountId, '127.0.0.1')).status).toBe(200)
    const detail = await getMicrosoftMessage(runtime, user, accountId, messageId)
    expect(await detail.json()).toMatchObject({ message: { body: 'Graph body', isRead: true, attachments: [{ filename: 'hello.txt', partId: '0' }] } })
    const attachment = await getMicrosoftAttachment(runtime, user, accountId, messageId, '0')
    expect(await attachment.text()).toBe('hello')
    const otherUser = { ...user, id: 'another-user' }
    expect((await getMicrosoftMessage(runtime, otherUser, accountId, messageId)).status).toBe(404)
    const account = (await microsoftAccountForSync(runtime, accountId))!
    account.accessTokenExpiresAt = 0
    await microsoftAccessToken(runtime, account, { now: now + 2, fetcher })
    expect(account).toMatchObject({ protocol: 'graph', refreshToken: 'graph-rotated', canMarkRead: true })
    expect(JSON.stringify(await (await listMicrosoftAccounts(runtime, user)).json())).not.toContain('graph-rotated')
    const version = await env.DB.prepare("SELECT version FROM mail_notification_versions WHERE user_id=? AND source='microsoft'").bind(user.id).first<{ version: number }>()
    expect(version!.version).toBeGreaterThan(1)
    await env.DB.prepare(`INSERT INTO notification_endpoints(id,user_id,chat_id,sources_json,enabled_at,created_at,updated_at)
      VALUES('graph-endpoint',?,'123456789','["microsoft"]',?,?,?)`).bind(user.id, now - 10, now - 10, now - 10).run()
    const next = graphMetadata({ id: 'new-message', receivedDateTime: new Date((now + 5) * 1000).toISOString(), isRead: false })
    await messageStatement(runtime, accountId, 'INBOX', null, next, now + 5).run()
    await messageStatement(runtime, accountId, 'INBOX', null, next, now + 6).run()
    expect(await env.DB.prepare("SELECT COUNT(*) AS n FROM notification_outbox WHERE endpoint_id='graph-endpoint'").first()).toEqual({ n: 1 })
  })
})
