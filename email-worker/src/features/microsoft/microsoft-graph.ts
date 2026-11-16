import { parseMicrosoftMessage } from './microsoft-message-parser'
import { MicrosoftTokenError } from './microsoft-token'
import type { MicrosoftFolder, MicrosoftMessageMetadata } from './microsoft-types'

const GRAPH_ROOT = 'https://graph.microsoft.com/v1.0'
const MAX_MESSAGE_BYTES = 10 * 1024 * 1024
const MESSAGE_FIELDS = 'id,internetMessageId,from,toRecipients,ccRecipients,subject,bodyPreview,receivedDateTime,sentDateTime,isRead,flag,hasAttachments'
type Recipient = { emailAddress?: { name?: string; address?: string } }
type GraphFolder = { id: string; displayName: string; childFolderCount: number; totalItemCount: number }
type GraphMessage = {
  id: string; internetMessageId?: string; from?: Recipient; toRecipients?: Recipient[]
  ccRecipients?: Recipient[]; subject?: string; bodyPreview?: string; receivedDateTime?: string
  sentDateTime?: string; isRead?: boolean; flag?: { flagStatus?: string }; hasAttachments?: boolean
}

export class MicrosoftGraphError extends Error {
  constructor(readonly code: string, readonly status: number, readonly retryable = false,
    readonly retryAfter = 0) {
    super(`Microsoft Graph request failed (${code}).`)
  }
}

export function graphUrl(path: string): string {
  const url = new URL(path.startsWith('/') ? `${GRAPH_ROOT}${path}` : path)
  if (url.origin !== 'https://graph.microsoft.com' || !url.pathname.startsWith('/v1.0/me/')
    || !url.href.startsWith('https://graph.microsoft.com/') || url.hash) {
    throw new MicrosoftGraphError('invalid_graph_url', 400)
  }
  return url.href
}

function folderId(path: string): string {
  if (path.toUpperCase() === 'INBOX') return 'inbox'
  if (!path.startsWith('graph:') || path.length <= 6) {
    throw new MicrosoftGraphError('folder_not_found', 404)
  }
  return encodeURIComponent(path.slice(6))
}

function timestamp(value?: string): number {
  const parsed = Date.parse(value || '')
  return Number.isFinite(parsed) ? Math.floor(parsed / 1000) : 0
}

function addresses(recipients: Recipient[] = []): string[] {
  return recipients.map(({ emailAddress }) => emailAddress?.name
    ? `${emailAddress.name} <${emailAddress.address || ''}>` : emailAddress?.address || '')
    .filter(Boolean)
}

export function graphMetadata(message: GraphMessage): MicrosoftMessageMetadata {
  if (!message.id) throw new MicrosoftGraphError('invalid_graph_response', 502, true)
  const isRead = message.isRead === true
  const isStarred = message.flag?.flagStatus === 'flagged'
  return {
    uid: null, graphMessageId: message.id, internetMessageId: message.internetMessageId || '',
    senderName: message.from?.emailAddress?.name || '',
    senderAddress: message.from?.emailAddress?.address || '',
    recipients: addresses(message.toRecipients), cc: addresses(message.ccRecipients),
    subject: (message.subject || '').slice(0, 998), preview: (message.bodyPreview || '').slice(0, 500),
    receivedAt: timestamp(message.receivedDateTime), sentAt: timestamp(message.sentDateTime) || null,
    sizeBytes: 0, flags: [...(isRead ? ['\\Seen'] : []), ...(isStarred ? ['\\Flagged'] : [])],
    isRead, isStarred, hasAttachments: message.hasAttachments === true,
  }
}

export class MicrosoftGraphClient {
  readonly protocol = 'graph' as const
  constructor(private token: string, readonly canMarkRead: boolean,
    private readonly fetcher: typeof fetch = fetch,
    private readonly renew?: () => Promise<string>) {}

  async request(path: string, init: RequestInit = {}): Promise<Response> {
    const url = graphUrl(path)
    const send = () => this.fetcher(url, { ...init, redirect: 'error',
      headers: { Accept: 'application/json', ...init.headers,
        Authorization: `Bearer ${this.token}`, Prefer: 'IdType="ImmutableId"' },
      signal: AbortSignal.timeout(30_000) })
    let response: Response
    try {
      response = await send()
      if (response.status === 401 && this.renew) {
        await response.body?.cancel()
        this.token = await this.renew()
        response = await send()
      }
    } catch (error) {
      if (error instanceof MicrosoftTokenError || (error instanceof Error && 'code' in error)) throw error
      throw new MicrosoftGraphError('graph_unavailable', 503, true)
    }
    if (!response.ok) {
      await response.body?.cancel()
      const code = response.status === 401 ? 'graph_auth_rejected'
        : response.status === 403 ? 'graph_access_rejected'
          : response.status === 404 ? 'remote_message_not_found'
            : response.status === 429 ? 'graph_throttled' : 'graph_unavailable'
      const retryAfter = Number(response.headers.get('Retry-After'))
      throw new MicrosoftGraphError(code, response.status,
        response.status === 429 || response.status >= 500,
        Number.isFinite(retryAfter) ? Math.max(0, retryAfter) : 0)
    }
    return response
  }

  async json<T>(path: string): Promise<T> {
    return (await this.request(path)).json<T>()
  }

  // A mailbox read probes the actual granted capability; User.Read is not required.
  async open(): Promise<void> { await this.examineFolder('INBOX') }
  async close(): Promise<void> { /* HTTP responses are consumed by each operation. */ }

  async validateIdentity(email: string): Promise<void> {
    const profile = await this.json<{ mail?: string; userPrincipalName?: string }>(
      '/me/?$select=mail,userPrincipalName',
    )
    if (![profile.mail, profile.userPrincipalName].some((value) => value?.toLowerCase() === email.toLowerCase())) {
      throw new MicrosoftGraphError('graph_identity_mismatch', 400)
    }
  }

  async examineFolder(path: string): Promise<{ uidValidity: null; exists: number }> {
    const folder = await this.json<GraphFolder>(`/me/mailFolders/${folderId(path)}?$select=id,totalItemCount`)
    return { uidValidity: null, exists: folder.totalItemCount }
  }

  async listFolders(): Promise<MicrosoftFolder[]> {
    const inbox = await this.json<GraphFolder>('/me/mailFolders/inbox?$select=id,displayName,totalItemCount')
    const folders: MicrosoftFolder[] = [{ path: 'INBOX', displayName: inbox.displayName || 'INBOX',
      flags: ['\\Inbox'], specialUse: '\\Inbox', uidValidity: null, lastUid: 0 }]
    const seen = new Set([inbox.id])
    const pending = ['/me/mailFolders?$select=id,displayName,childFolderCount&$top=100']
    let requests = 0
    while (pending.length) {
      if (++requests > 100 || folders.length > 1000) {
        throw new MicrosoftGraphError('response_too_large', 413)
      }
      const page = await this.json<{ value: GraphFolder[]; '@odata.nextLink'?: string }>(pending.shift()!)
      for (const folder of page.value) {
        // Inbox's children must also be visited even though INBOX is already in the list.
        if (folder.childFolderCount > 0 && (!seen.has(folder.id) || folder.id === inbox.id)) {
          pending.push(`/me/mailFolders/${encodeURIComponent(folder.id)}/childFolders?$select=id,displayName,childFolderCount&$top=100`)
        }
        if (seen.has(folder.id)) continue
        seen.add(folder.id)
        folders.push({ path: `graph:${folder.id}`, displayName: folder.displayName,
          flags: [], specialUse: '', uidValidity: null, lastUid: 0 })
      }
      if (page['@odata.nextLink']) pending.push(graphUrl(page['@odata.nextLink']))
    }
    return folders
  }

  async listMessages(path: string, limit: number): Promise<MicrosoftMessageMetadata[]> {
    const maximum = Math.min(500, Math.max(1, limit))
    const messages: MicrosoftMessageMetadata[] = []
    const seen = new Set<string>()
    const query = new URLSearchParams({ '$select': MESSAGE_FIELDS,
      '$orderby': 'receivedDateTime desc', '$top': String(Math.min(100, maximum)) })
    let next = `/me/mailFolders/${folderId(path)}/messages?${query}`
    const visited = new Set<string>()
    while (next && messages.length < maximum) {
      if (visited.has(next) || visited.size >= 50) throw new MicrosoftGraphError('invalid_graph_response', 502)
      visited.add(next)
      const page = await this.json<{ value: GraphMessage[]; '@odata.nextLink'?: string }>(next)
      for (const value of page.value) {
        if (seen.has(value.id)) continue
        seen.add(value.id)
        messages.push(graphMetadata(value))
        if (messages.length === maximum) break
      }
      next = page['@odata.nextLink'] ? graphUrl(page['@odata.nextLink']) : ''
    }
    return messages
  }

  async getMessage(path: string, messageId: string) {
    const response = await this.request(`/me/mailFolders/${folderId(path)}/messages/${encodeURIComponent(messageId)}/$value`,
      { headers: { Accept: 'message/rfc822' } })
    const reader = response.body?.getReader()
    if (!reader) throw new MicrosoftGraphError('invalid_graph_response', 502)
    const chunks: Uint8Array[] = []
    let length = 0
    try {
      while (true) {
        const { done, value } = await reader.read()
        if (done) break
        length += value.byteLength
        if (length > MAX_MESSAGE_BYTES) throw new MicrosoftGraphError('response_too_large', 413)
        chunks.push(value)
      }
    } finally { await reader.cancel() }
    const bytes = new Uint8Array(length)
    let offset = 0
    for (const chunk of chunks) { bytes.set(chunk, offset); offset += chunk.byteLength }
    return parseMicrosoftMessage(bytes, messageId)
  }

  async markSeen(path: string, messageId: string): Promise<void> {
    if (!this.canMarkRead) throw new MicrosoftGraphError('graph_readonly', 403)
    const response = await this.request(`/me/mailFolders/${folderId(path)}/messages/${encodeURIComponent(messageId)}`, {
      method: 'PATCH', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ isRead: true }),
    })
    await response.body?.cancel()
  }
}
