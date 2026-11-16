import type { Env } from '../../app/types'
import type { MicrosoftGraphClient } from './microsoft-graph'
import { MicrosoftStoreError } from './microsoft-store'
import { messageStatement } from './microsoft-sync'

export async function refreshMicrosoftGraphFolder(
  env: Env, accountId: string, folderPath: string, limit: number,
  client: MicrosoftGraphClient, now: number,
): Promise<{ uidValidity: null; indexed: number }> {
  const folder = await env.DB.prepare(
    'SELECT path FROM microsoft_imap_folders WHERE account_id = ? AND path = ? LIMIT 1',
  ).bind(accountId, folderPath).first()
  if (!folder) throw new MicrosoftStoreError(404, 'folder_not_found', 'Microsoft 文件夹不存在。')
  const { results: existing } = await env.DB.prepare(
    `SELECT graph_message_id FROM microsoft_imap_messages
     WHERE account_id = ? AND folder_path = ? AND graph_message_id != ''`,
  ).bind(accountId, folderPath).all<{ graph_message_id: string }>()
  const messages = await client.listMessages(folderPath, Math.min(500, Math.max(limit, existing.length)))
  const remoteIds = new Set(messages.map((message) => message.graphMessageId))
  const statements = messages.map((message) => messageStatement(env, accountId, folderPath, null, message, now))
  // The index is a bounded latest-message window. Do not delete anything until all pages succeeded.
  for (const row of existing) {
    if (!remoteIds.has(row.graph_message_id)) statements.push(env.DB.prepare(
      `DELETE FROM microsoft_imap_messages
       WHERE account_id = ? AND folder_path = ? AND graph_message_id = ?`,
    ).bind(accountId, folderPath, row.graph_message_id))
  }
  statements.push(env.DB.prepare(
    'UPDATE microsoft_imap_folders SET last_listed_at = ? WHERE account_id = ? AND path = ?',
  ).bind(now, accountId, folderPath))
  await env.DB.batch(statements)
  return { uidValidity: null, indexed: messages.length }
}
