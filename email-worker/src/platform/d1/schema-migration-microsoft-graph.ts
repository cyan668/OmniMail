export const MICROSOFT_GRAPH_MIGRATION = '0042_microsoft_graph.sql'

export const MICROSOFT_GRAPH_RECOVERY = {
  name: MICROSOFT_GRAPH_MIGRATION,
  statements: [
    `ALTER TABLE microsoft_imap_accounts ADD COLUMN protocol TEXT NOT NULL DEFAULT 'imap'
  CHECK (protocol IN ('imap', 'graph'))`,
    `ALTER TABLE microsoft_imap_accounts ADD COLUMN can_mark_read INTEGER NOT NULL DEFAULT 1
  CHECK (can_mark_read IN (0, 1))`,
    `CREATE TABLE microsoft_messages_v2 (
  id TEXT PRIMARY KEY,
  account_id TEXT NOT NULL REFERENCES microsoft_imap_accounts(id) ON DELETE CASCADE,
  folder_path TEXT NOT NULL,
  uid_validity INTEGER,
  imap_uid INTEGER,
  internet_message_id TEXT NOT NULL DEFAULT '',
  sender_name TEXT NOT NULL DEFAULT '',
  sender_address TEXT NOT NULL DEFAULT '',
  recipients_json TEXT NOT NULL DEFAULT '[]',
  cc_json TEXT NOT NULL DEFAULT '[]',
  subject TEXT NOT NULL DEFAULT '',
  preview TEXT NOT NULL DEFAULT '',
  received_at INTEGER NOT NULL,
  sent_at INTEGER,
  size_bytes INTEGER NOT NULL DEFAULT 0 CHECK (size_bytes >= 0),
  flags_json TEXT NOT NULL DEFAULT '[]',
  is_read INTEGER NOT NULL DEFAULT 0 CHECK (is_read IN (0, 1)),
  is_starred INTEGER NOT NULL DEFAULT 0 CHECK (is_starred IN (0, 1)),
  has_attachments INTEGER NOT NULL DEFAULT 0 CHECK (has_attachments IN (0, 1)),
  created_at INTEGER NOT NULL,
  updated_at INTEGER NOT NULL,
  graph_message_id TEXT NOT NULL DEFAULT '',
  UNIQUE (account_id, folder_path, uid_validity, imap_uid),
  FOREIGN KEY (account_id, folder_path)
    REFERENCES microsoft_imap_folders(account_id, path) ON DELETE CASCADE,
  CHECK (
    (graph_message_id = '' AND uid_validity IS NOT NULL AND uid_validity > 0
      AND imap_uid IS NOT NULL AND imap_uid > 0)
    OR (graph_message_id != '' AND uid_validity IS NULL AND imap_uid IS NULL)
  )
)`,
    `INSERT INTO microsoft_messages_v2 SELECT *, '' FROM microsoft_imap_messages`,
    `DROP TABLE microsoft_imap_messages`,
    `ALTER TABLE microsoft_messages_v2 RENAME TO microsoft_imap_messages`,
    `CREATE UNIQUE INDEX idx_microsoft_graph_message
  ON microsoft_imap_messages(account_id, folder_path, graph_message_id)
  WHERE graph_message_id != ''`,
    `CREATE INDEX idx_microsoft_imap_messages_folder_date
  ON microsoft_imap_messages(account_id, folder_path, received_at DESC, id DESC)`,
    `CREATE INDEX idx_microsoft_imap_messages_date
  ON microsoft_imap_messages(received_at DESC, id DESC, account_id)`,
    `CREATE TRIGGER trg_microsoft_notification_insert AFTER INSERT ON microsoft_imap_messages BEGIN
  INSERT INTO mail_notification_versions (user_id,source,version) SELECT user_id, 'microsoft', 1 FROM microsoft_imap_accounts WHERE id = NEW.account_id AND EXISTS (SELECT 1 FROM users WHERE users.id=microsoft_imap_accounts.user_id) ON CONFLICT(user_id,source) DO UPDATE SET version=version+1;
END`,
    `CREATE TRIGGER trg_microsoft_notification_delete AFTER DELETE ON microsoft_imap_messages BEGIN
  INSERT INTO mail_notification_versions (user_id,source,version) SELECT user_id, 'microsoft', 1 FROM microsoft_imap_accounts WHERE id = OLD.account_id AND EXISTS (SELECT 1 FROM users WHERE users.id=microsoft_imap_accounts.user_id) ON CONFLICT(user_id,source) DO UPDATE SET version=version+1;
END`,
    `CREATE TRIGGER trg_microsoft_notification_update AFTER UPDATE OF account_id, sender_name, sender_address, subject, is_read, received_at, id, folder_path ON microsoft_imap_messages
WHEN OLD.account_id IS NOT NEW.account_id OR OLD.sender_name IS NOT NEW.sender_name OR OLD.sender_address IS NOT NEW.sender_address OR OLD.subject IS NOT NEW.subject OR OLD.is_read IS NOT NEW.is_read OR OLD.received_at IS NOT NEW.received_at OR OLD.id IS NOT NEW.id OR OLD.folder_path IS NOT NEW.folder_path
BEGIN
  INSERT INTO mail_notification_versions (user_id,source,version) SELECT DISTINCT user_id, 'microsoft', 1 FROM microsoft_imap_accounts WHERE id IN (OLD.account_id,NEW.account_id) AND EXISTS (SELECT 1 FROM users WHERE users.id=microsoft_imap_accounts.user_id) ON CONFLICT(user_id,source) DO UPDATE SET version=version+1;
END`,
    `CREATE TRIGGER trg_telegram_microsoft_insert
AFTER INSERT ON microsoft_imap_messages
WHEN NEW.is_read = 0 AND NEW.folder_path = 'INBOX' COLLATE NOCASE
BEGIN
  INSERT OR IGNORE INTO notification_outbox
    (id, endpoint_id, source, account_id, message_id, next_attempt_at, created_at)
  SELECT lower(hex(randomblob(16))), e.id, 'microsoft', NEW.account_id, NEW.id, unixepoch(), unixepoch()
  FROM microsoft_imap_accounts a JOIN notification_endpoints e ON e.user_id = a.user_id
  JOIN microsoft_imap_folders f ON f.account_id = NEW.account_id AND f.path = NEW.folder_path
  WHERE a.id = NEW.account_id AND a.last_synced_at IS NOT NULL
    AND ((a.protocol = 'imap' AND f.uid_validity = NEW.uid_validity AND NEW.imap_uid > f.last_uid)
      OR (a.protocol = 'graph' AND NEW.graph_message_id != '' AND NEW.received_at >= a.last_synced_at))
    AND e.enabled = 1 AND e.status = 'active' AND NEW.created_at >= e.enabled_at
    AND EXISTS (SELECT 1 FROM json_each(e.sources_json) WHERE value = 'microsoft');
END`,
  ],
} as const
