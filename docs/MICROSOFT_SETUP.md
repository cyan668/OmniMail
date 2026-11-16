# Microsoft 邮箱设置指南

> 新部署推荐使用统一 `MAIL_CREDENTIALS_KEY`；旧 `MICROSOFT_CREDENTIALS_KEY` 继续兼容。升级时先保留旧值，按[迁移指南](MAIL_CREDENTIALS.md)由主管理员主动迁移。

OmniMail 支持 IMAP OAuth2 和 Microsoft Graph 两种收信方式，使用用户提供的 refresh token
与 Client ID。导入时先验证 IMAP；IMAP 授权范围不匹配时尝试 Graph 的已授权权限，并保存
实际通过邮箱验证的连接方式。现有 IMAP 账号继续使用原连接方式。

首期支持 Azure Global 上的 Outlook.com、Hotmail、Live，以及租户管理员允许委托邮箱访问的
Microsoft 365 委托式账号。不支持世纪互联、中国区、GCC High 或 DoD 端点。

## 1. 部署配置

1. 在 Worker 的 **Variables & Secrets** 中新增 Secret `MAIL_CREDENTIALS_KEY`。
   值必须至少包含 32 个随机 UTF-8 字节，并在迁移或恢复部署时保持不变。
2. 可选新增 Text 变量 `MICROSOFT_MAIL_ENABLED=true`。设为 `false` 会隐藏入口并停止定时入队，
   但不会删除已保存的账号、密文或索引。
3. 使用 `npm run deploy` 应用 D1 迁移，包括 `0042_microsoft_graph.sql`，然后部署 Worker。
   该迁移保留旧 IMAP 索引、凭据和通知触发器，并增加 Graph 邮件身份字段。
4. 确认 `MAIL_QUEUE` producer/consumer 与 `*/5 * * * *` Cron 已按 `wrangler.jsonc` 绑定。
5. 主管理员可在 **系统设置 → 邮箱功能入口** 中隐藏或恢复 Microsoft 入口。

本地开发可复制示例变量：

```text
MAIL_CREDENTIALS_KEY=replace-with-at-least-32-random-bytes
MICROSOFT_MAIL_ENABLED=true
```

不要把真实密钥、refresh token、access token 或密码提交到 Git。

## 2. OAuth2 准备

导入 OAuth2 账号需要同一应用和用户配套的：

- 邮箱地址；
- refresh token；
- Client ID（UUID）；
- authority：`common`、`consumers`、`organizations` 或具体 tenant UUID。

签发 refresh token 时，IMAP 账号需要委托式 Outlook IMAP 权限
`https://outlook.office.com/IMAP.AccessAsUser.All` 和 `offline_access`。
Graph 账号需要 `Mail.Read` 或 `Mail.ReadWrite`、`User.Read`（核对邮箱身份）与 `offline_access`。
只有 `Mail.Read` 时仍可收信、打开正文和附件，但不会修改远端已读状态。OmniMail 不内置或借用
任何第三方 Client ID，也不负责绕过租户同意、条件访问或管理员策略。授权与 refresh token 行为见
[Microsoft IMAP OAuth 文档](https://learn.microsoft.com/en-us/exchange/client-developer/legacy-protocols/how-to-authenticate-an-imap-pop-smtp-application-by-using-oauth)和
[Microsoft OAuth 授权码流程](https://learn.microsoft.com/en-us/entra/identity-platform/v2-oauth2-auth-code-flow)。

IMAP 账号的令牌刷新请求为：

```text
POST https://login.microsoftonline.com/{authority}/oauth2/v2.0/token
scope=https://outlook.office.com/IMAP.AccessAsUser.All offline_access
```

Graph 账号使用同一 token endpoint，scope 为
`https://graph.microsoft.com/.default offline_access`，只取得应用已获同意的 Graph 权限。
填写邮箱须与 Graph `/me` 返回的主邮箱或 userPrincipalName 一致。

成功兑换后才会验证 Microsoft 邮箱。若 Microsoft 返回轮换后的 refresh token，OmniMail 会在
账号级刷新租约保护下原子替换旧密文。XOAUTH2 失败时只允许强制刷新并重试一次，不会改用密码。

## 3. 导入格式

Microsoft 工作区提供 **分字段录入** 和 **批量导入** 两种入口。分字段录入适合单个账号，或
password 本身包含字面量 `----` 的组合凭据；批量导入每行使用以下一种格式：

```text
email----password----refresh_token----client_id
email----password----client_id----refresh_token
email--------refresh_token----client_id
```

- 四字段格式的最后两段可以互换；浏览器根据其中唯一的 UUID 自动识别 Client ID。两段都不是
  UUID 或两段都是 UUID 时会拒绝该行，避免猜测。
- 四字段中只要 refresh token 与 Client ID 齐全，就使用 OAuth2。用户确认后，组合里的 password
  会用独立 AAD 上下文执行 AES-GCM 加密，保存到 `combination_password_cipher`；它不会进入
  IMAP LOGIN，也不会作为 OAuth2 失败后的认证回退。
- 两字段 `email----password` 不再支持；前端与结构化 API 都会拒绝。
- 八个连续连字符表示 OAuth2 四字段格式中的 password 为空。
- 预览只显示规范化邮箱、识别模式和脱敏 Client ID，不显示密码或完整 token。
- 批量导入分为“输入账号”和“安全确认”两步；输入框高度固定，安全预览出现前不会上传凭据。
- 每批最多 25 个有效账号；浏览器逐项提交并显示真实验证进度。当前项会显示转圈，成功后显示
  对勾并从列表移出；失败项保留具体错误。用户/IP 验证窗口允许连续处理两批完整批次，超过后
  才限速。
- 若密码本身包含字面量 `----`，请改用 **分字段录入**，不要让解析器猜测字段位置。

## 4. 验证与收信

连接时 Worker 会先完成真实 token 兑换。IMAP 验证 XOAUTH2、`LIST` 和 INBOX `EXAMINE`；
Graph 核对账号身份、INBOX 访问及文件夹列表。验证成功后才保存账号。

连接成功后：

- 后台约每 5 分钟同步一次 INBOX 元数据，每账号最多保留最近 500 条；
- 工作区可聚合全部账号 INBOX，或选择单账号与服务器返回的文件夹；
- 单次读取数量可选 25、50、100 或 200，服务端仍强制限制为 1–200；
- “全部 Microsoft”范围的同步按钮会为所有已连接账号逐个请求 INBOX Queue 同步；
- 单账号范围的“远程刷新”通过当前连接方式只读更新文件夹索引；Graph 使用 immutable ID，
  按完整 nextLink 分页读取最近邮件，非 INBOX 文件夹使用服务端返回的 `graph:` 路径；
- 正文和附件按需读取 MIME：IMAP 使用 `BODY.PEEK[]`，Graph 使用 `messages/{id}/$value`。
  正文读取成功后，有写权限的未读邮件同步已读：IMAP 写 `\Seen`，Graph 写 `isRead:true`；
- 已读写入失败不会阻断正文显示，重新打开可重试；移动、删除、归档、星标和其他 flags 写入均未开放；
- 删除连接只清理 OmniMail 本地密文与索引，不删除远端邮件。OAuth2 用户还应在 Microsoft
  账户或租户应用授权页面撤销不再使用的授权。

这是轮询式定时收信，不是 IMAP IDLE 或秒级推送。

## 5. 工作区操作

- 范围选择器默认是“全部 Microsoft”，聚合所有账号的 INBOX；选中账号后才显示其文件夹。
- 顶部复制按钮在全部范围复制第一个账号邮箱，单账号范围复制当前邮箱；范围面板中每个账号也有
  独立复制按钮。
- 全部范围使用“同步全部 Microsoft 账号”；单账号范围使用“远程刷新当前文件夹”。
- 账号管理支持修改备注、验证、同步、替换 OAuth2 凭据、单个断开和批量断开。
- 批量断开会逐个删除 OmniMail 本地记录，并显示当前进度；操作前需要二次确认。
- 连接、批量导入和账号管理共用稳定高度的响应式弹窗；内容过长时只在弹窗内部滚动。

## 6. 真实账号上线验收

代码测试使用受控协议响应，不包含任何真实凭据。正式启用前，请用一个专用 Outlook.com 测试
账号在部署后的工作区完成以下探针；整个过程不要截图、记录或复制 token 到日志：

1. OAuth2 导入成功，账号状态为“已连接”；
2. 文件夹列表可刷新，INBOX 可读取；
3. 能打开一封未读的纯文本或 HTML 邮件，并确认 Microsoft 端已标记为已读；
4. 能下载一个不超过 5 MiB 的测试附件；
5. 在全部范围触发同步，确认每个账号均进入 Queue；切换单账号后确认当前文件夹可远程刷新；
6. 顶部复制按钮和范围面板复制按钮只复制预期邮箱，不泄露其他凭据；
7. 手动同步入队后更新时间变化，下一次 Cron 同步不产生重复记录；
8. 撤销应用授权后，账号进入凭据或权限错误且不会无限重试；
9. 如需宣称 Microsoft 365 支持，再用受控工作/学校账号重复以上步骤，并确认租户允许 IMAP。

OmniMail 不使用 ROPC、密码 LOGIN、网页登录自动化、代理或其他规避措施。

## 7. 安全边界与故障排查

- OAuth 主机、authority 形式、IMAP 主机和端口、Graph `/v1.0/me/` URL 均由服务端白名单固定，导入数据不能指定 URL、
  主机或端口。
- 账号、文件夹、邮件和附件查询都同时校验当前用户归属；消息远端身份绑定
  IMAP 的 `folder + UIDVALIDITY + UID` 或 Graph 的 `folder + immutable ID`。
- 唯一允许的远端写入是选中已校验文件夹后，按 IMAP UID 或 Graph immutable ID 标记已读；不接受客户端传入
  IMAP 命令、flags、主机或端口。
- API 响应使用 `private, no-store`；审计只记录脱敏邮箱和认证模式。
- `invalid_grant` 通常表示 refresh token 失效或被撤销；请重新授权并替换凭据。
- `imap_scope_missing` 表示 token 不含 Outlook IMAP 委托 scope。
- `graph_scope_missing` 表示缺少 Graph Mail.Read/Mail.ReadWrite；`graph_access_rejected`
  表示 Microsoft 拒绝资源访问；`graph_identity_mismatch` 表示邮箱地址和授权身份不匹配。
- 部分 Graph-only token 请求 IMAP scope 时会返回 `invalid_grant`；新连接会继续验证 Graph，
  不能仅凭 IMAP 的这条错误判断邮箱或 refresh token 已失效。
- `imap_access_rejected` 或 `permission_error` 可能表示租户关闭 IMAP、缺少同意或条件访问阻止。
- `credential_decryption_failed` 表示部署密钥与保存凭据时不一致；恢复原
  `MAIL_CREDENTIALS_KEY`（旧密文对应 `MICROSOFT_CREDENTIALS_KEY`），或断开后重新连接账号。

完整端点与响应说明见 [Microsoft API 参考](api/microsoft.md)。
