# Microsoft Graph API v1.0 接口覆盖率检查

本文档记录了 msgraph-orm-js SDK 对 Microsoft Graph REST API v1.0 的接口覆盖情况。

## 覆盖率统计

- ✅ **已完整实现**: 95+个主要 API 端点
- ⚠️ **部分实现**: 5个 API 端点
- ❌ **未实现**: 少数高级/特殊场景 API

## 详细覆盖情况

### 1. 用户和身份管理 (Users & Identity) ✅

| API 端点 | 状态 | ORM 方法 | 权限 |
|---------|------|----------|------|
| /users | ✅ | `orm.users` | User.Read.All |
| /users/{id} | ✅ | `orm.users.findById()` | User.Read.All |
| /groups | ✅ | `orm.groups` | Group.Read.All |
| /groups/{id}/members | ✅ | `orm.groupMembers()` | GroupMember.Read.All |
| /groups/{id}/owners | ✅ | `orm.groupOwners()` | Group.Read.All |
| /devices | ✅ | `orm.devices` | Device.Read.All |
| /applications | ✅ | `orm.applications` | Application.Read.All |
| /servicePrincipals | ✅ | `orm.servicePrincipals` | Application.Read.All |
| /directoryRoles | ✅ | `orm.directoryRoles` | RoleManagement.Read.Directory |
| /directoryRoleTemplates | ✅ | `orm.directoryRoleTemplates` | RoleManagement.Read.Directory |
| /directoryRoles/{id}/members | ✅ | `orm.directoryRoleMembers()` | RoleManagement.Read.Directory |
| /directory/administrativeUnits | ✅ | `orm.administrativeUnits` | AdministrativeUnit.Read.All |
| /users/{id}/licenseDetails | ✅ | `orm.getUserLicenses()` | User.Read.All |
| /subscribedSkus | ✅ | `orm.getSubscribedSkus()` | Organization.Read.All |
| /users/{id}/assignLicense | ✅ | `orm.assignUserLicense()` | User.ReadWrite.All |
| /invitations | ✅ | `orm.createInvitation()` | User.Invite.All |
| /organization | ✅ | `orm.organization()` | Organization.Read.All |
| /me | ✅ | `orm.me()` | User.Read |
| /directoryObjects/{id} | ✅ | `orm.getDirectoryObject()` | Directory.Read.All |
| /directoryObjects/getByIds | ✅ | `orm.getDirectoryObjectsByIds()` | Directory.Read.All |
| /directoryObjects/{id}/checkMemberGroups | ✅ | `orm.checkMemberGroups()` | Directory.Read.All |
| /directoryObjects/{id}/getMemberObjects | ✅ | `orm.getMemberObjects()` | Directory.Read.All |

### 2. 日历 (Calendar) ✅

| API 端点 | 状态 | ORM 方法 | 权限 |
|---------|------|----------|------|
| /users/{id}/events | ✅ | `orm.events()` | Calendars.Read |
| /users/{id}/calendar/events | ✅ | `orm.calendarEvents()` | Calendars.Read |
| /users/{id}/calendars/{id}/events | ✅ | `orm.calendarEventsById()` | Calendars.Read |
| /groups/{id}/events | ✅ | `orm.groupEvents()` | Group.Read.All |
| /users/{id}/calendars | ✅ | `orm.calendars()` | Calendars.Read |
| /users/{id}/calendar | ✅ | `orm.getUserCalendar()` | Calendars.Read |
| /users/{id}/calendarGroups | ✅ | `orm.calendarGroups()` | Calendars.Read |
| /users/{id}/calendarGroups/{id}/calendars | ✅ | `orm.calendarsInGroup()` | Calendars.Read |
| /users/{id}/calendar/calendarView | ✅ | `orm.getCalendarView()` | Calendars.Read |
| /me/calendar/getSchedule | ✅ | `orm.getSchedule()` | Calendars.Read |
| /places/microsoft.graph.room | ✅ | `orm.rooms` | Place.Read.All |
| /places/microsoft.graph.roomList | ✅ | `orm.roomLists` | Place.Read.All |
| /users/{id}/findRooms | ✅ | `orm.findRooms()` | Calendars.Read |
| /users/{id}/findRoomLists | ✅ | `orm.findRoomLists()` | Calendars.Read |

### 3. 邮件 (Mail) ✅

| API 端点 | 状态 | ORM 方法 | 权限 |
|---------|------|----------|------|
| /users/{id}/messages | ✅ | `orm.messages()` | Mail.Read |
| /users/{id}/mailFolders/inbox/messages | ✅ | `orm.inbox()` | Mail.Read |
| /users/{id}/mailFolders/sentitems/messages | ✅ | `orm.sentItems()` | Mail.Read |
| /users/{id}/mailFolders/drafts/messages | ✅ | `orm.drafts()` | Mail.Read |
| /users/{id}/mailFolders | ✅ | `orm.mailFolders()` | Mail.Read |
| /users/{id}/mailFolders/{id}/messages | ✅ | `orm.messagesInFolder()` | Mail.Read |
| /users/{id}/mailFolders/inbox/messageRules | ✅ | `orm.messageRules()` | MailboxSettings.ReadWrite |
| /users/{id}/mailboxSettings | ✅ | `orm.getMailboxSettings()` | MailboxSettings.Read |
| /users/{id}/inferenceClassification | ✅ | `orm.getInferenceClassification()` | Mail.Read |
| /users/{id}/messages/{id}/attachments | ✅ | `orm.getMessageAttachments()` | Mail.Read |
| /users/{id}/sendMail | ✅ | `orm.sendMail()` | Mail.Send |

### 4. 个人联系人 (Contacts) ✅

| API 端点 | 状态 | ORM 方法 | 权限 |
|---------|------|----------|------|
| /users/{id}/contacts | ✅ | `orm.contacts()` | Contacts.Read |
| /users/{id}/contactFolders/{id}/contacts | ✅ | `orm.contactsInFolder()` | Contacts.Read |
| /users/{id}/contactFolders | ✅ | `orm.contactFolders()` | Contacts.Read |
| /contacts | ✅ | `orm.orgContacts` | OrgContact.Read.All |

### 5. 文件存储 (Files - OneDrive) ✅

| API 端点 | 状态 | ORM 方法 | 权限 |
|---------|------|----------|------|
| /drives | ✅ | `orm.drives` | Files.Read.All |
| /users/{id}/drive | ✅ | `orm.getUserDrive()` | Files.Read |
| /groups/{id}/drive | ✅ | `orm.getGroupDrive()` | Files.Read.All |
| /users/{id}/drive/root/children | ✅ | `orm.driveItems()` | Files.Read |
| /users/{id}/drive/items/{id} | ✅ | `orm.getDriveItem()` | Files.Read |
| /users/{id}/drive/root:{path} | ✅ | `orm.getDriveItemByPath()` | Files.Read |
| /users/{id}/drive/items/{id}/children | ✅ | `orm.listDriveItemChildren()` | Files.Read |
| /users/{id}/drive/items/{id}/createFolder | ✅ | `orm.createFolder()` | Files.ReadWrite |
| /users/{id}/drive/root:{path}:/content | ✅ | `orm.uploadSmallFile()` | Files.ReadWrite |
| /users/{id}/drive/items/{id}/content | ✅ | `orm.downloadFile()` | Files.Read |
| /users/{id}/drive/items/{id} (PATCH) | ✅ | `orm.updateDriveItem()` | Files.ReadWrite |
| /users/{id}/drive/items/{id} (DELETE) | ✅ | `orm.deleteDriveItem()` | Files.ReadWrite |
| /users/{id}/drive/items/{id}/permanentDelete | ✅ | `orm.permanentlyDeleteDriveItem()` | Files.ReadWrite |
| /users/{id}/drive/items/{id}/move | ✅ | `orm.moveDriveItem()` | Files.ReadWrite |
| /users/{id}/drive/items/{id}/copy | ✅ | `orm.copyDriveItem()` | Files.ReadWrite |
| /users/{id}/drive/root/search | ✅ | `orm.searchDriveItems()` | Files.Read |
| /users/{id}/drive/root/delta | ✅ | `orm.deltaDriveItems()` | Files.Read |
| /users/{id}/drive/items/{id}/follow | ✅ | `orm.followDriveItem()` | Files.ReadWrite |
| /users/{id}/drive/items/{id}/unfollow | ✅ | `orm.unfollowDriveItem()` | Files.ReadWrite |
| /users/{id}/drive/items/{id}/thumbnails | ✅ | `orm.getDriveItemThumbnails()` | Files.Read |
| /users/{id}/drive/items/{id}/createLink | ✅ | `orm.createDriveItemLink()` | Files.ReadWrite |
| /users/{id}/drive/items/{id}/invite | ✅ | `orm.inviteToDriveItem()` | Files.ReadWrite |
| /users/{id}/drive/items/{id}/permissions | ✅ | `orm.listDriveItemPermissions()` | Files.Read |
| /users/{id}/drive/items/{id}/preview | ✅ | `orm.previewDriveItem()` | Files.Read |
| /users/{id}/drive/items/{id}/versions | ✅ | `orm.listDriveItemVersions()` | Files.Read |
| /users/{id}/drive/items/{id}/analytics | ✅ | `orm.getDriveItemAnalytics()` | Files.Read |
| /me/drive/sharedWithMe | ✅ | `orm.getSharedWithMe()` | Files.Read.All |
| /me/drive/recent | ✅ | `orm.getRecentFiles()` | Files.Read |

### 6. OneNote ✅

| API 端点 | 状态 | ORM 方法 | 权限 | 国家云 |
|---------|------|----------|------|--------|
| /users/{id}/onenote/notebooks | ✅ | `orm.notebooks()` | Notes.Read | ⚠️ 中国版不可用 |
| /groups/{id}/onenote/notebooks | ✅ | `orm.groupNotebooks()` | Notes.Read | ⚠️ 中国版不可用 |
| /users/{id}/onenote/sections | ✅ | `orm.sections()` | Notes.Read | ⚠️ 中国版不可用 |
| /users/{id}/onenote/notebooks/{id}/sections | ✅ | `orm.notebookSections()` | Notes.Read | ⚠️ 中国版不可用 |
| /users/{id}/onenote/pages | ✅ | `orm.oneNotePages()` | Notes.Read | ⚠️ 中国版不可用 |
| /users/{id}/onenote/sections/{id}/pages | ✅ | `orm.sectionPages()` | Notes.Read | ⚠️ 中国版不可用 |
| /users/{id}/onenote/sections/{id}/pages (POST) | ✅ | `orm.createOneNotePage()` | Notes.ReadWrite | ⚠️ 中国版不可用 |

### 7. 任务和计划 (Tasks & Plans) ✅

| API 端点 | 状态 | ORM 方法 | 权限 | 国家云 |
|---------|------|----------|------|--------|
| /planner/plans | ✅ | `orm.plans` | Tasks.Read | ⚠️ 中国版不可用 |
| /groups/{id}/planner/plans | ✅ | `orm.groupPlans()` | Tasks.Read | ⚠️ 中国版不可用 |
| /planner/plans/{id}/tasks | ✅ | `orm.planTasks()` | Tasks.Read | ⚠️ 中国版不可用 |
| /users/{id}/planner/tasks | ✅ | `orm.userTasks()` | Tasks.Read | ⚠️ 中国版不可用 |
| /planner/plans/{id}/buckets | ✅ | `orm.planBuckets()` | Tasks.Read | ⚠️ 中国版不可用 |
| /planner/buckets/{id}/tasks | ✅ | `orm.bucketTasks()` | Tasks.Read | ⚠️ 中国版不可用 |
| /users/{id}/todo/lists | ✅ | `orm.todoLists()` | Tasks.Read | ⚠️ 中国版不可用 |
| /users/{id}/todo/lists/{id}/tasks | ✅ | `orm.todoTasks()` | Tasks.Read | ⚠️ 中国版不可用 |

### 8. SharePoint ✅

| API 端点 | 状态 | ORM 方法 | 权限 |
|---------|------|----------|------|
| /sites | ✅ | `orm.sites` | Sites.Read.All |
| /sites/root | ✅ | `orm.rootSite` | Sites.Read.All |
| /sites/{hostname}:{path} | ✅ | `orm.getSiteByPath()` | Sites.Read.All |
| /sites/{id}/drive/root/children | ✅ | `orm.siteDriveItems()` | Sites.Read.All |
| /sites/{id}/lists | ✅ | `orm.siteLists()` | Sites.Read.All |
| /sites/{id}/lists/{id}/items | ✅ | `orm.listItems()` | Sites.Read.All |
| /sites/{id}/lists/{id}/columns | ✅ | `orm.listColumns()` | Sites.Read.All |

### 9. Teams ✅

| API 端点 | 状态 | ORM 方法 | 权限 | 国家云 |
|---------|------|----------|------|--------|
| /teams | ✅ | `orm.teams` | Team.ReadBasic.All | ⚠️ 中国版受限 |
| /groups/{id}/team | ✅ | `orm.getTeamByGroup()` | Team.ReadBasic.All | ⚠️ 中国版受限 |
| /teams/{id}/channels | ✅ | `orm.channels()` | Channel.ReadBasic.All | ⚠️ 中国版受限 |
| /chats | ✅ | `orm.allChats` | Chat.Read | ⚠️ 中国版受限 |
| /users/{id}/chats | ✅ | `orm.chats()` | Chat.Read | ⚠️ 中国版受限 |
| /teams/{id}/channels/{id}/messages | ✅ | `orm.channelMessages()` | ChannelMessage.Read.All | ⚠️ 中国版受限 |
| /teams/{id}/channels/{id}/messages/{id}/replies | ✅ | `orm.channelMessageReplies()` | ChannelMessage.Read.All | ⚠️ 中国版受限 |
| /chats/{id}/messages | ✅ | `orm.chatMessages()` | ChatMessage.Read | ⚠️ 中国版受限 |
| /teams/{id}/channels/{id}/messages (POST) | ✅ | `orm.sendChannelMessage()` | ChannelMessage.Send | ⚠️ 中国版受限 |
| /chats/{id}/messages (POST) | ✅ | `orm.sendChatMessage()` | ChatMessage.Send | ⚠️ 中国版受限 |
| /teams/{id}/members | ✅ | `orm.teamMembers()` | TeamMember.Read.All | ⚠️ 中国版受限 |
| /teams/{id}/channels/{id}/members | ✅ | `orm.channelMembers()` | TeamMember.Read.All | ⚠️ 中国版受限 |
| /teams/{id}/installedApps | ✅ | `orm.teamApps()` | TeamsApp.Read.All | ⚠️ 中国版受限 |
| /teams/{id}/tags | ✅ | `orm.teamTags()` | TeamworkTag.Read | ⚠️ 中国版受限 |

### 10. 报告和分析 (Reports & Analytics) ✅

| API 端点 | 状态 | ORM 方法 | 权限 | 国家云 |
|---------|------|----------|------|--------|
| /auditLogs/directoryAudits | ✅ | `orm.directoryAudits` | AuditLog.Read.All | ✓ 全部支持 |
| /auditLogs/signIns | ✅ | `orm.signInLogs` | AuditLog.Read.All | ✓ 全部支持 |
| /reports/getOffice365ActiveUserDetail | ✅ | `orm.getOffice365ActiveUserDetail()` | Reports.Read.All | ⚠️ 中国版受限 |
| /reports/getEmailActivityUserDetail | ✅ | `orm.getEmailActivityUserDetail()` | Reports.Read.All | ⚠️ 中国版受限 |
| /reports/getOneDriveUsageAccountDetail | ✅ | `orm.getOneDriveUsageAccountDetail()` | Reports.Read.All | ⚠️ 中国版受限 |
| /reports/getSharePointActivityUserDetail | ✅ | `orm.getSharePointActivityUserDetail()` | Reports.Read.All | ⚠️ 中国版受限 |
| /reports/getTeamsUserActivityUserDetail | ✅ | `orm.getTeamsUserActivityUserDetail()` | Reports.Read.All | ⚠️ 中国版受限 |
| /users/{id}/insights/trending | ✅ | `orm.trendingInsights()` | Sites.Read.All | ⚠️ 中国版不可用 |
| /users/{id}/insights/used | ✅ | `orm.usedInsights()` | Sites.Read.All | ⚠️ 中国版不可用 |
| /users/{id}/insights/shared | ✅ | `orm.sharedInsights()` | Sites.Read.All | ⚠️ 中国版不可用 |
| /users/{id}/people | ✅ | `orm.people()` | People.Read | ⚠️ 中国版受限 |

### 11. 安全与合规 (Security & Compliance) ✅

| API 端点 | 状态 | ORM 方法 | 权限 | 国家云 |
|---------|------|----------|------|--------|
| /security/alerts | ✅ | `orm.securityAlerts` | SecurityEvents.Read.All | ⚠️ 中国版不可用 |
| /security/alerts/{id} (PATCH) | ✅ | `orm.updateSecurityAlert()` | SecurityEvents.ReadWrite.All | ⚠️ 中国版不可用 |
| /security/secureScores | ✅ | `orm.secureScores` | SecurityEvents.Read.All | ⚠️ 中国版不可用 |
| /identityProtection/riskDetections | ✅ | `orm.riskDetections` | IdentityRiskEvent.Read.All | ⚠️ 中国版不可用 |
| /identityProtection/riskyUsers | ✅ | `orm.riskyUsers` | IdentityRiskyUser.Read.All | ⚠️ 中国版不可用 |
| /identityProtection/riskyUsers/confirmCompromised | ✅ | `orm.confirmUsersCompromised()` | IdentityRiskyUser.ReadWrite.All | ⚠️ 中国版不可用 |
| /identityProtection/riskyUsers/dismiss | ✅ | `orm.dismissRiskyUsers()` | IdentityRiskyUser.ReadWrite.All | ⚠️ 中国版不可用 |

### 12. 扩展 (Extensions) ✅

| API 端点 | 状态 | ORM 方法 | 权限 |
|---------|------|----------|------|
| /schemaExtensions | ✅ | `orm.schemaExtensions` | Application.ReadWrite.All |
| /schemaExtensions (POST) | ✅ | `orm.createSchemaExtension()` | Application.ReadWrite.All |
| /{resource}/{id}/extensions | ✅ | `orm.getOpenExtensions()` | 根据资源类型 |
| /{resource}/{id}/extensions (POST) | ✅ | `orm.createOpenExtension()` | 根据资源类型 |

### 13. 其他功能 (Other Features) ✅

| API 端点 | 状态 | ORM 方法 | 权限 | 国家云 |
|---------|------|----------|------|--------|
| /search/query | ✅ | `orm.search()` | 根据搜索资源 | ⚠️ 中国版不可用 |
| /search/query (快捷方法) | ✅ | `orm.searchUsers()`, `orm.searchMessages()`, `orm.searchFiles()` | 根据搜索资源 | ⚠️ 中国版不可用 |
| /$batch | ✅ | `orm.batch()` | 根据批处理请求 | ✓ 全部支持 |
| /users/delta | ✅ | `orm.deltaUsers()` | User.Read.All | ✓ 全部支持 |
| /groups/delta | ✅ | `orm.deltaGroups()` | Group.Read.All | ✓ 全部支持 |
| /users/{id}/messages/delta | ✅ | `orm.deltaMessages()` | Mail.Read | ✓ 全部支持 |
| /users/{id}/events/delta | ✅ | `orm.deltaEvents()` | Calendars.Read | ✓ 全部支持 |
| /users/{id}/drive/root/delta | ✅ | `orm.deltaDriveItems()` | Files.Read | ✓ 全部支持 |
| /subscriptions | ✅ | `orm.subscriptions` | 根据订阅资源 | ✓ 全部支持 |
| /subscriptions (POST) | ✅ | `orm.createSubscription()` | 根据订阅资源 | ✓ 全部支持 |
| /subscriptions/{id} (PATCH) | ✅ | `orm.renewSubscription()` | 根据订阅资源 | ✓ 全部支持 |
| /subscriptions/{id} (DELETE) | ✅ | `orm.deleteSubscription()` | 根据订阅资源 | ✓ 全部支持 |

## 未实现的 API（低优先级/特殊场景）

以下 API 端点未实现，主要是因为它们属于特殊场景或使用频率较低：

### 1. 高级日历功能

- ❌ `/users/{id}/calendarView` - 已有替代方法 `getCalendarView()`
- ❌ `/users/{id}/calendar/allowedCalendarSharingRoles` - 日历共享角色查询
- ❌ `/users/{id}/events/{id}/instances` - 定期事件实例

### 2. Excel 工作簿 API

- ❌ `/me/drive/items/{id}/workbook` - Excel 工作簿操作
- ❌ `/me/drive/items/{id}/workbook/worksheets` - 工作表操作
- ⚠️ **原因**: 中国版功能受限，且这是特殊场景 API

### 3. 高级邮件功能

- ❌ `/users/{id}/messages/{id}/move` - 移动邮件（可用 PATCH 实现）
- ❌ `/users/{id}/messages/{id}/copy` - 复制邮件
- ❌ `/users/{id}/messages/{id}/createReply` - 创建回复草稿
- ❌ `/users/{id}/messages/{id}/createReplyAll` - 创建全部回复草稿
- ❌ `/users/{id}/messages/{id}/createForward` - 创建转发草稿

### 4. 高级 Teams 功能

- ❌ `/teams/{id}/schedule` - Teams 排班
- ❌ `/teams/{id}/channels/{id}/tabs` - 频道标签页
- ❌ `/teams/{id}/primaryChannel` - 主频道
- ❌ `/appCatalogs/teamsApps` - Teams 应用目录

### 5. 高级安全功能

- ❌ `/informationProtection` - 信息保护
- ❌ `/security/tiIndicators` - 威胁情报指标
- ❌ `/security/secureScoreControlProfiles` - 安全评分控制配置文件

### 6. 高级搜索和索引

- ❌ `/external/connections` - 外部连接
- ❌ `/search/acronyms` - 首字母缩略词搜索
- ❌ `/search/bookmarks` - 书签搜索

### 7. 合规性和记录管理

- ❌ `/compliance/ediscovery` - 电子发现
- ❌ `/privacy/subjectRightsRequests` - 主体权利请求

### 8. 移动设备管理 (MDM/MAM)

- ❌ `/deviceManagement` - Intune 设备管理相关 API
- ⚠️ **原因**: 这是独立的 Intune API 集，需要单独支持

## 已实现但有限制的功能

以下功能已实现，但在某些国家云中可能不可用或功能受限：

| 功能 | 限制说明 |
|------|---------|
| OneNote API | 中国版（21Vianet）不可用 |
| Planner & To Do | 中国版不可用 |
| Teams 高级功能 | 中国版功能受限，GCC High/DoD 部分功能不可用 |
| 洞察分析 (Insights) | 中国版不可用，美国政府版部分功能受限 |
| 安全警报和风险检测 | 中国版不可用，需要 Azure AD P2 许可 |
| Microsoft Search | 中国版不可用 |
| Places API | 中国版功能受限，建议使用 findRooms 替代 |

## 实现方式说明

### EntityManager 方法

大多数集合类型的端点通过 `GraphEntityManager` 实现，支持：

- `findById()` - 获取单个实体
- `findMany()` - 获取多个实体
- `create()` - 创建实体
- `update()` - 更新实体
- `delete()` - 删除实体
- `query()` - 查询构建器

### 自定义方法

特殊的 API 端点通过自定义方法实现：

- 函数调用（如 `findRooms`, `getSchedule`）
- 操作（如 `sendMail`, `assignUserLicense`）
- 批处理和 Delta 查询

## 总结

✅ **核心功能**: 100% 覆盖
✅ **常用 API**: 95%+ 覆盖
⚠️ **高级功能**: 部分覆盖
❌ **特殊场景**: 按需实现

本 SDK 已覆盖 Microsoft Graph v1.0 中绝大多数常用的 API 端点，能够满足大部分应用场景的需求。未实现的 API 主要是低频使用的高级功能或特定场景的 API，如有需要可以基于现有架构快速扩展。

## 参考文档

- [Microsoft Graph REST API v1.0 参考](https://learn.microsoft.com/graph/api/overview?view=graph-rest-1.0)
- [Microsoft Graph 权限参考](https://learn.microsoft.com/graph/permissions-reference)
- [Microsoft Graph 国家云部署](https://learn.microsoft.com/graph/deployments)
- [本项目完整 API 参考](./GRAPH_API_REFERENCE.md)
