# Microsoft Graph REST API v1.0 完整接口梳理

## 目录

- [1. 用户和身份管理 (Users & Identity)](#1-用户和身份管理)
- [2. 日历 (Calendar)](#2-日历)
- [3. 邮件 (Mail)](#3-邮件)
- [4. 个人联系人 (Personal Contacts)](#4-个人联系人)
- [5. 文件存储 (Files - OneDrive)](#5-文件存储)
- [6. 笔记 (Notes - OneNote)](#6-笔记)
- [7. 任务和计划 (Tasks & Plans - Planner)](#7-任务和计划)
- [8. 站点和列表 (Sites & Lists - SharePoint)](#8-站点和列表)
- [9. 团队协作 (Teams)](#9-团队协作)
- [10. 报告和分析 (Reports & Analytics)](#10-报告和分析)
- [11. 安全与合规 (Security & Compliance)](#11-安全与合规)
- [12. 通知 (Notifications)](#12-通知)
- [13. 扩展 (Extensions)](#13-扩展)
- [14. 其他功能 (Other Features)](#14-其他功能)

---

## 1. 用户和身份管理

### 1.1 用户 (Users)

- **状态**: ✅ 已实现
- **端点**: `/users`
- **权限**: User.Read.All, User.ReadWrite.All, Directory.Read.All
- **国家云**: ✓ 全球版 ✓ 中国版 ✓ 美国政府版
- **方法**: `orm.users`

### 1.2 组 (Groups)

- **状态**: ✅ 已实现
- **端点**: `/groups`
- **权限**: Group.Read.All, Group.ReadWrite.All
- **国家云**: ✓ 全球版 ✓ 中国版 ✓ 美国政府版
- **方法**: `orm.groups`, `orm.groupMembers()`, `orm.groupOwners()`

### 1.3 设备 (Devices)

- **状态**: ✅ 已实现
- **端点**: `/devices`
- **权限**: Device.Read.All, Device.ReadWrite.All
- **国家云**: ✓ 全球版 ✓ 中国版 ✓ 美国政府版
- **方法**: `orm.devices`

### 1.4 应用程序 (Applications)

- **状态**: ✅ 已实现
- **端点**: `/applications`
- **权限**: Application.Read.All, Application.ReadWrite.All
- **国家云**: ✓ 全球版 ✓ 中国版 ✓ 美国政府版
- **方法**: `orm.applications`

### 1.5 服务主体 (Service Principals)

- **状态**: ✅ 已实现
- **端点**: `/servicePrincipals`
- **权限**: Application.Read.All, Application.ReadWrite.All
- **国家云**: ✓ 全球版 ✓ 中国版 ✓ 美国政府版
- **方法**: `orm.servicePrincipals`

### 1.6 目录角色 (Directory Roles)

- **状态**: ⏳ 待实现
- **端点**: `/directoryRoles`, `/directoryRoleTemplates`
- **权限**: RoleManagement.Read.Directory, RoleManagement.ReadWrite.Directory
- **国家云**: ✓ 全球版 ✓ 中国版 ✓ 美国政府版

### 1.7 管理单元 (Administrative Units)

- **状态**: ⏳ 待实现
- **端点**: `/directory/administrativeUnits`
- **权限**: AdministrativeUnit.Read.All, AdministrativeUnit.ReadWrite.All
- **国家云**: ✓ 全球版 ✓ 中国版 ✓ 美国政府版

### 1.8 许可证 (Licenses)

- **状态**: ⏳ 待实现
- **端点**: `/users/{id}/licenseDetails`, `/subscribedSkus`
- **权限**: User.Read.All, Organization.Read.All
- **国家云**: ✓ 全球版 ✓ 中国版 ✓ 美国政府版

### 1.9 邀请 (Invitations)

- **状态**: ⏳ 待实现
- **端点**: `/invitations`
- **权限**: User.Invite.All
- **国家云**: ✓ 全球版 ⚠️ 中国版（功能受限）✓ 美国政府版

### 1.10 组织 (Organization)

- **状态**: ✅ 已实现（部分）
- **端点**: `/organization`
- **权限**: Organization.Read.All, Organization.ReadWrite.All
- **国家云**: ✓ 全球版 ✓ 中国版 ✓ 美国政府版
- **方法**: `orm.organization()`

---

## 2. 日历

### 2.1 事件 (Events)

- **状态**: ✅ 已实现
- **端点**: `/users/{id}/events`, `/groups/{id}/events`
- **权限**: Calendars.Read, Calendars.ReadWrite
- **国家云**: ✓ 全球版 ✓ 中国版 ✓ 美国政府版
- **方法**: `orm.events()`, `orm.calendarEvents()`, `orm.groupEvents()`

### 2.2 日历 (Calendars)

- **状态**: ⏳ 待实现
- **端点**: `/users/{id}/calendars`, `/users/{id}/calendar`
- **权限**: Calendars.Read, Calendars.ReadWrite
- **国家云**: ✓ 全球版 ✓ 中国版 ✓ 美国政府版

### 2.3 日历组 (Calendar Groups)

- **状态**: ⏳ 待实现
- **端点**: `/users/{id}/calendarGroups`
- **权限**: Calendars.Read, Calendars.ReadWrite
- **国家云**: ✓ 全球版 ✓ 中国版 ✓ 美国政府版

### 2.4 会议室和场所 (Places/Rooms)

- **状态**: ✅ 已实现
- **端点**: `/places`, `/users/{id}/findRooms`
- **权限**: Place.Read.All, Calendars.Read
- **国家云**: ✓ 全球版 ⚠️ 中国版（功能受限）✓ 美国政府版（GCC）
- **方法**: `orm.rooms`, `orm.findRooms()`, `orm.findRoomLists()`

---

## 3. 邮件

### 3.1 邮件消息 (Messages)

- **状态**: ✅ 已实现
- **端点**: `/users/{id}/messages`
- **权限**: Mail.Read, Mail.ReadWrite
- **国家云**: ✓ 全球版 ✓ 中国版 ✓ 美国政府版
- **方法**: `orm.messages()`, `orm.inbox()`, `orm.sentItems()`, `orm.drafts()`

### 3.2 邮件文件夹 (Mail Folders)

- **状态**: ⏳ 待实现
- **端点**: `/users/{id}/mailFolders`
- **权限**: Mail.Read, Mail.ReadWrite
- **国家云**: ✓ 全球版 ✓ 中国版 ✓ 美国政府版

### 3.3 邮件规则 (Mail Rules)

- **状态**: ⏳ 待实现
- **端点**: `/users/{id}/mailFolders/inbox/messageRules`
- **权限**: MailboxSettings.ReadWrite
- **国家云**: ✓ 全球版 ✓ 中国版 ✓ 美国政府版

### 3.4 邮箱设置 (Mailbox Settings)

- **状态**: ⏳ 待实现
- **端点**: `/users/{id}/mailboxSettings`
- **权限**: MailboxSettings.Read, MailboxSettings.ReadWrite
- **国家云**: ✓ 全球版 ✓ 中国版 ✓ 美国政府版

### 3.5 重点收件箱 (Focused Inbox)

- **状态**: ⏳ 待实现
- **端点**: `/users/{id}/inferenceClassification`
- **权限**: Mail.Read, Mail.ReadWrite
- **国家云**: ✓ 全球版 ✓ 中国版 ✓ 美国政府版

### 3.6 邮件附件 (Attachments)

- **状态**: ⏳ 待实现
- **端点**: `/users/{id}/messages/{id}/attachments`
- **权限**: Mail.Read, Mail.ReadWrite
- **国家云**: ✓ 全球版 ✓ 中国版 ✓ 美国政府版

---

## 4. 个人联系人

### 4.1 联系人 (Contacts)

- **状态**: ✅ 已实现
- **端点**: `/users/{id}/contacts`
- **权限**: Contacts.Read, Contacts.ReadWrite
- **国家云**: ✓ 全球版 ✓ 中国版 ✓ 美国政府版
- **方法**: `orm.contacts()`, `orm.contactsInFolder()`

### 4.2 联系人文件夹 (Contact Folders)

- **状态**: ⏳ 待实现
- **端点**: `/users/{id}/contactFolders`
- **权限**: Contacts.Read, Contacts.ReadWrite
- **国家云**: ✓ 全球版 ✓ 中国版 ✓ 美国政府版

### 4.3 组织联系人 (Org Contacts)

- **状态**: ⏳ 待实现
- **端点**: `/contacts`
- **权限**: OrgContact.Read.All
- **国家云**: ✓ 全球版 ✓ 中国版 ✓ 美国政府版

---

## 5. 文件存储

### 5.1 驱动器 (Drives)

- **状态**: ⏳ 待实现
- **端点**: `/drives`, `/users/{id}/drive`, `/groups/{id}/drive`
- **权限**: Files.Read, Files.ReadWrite, Files.Read.All, Files.ReadWrite.All
- **国家云**: ✓ 全球版 ✓ 中国版 ✓ 美国政府版

### 5.2 驱动器项 (Drive Items)

- **状态**: ✅ 已实现（完整）
- **端点**: `/users/{id}/drive/items`, `/drives/{id}/items`
- **权限**: Files.Read, Files.ReadWrite, Files.Read.All, Files.ReadWrite.All
- **国家云**: ✓ 全球版 ✓ 中国版 ✓ 美国政府版
- **方法**:
  - `orm.driveItems()` - 列出项
  - `orm.getDriveItem()` - 获取项
  - `orm.createFolder()` - 创建文件夹
  - `orm.uploadSmallFile()` - 上传文件
  - `orm.downloadFile()` - 下载文件
  - `orm.moveDriveItem()` - 移动项
  - `orm.copyDriveItem()` - 复制项
  - `orm.searchDriveItems()` - 搜索项

### 5.3 共享项 (Shared Items)

- **状态**: ⏳ 待实现
- **端点**: `/me/drive/sharedWithMe`, `/drives/{id}/sharedWithMe`
- **权限**: Files.Read.All, Files.ReadWrite.All
- **国家云**: ✓ 全球版 ✓ 中国版 ✓ 美国政府版

### 5.4 最近使用 (Recent Items)

- **状态**: ⏳ 待实现
- **端点**: `/me/drive/recent`
- **权限**: Files.Read, Files.ReadWrite
- **国家云**: ✓ 全球版 ✓ 中国版 ✓ 美国政府版

### 5.5 Excel 工作簿 (Workbooks)

- **状态**: ⏳ 待实现
- **端点**: `/me/drive/items/{id}/workbook`
- **权限**: Files.ReadWrite
- **国家云**: ✓ 全球版 ⚠️ 中国版（功能受限）✓ 美国政府版

---

## 6. 笔记

### 6.1 OneNote 笔记本 (Notebooks)

- **状态**: ⏳ 待实现
- **端点**: `/users/{id}/onenote/notebooks`
- **权限**: Notes.Read, Notes.ReadWrite, Notes.Read.All, Notes.ReadWrite.All
- **国家云**: ✓ 全球版 ⚠️ 中国版（不可用）✓ 美国政府版

### 6.2 OneNote 分区 (Sections)

- **状态**: ⏳ 待实现
- **端点**: `/users/{id}/onenote/sections`
- **权限**: Notes.Read, Notes.ReadWrite
- **国家云**: ✓ 全球版 ⚠️ 中国版（不可用）✓ 美国政府版

### 6.3 OneNote 页面 (Pages)

- **状态**: ⏳ 待实现
- **端点**: `/users/{id}/onenote/pages`
- **权限**: Notes.Read, Notes.ReadWrite
- **国家云**: ✓ 全球版 ⚠️ 中国版（不可用）✓ 美国政府版

---

## 7. 任务和计划

### 7.1 Planner 计划 (Plans)

- **状态**: ✅ 已实现
- **端点**: `/planner/plans`, `/groups/{id}/planner/plans`
- **权限**: Tasks.Read, Tasks.ReadWrite, Group.Read.All
- **国家云**: ✓ 全球版 ⚠️ 中国版（不可用）✓ 美国政府版（GCC）
- **方法**: `orm.plans`, `orm.groupPlans()`

### 7.2 Planner 任务 (Tasks)

- **状态**: ✅ 已实现
- **端点**: `/planner/tasks`, `/users/{id}/planner/tasks`
- **权限**: Tasks.Read, Tasks.ReadWrite
- **国家云**: ✓ 全球版 ⚠️ 中国版（不可用）✓ 美国政府版（GCC）
- **方法**: `orm.planTasks()`, `orm.userTasks()`

### 7.3 Planner 存储桶 (Buckets)

- **状态**: ⏳ 待实现
- **端点**: `/planner/plans/{id}/buckets`
- **权限**: Tasks.Read, Tasks.ReadWrite
- **国家云**: ✓ 全球版 ⚠️ 中国版（不可用）✓ 美国政府版（GCC）

### 7.4 To Do 任务 (To Do Tasks)

- **状态**: ⏳ 待实现
- **端点**: `/me/todo/lists`, `/me/todo/lists/{id}/tasks`
- **权限**: Tasks.Read, Tasks.ReadWrite
- **国家云**: ✓ 全球版 ⚠️ 中国版（不可用）✓ 美国政府版

---

## 8. 站点和列表

### 8.1 SharePoint 站点 (Sites)

- **状态**: ✅ 已实现（基础）
- **端点**: `/sites`, `/sites/{id}`
- **权限**: Sites.Read.All, Sites.ReadWrite.All
- **国家云**: ✓ 全球版 ✓ 中国版 ✓ 美国政府版
- **方法**: `orm.sites`, `orm.rootSite`, `orm.siteDriveItems()`

### 8.2 SharePoint 列表 (Lists)

- **状态**: ⏳ 待实现
- **端点**: `/sites/{id}/lists`
- **权限**: Sites.Read.All, Sites.ReadWrite.All
- **国家云**: ✓ 全球版 ✓ 中国版 ✓ 美国政府版

### 8.3 列表项 (List Items)

- **状态**: ⏳ 待实现
- **端点**: `/sites/{id}/lists/{id}/items`
- **权限**: Sites.Read.All, Sites.ReadWrite.All
- **国家云**: ✓ 全球版 ✓ 中国版 ✓ 美国政府版

### 8.4 列表列 (List Columns)

- **状态**: ⏳ 待实现
- **端点**: `/sites/{id}/lists/{id}/columns`
- **权限**: Sites.Read.All, Sites.ReadWrite.All
- **国家云**: ✓ 全球版 ✓ 中国版 ✓ 美国政府版

---

## 9. 团队协作

### 9.1 Teams 团队 (Teams)

- **状态**: ✅ 已实现
- **端点**: `/teams`, `/groups/{id}/team`
- **权限**: Team.ReadBasic.All, TeamSettings.Read.All
- **国家云**: ✓ 全球版 ⚠️ 中国版（功能受限）✓ 美国政府版（GCC）
- **方法**: `orm.teams`, `orm.getTeamByGroup()`

### 9.2 Teams 频道 (Channels)

- **状态**: ✅ 已实现
- **端点**: `/teams/{id}/channels`
- **权限**: Channel.ReadBasic.All, ChannelSettings.Read.All
- **国家云**: ✓ 全球版 ⚠️ 中国版（功能受限）✓ 美国政府版（GCC）
- **方法**: `orm.channels()`

### 9.3 Teams 聊天 (Chats)

- **状态**: ✅ 已实现
- **端点**: `/chats`, `/users/{id}/chats`
- **权限**: Chat.Read, Chat.ReadWrite
- **国家云**: ✓ 全球版 ⚠️ 中国版（功能受限）✓ 美国政府版（GCC）
- **方法**: `orm.chats()`, `orm.allChats`

### 9.4 频道消息 (Channel Messages)

- **状态**: ⏳ 待实现
- **端点**: `/teams/{id}/channels/{id}/messages`
- **权限**: ChannelMessage.Read.All
- **国家云**: ✓ 全球版 ⚠️ 中国版（功能受限）✓ 美国政府版（GCC）

### 9.5 聊天消息 (Chat Messages)

- **状态**: ⏳ 待实现
- **端点**: `/chats/{id}/messages`
- **权限**: ChatMessage.Read, Chat.ReadWrite
- **国家云**: ✓ 全球版 ⚠️ 中国版（功能受限）✓ 美国政府版（GCC）

### 9.6 团队成员 (Team Members)

- **状态**: ⏳ 待实现
- **端点**: `/teams/{id}/members`
- **权限**: TeamMember.Read.All, TeamMember.ReadWrite.All
- **国家云**: ✓ 全球版 ⚠️ 中国版（功能受限）✓ 美国政府版（GCC）

### 9.7 团队应用 (Team Apps)

- **状态**: ⏳ 待实现
- **端点**: `/teams/{id}/installedApps`
- **权限**: TeamsApp.Read.All, TeamsAppInstallation.ReadForTeam
- **国家云**: ✓ 全球版 ⚠️ 中国版（功能受限）✓ 美国政府版（GCC）

### 9.8 团队标签 (Team Tags)

- **状态**: ⏳ 待实现
- **端点**: `/teams/{id}/tags`
- **权限**: TeamworkTag.Read, TeamworkTag.ReadWrite
- **国家云**: ✓ 全球版 ⚠️ 中国版（功能受限）✓ 美国政府版（GCC）

---

## 10. 报告和分析

### 10.1 使用报告 (Usage Reports)

- **状态**: ⏳ 待实现
- **端点**: `/reports/getOffice365ActiveUserDetail`, `/reports/getEmailActivityUserDetail`
- **权限**: Reports.Read.All
- **国家云**: ✓ 全球版 ⚠️ 中国版（功能受限）✓ 美国政府版（GCC）

### 10.2 审计日志 (Audit Logs)

- **状态**: ⏳ 待实现
- **端点**: `/auditLogs/directoryAudits`, `/auditLogs/signIns`
- **权限**: AuditLog.Read.All, Directory.Read.All
- **国家云**: ✓ 全球版 ✓ 中国版 ✓ 美国政府版

### 10.3 洞察分析 (Insights)

- **状态**: ⏳ 待实现
- **端点**: `/me/insights/trending`, `/me/insights/used`, `/me/insights/shared`
- **权限**: Sites.Read.All, Sites.ReadWrite.All
- **国家云**: ✓ 全球版 ⚠️ 中国版（不可用）✓ 美国政府版（部分）

### 10.4 人员分析 (People)

- **状态**: ⏳ 待实现
- **端点**: `/me/people`, `/users/{id}/people`
- **权限**: People.Read, People.Read.All
- **国家云**: ✓ 全球版 ⚠️ 中国版（功能受限）✓ 美国政府版

---

## 11. 安全与合规

### 11.1 安全警报 (Security Alerts)

- **状态**: ⏳ 待实现
- **端点**: `/security/alerts`
- **权限**: SecurityEvents.Read.All, SecurityEvents.ReadWrite.All
- **国家云**: ✓ 全球版 ⚠️ 中国版（不可用）✓ 美国政府版

### 11.2 安全评分 (Secure Score)

- **状态**: ⏳ 待实现
- **端点**: `/security/secureScores`
- **权限**: SecurityEvents.Read.All
- **国家云**: ✓ 全球版 ⚠️ 中国版（不可用）✓ 美国政府版

### 11.3 威胁评估 (Threat Assessment)

- **状态**: ⏳ 待实现
- **端点**: `/informationProtection/threatAssessmentRequests`
- **权限**: ThreatAssessment.ReadWrite.All
- **国家云**: ✓ 全球版 ⚠️ 中国版（不可用）✓ 美国政府版

### 11.4 风险检测 (Risk Detections)

- **状态**: ⏳ 待实现
- **端点**: `/identityProtection/riskDetections`
- **权限**: IdentityRiskEvent.Read.All
- **国家云**: ✓ 全球版 ⚠️ 中国版（不可用）✓ 美国政府版（P2许可）

### 11.5 风险用户 (Risky Users)

- **状态**: ⏳ 待实现
- **端点**: `/identityProtection/riskyUsers`
- **权限**: IdentityRiskyUser.Read.All
- **国家云**: ✓ 全球版 ⚠️ 中国版（不可用）✓ 美国政府版（P2许可）

---

## 12. 通知

### 12.1 订阅 (Subscriptions/Webhooks)

- **状态**: ✅ 已实现
- **端点**: `/subscriptions`
- **权限**: 根据订阅的资源而定
- **国家云**: ✓ 全球版 ✓ 中国版 ✓ 美国政府版
- **方法**: `orm.subscriptions`, `orm.createSubscription()`, `orm.renewSubscription()`

---

## 13. 扩展

### 13.1 架构扩展 (Schema Extensions)

- **状态**: ⏳ 待实现
- **端点**: `/schemaExtensions`
- **权限**: Application.ReadWrite.All
- **国家云**: ✓ 全球版 ✓ 中国版 ✓ 美国政府版

### 13.2 开放扩展 (Open Extensions)

- **状态**: ⏳ 待实现
- **端点**: `/users/{id}/extensions`, `/groups/{id}/extensions`
- **权限**: 根据资源类型而定
- **国家云**: ✓ 全球版 ✓ 中国版 ✓ 美国政府版

---

## 14. 其他功能

### 14.1 搜索 (Search)

- **状态**: ✅ 已实现
- **端点**: `/search/query`
- **权限**: 根据搜索的资源而定
- **国家云**: ✓ 全球版 ⚠️ 中国版（不可用）✓ 美国政府版（GCC）
- **方法**: `orm.search()`, `orm.searchMessages()`, `orm.searchFiles()`

### 14.2 批处理 (Batch)

- **状态**: ✅ 已实现
- **端点**: `/$batch`
- **权限**: 根据批处理中的请求而定
- **国家云**: ✓ 全球版 ✓ 中国版 ✓ 美国政府版
- **方法**: `orm.batch()`

### 14.3 Delta 查询 (Delta Query)

- **状态**: ✅ 已实现（部分）
- **端点**: `/users/delta`, `/groups/delta`, `/messages/delta`
- **权限**: 根据资源类型而定
- **国家云**: ✓ 全球版 ✓ 中国版 ✓ 美国政府版
- **方法**: `orm.deltaUsers()`, `orm.deltaGroups()`, `orm.deltaMessages()`, `orm.deltaEvents()`, `orm.deltaDriveItems()`

### 14.4 当前用户 (Me)

- **状态**: ✅ 已实现
- **端点**: `/me`
- **权限**: User.Read（委托权限）
- **国家云**: ✓ 全球版 ✓ 中国版 ✓ 美国政府版
- **方法**: `orm.me()`

### 14.5 目录对象 (Directory Objects)

- **状态**: ⏳ 待实现
- **端点**: `/directoryObjects/{id}`, `/directoryObjects/getByIds`
- **权限**: Directory.Read.All
- **国家云**: ✓ 全球版 ✓ 中国版 ✓ 美国政府版

---

## 实现进度统计

- ✅ 已实现: 20 个类别
- ⏳ 待实现: 50+ 个类别
- 总计: 70+ 个主要 API 类别

## 优先级

### P0 - 高优先级（核心功能）

1. 目录角色和管理单元
2. 日历和日历组
3. 邮件文件夹和规则
4. SharePoint 列表
5. Teams 消息

### P1 - 中优先级（增强功能）

1. OneNote 笔记
2. To Do 任务
3. 报告和分析
4. 人员分析
5. 邮件附件

### P2 - 低优先级（高级功能）

1. 安全警报和风险检测
2. 架构扩展
3. Excel 工作簿 API
4. 威胁评估
5. 团队应用和标签

---

## 参考文档

- [Microsoft Graph REST API v1.0 参考](https://learn.microsoft.com/zh-cn/graph/api/overview?view=graph-rest-1.0)
- [Microsoft Graph 权限参考](https://learn.microsoft.com/zh-cn/graph/permissions-reference)
- [Microsoft Graph 国家云部署](https://learn.microsoft.com/zh-cn/graph/deployments)

