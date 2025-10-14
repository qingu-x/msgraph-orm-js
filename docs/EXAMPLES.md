# Graph ORM 使用示例（新版本）

本文档展示如何使用新版 Graph ORM 的各种功能。

## 目录

- [初始化](#初始化)
- [用户管理](#用户管理)
- [组管理](#组管理)
- [文件操作](#文件操作)
- [邮件操作](#邮件操作)
- [日历操作](#日历操作)
- [高级查询](#高级查询)
- [增量查询](#增量查询)
- [批处理](#批处理)

## 初始化

```typescript
import { GraphClient, defaultEndpoints, createGraphORM } from '@qingu-x/msgraph-orm-js';

// 方式一：使用客户端凭据（应用权限）
const graphClient = new GraphClient(
  {
    tenantId: 'your-tenant-id',
    clientId: 'your-client-id',
    clientSecret: 'your-client-secret'
  },
  defaultEndpoints.global  // 全球版端点
);

const client = graphClient.getGraphClient();

// 创建 ORM 实例
const orm = createGraphORM(client);
```

### 使用委托权限（可选）

如果你有现成的访问令牌：

```typescript
import { Client } from '@microsoft/microsoft-graph-client';
import { createGraphORM } from '@qingu-x/msgraph-orm-js';

// 使用自定义认证提供者
const client = Client.initWithMiddleware({
  authProvider: {
    getAccessToken: async () => {
      // 返回访问令牌
      return 'your-access-token';
    }
  }
});

// 创建 ORM 实例
const orm = createGraphORM(client);
```

## 架构说明：Repository vs Service

新版 ORM 采用清晰的架构设计，所有请求都通过统一的 QueryBuilder 处理：

### Repository（仓储）- 实体 CRUD 和查询

用于实体的标准 CRUD 操作，支持查询构建器：

```typescript
// ✅ 使用 Repository 的场景：
// 1. 需要复杂查询条件
const unreadMessages = await orm.users.messages('user-id')
  .query()
  .where('isRead', 'eq', false)
  .orderBy('receivedDateTime', 'desc')
  .top(10)
  .get();

// 2. 标准 CRUD 操作
const event = await orm.users.events('user-id').findById('event-id');
await orm.users.events('user-id').create({ subject: '会议' });
await orm.users.events('user-id').update('event-id', { subject: '新标题' });
await orm.users.events('user-id').delete('event-id');

// 3. 实体特定操作
await orm.users.messages('user-id').reply('message-id', '收到');
await orm.users.events('user-id').accept('event-id', '我会参加');

// 4. 调试和监控（新功能）
const query = orm.users.query().where('department', 'eq', 'IT');
const debugInfo = query.getDebug();
console.log('请求信息:', debugInfo);
```

### Service（服务）- 业务逻辑

用于特定领域的业务功能：

```typescript
// ✅ 使用 Service 的场景：
// 1. 业务操作（不是简单的 CRUD）
await orm.mail.send('user-id', message);  // 发送邮件
await orm.files.uploadSmallFile('user-id', path, buffer);  // 上传文件

// 2. 特殊 API
const calendarView = await orm.calendar.getCalendarView('user-id', start, end);
const rooms = await orm.calendar.findRooms('user-id');

// 3. 快捷访问
const inbox = await orm.mail.getInbox('user-id', 20);  // 直接获取收件箱
const drafts = await orm.mail.getDrafts('user-id');    // 直接获取草稿箱

// 4. 调试和监控（新功能）
const debugInfo = orm.mail.getDebug();  // 获取最后请求的调试信息
console.log('邮件服务调试信息:', debugInfo);
```

### 选择原则

| 功能 | 使用 Repository | 使用 Service |
|------|----------------|--------------|
| 查询过滤 | ✅ | ❌ |
| 复杂条件 | ✅ | ❌ |
| CRUD 操作 | ✅ | 部分支持 |
| 业务逻辑 | ❌ | ✅ |
| 特殊 API | ❌ | ✅ |
| 快捷方法 | ❌ | ✅ |
| 调试功能 | ✅ | ✅ |
| 自定义 Headers | ✅ | ✅ |

## 用户管理

### 基本操作

```typescript
// 获取所有用户
const users = await orm.users.findMany();
console.log('用户总数:', users.meta.count);
console.log('用户列表:', users.data);

// 通过 ID 获取用户
const user = await orm.users.findById('user-id');
console.log('用户:', user.displayName);

// 通过邮箱查找用户
const userByEmail = await orm.users.findByEmail('user@example.com');

// 通过 UPN 查找用户
const userByUpn = await orm.users.findByUserPrincipalName('user@contoso.com');

// 检查用户是否存在
const exists = await orm.users.exists('user-id');

// 创建用户
const newUser = await orm.users.create({
  displayName: '张三',
  userPrincipalName: 'zhangsan@contoso.com',
  mailNickname: 'zhangsan',
  accountEnabled: true,
  passwordProfile: {
    password: 'TempPassword123!',
    forceChangePasswordNextSignIn: true
  }
});

// 更新用户
const updatedUser = await orm.users.update('user-id', {
  jobTitle: '高级工程师',
  department: '技术部'
});

// 删除用户
await orm.users.delete('user-id');
```

### 用户关系

```typescript
// 获取用户所属的组
const userGroups = await orm.users.groups('user-id').findMany();

// 获取用户的经理
const manager = await orm.users.getManager('user-id');

// 获取用户的直接下属
const directReports = await orm.users.directReports('user-id').findMany();

// 获取用户的邮件（使用 Repository，支持查询构建器）
const messages = await orm.users.messages('user-id')
  .query()
  .where('isRead', 'eq', false)
  .top(10)
  .get();

// 获取用户的事件（使用 Repository，支持查询构建器）
const events = await orm.users.events('user-id')
  .query()
  .where('start/dateTime', 'ge', new Date().toISOString())
  .orderBy('start/dateTime', 'asc')
  .get();

// 获取用户的文件（使用 Repository，支持查询构建器）
const files = await orm.users.driveItems('user-id')
  .query()
  .where('file', 'ne', null)
  .orderBy('lastModifiedDateTime', 'desc')
  .get();
```

### 用户许可证

```typescript
// 获取用户许可证
const licenses = await orm.users.getLicenses('user-id');

// 分配许可证
await orm.users.assignLicense('user-id', [
  {
    skuId: 'sku-id',
    disabledPlans: []
  }
], []);
```

### 用户设置

```typescript
// 获取邮箱设置
const mailboxSettings = await orm.users.getMailboxSettings('user-id');

// 更新邮箱设置
await orm.users.updateMailboxSettings('user-id', {
  timeZone: 'China Standard Time',
  language: {
    locale: 'zh-CN'
  }
});
```

### 用户照片

```typescript
// 获取用户照片元数据
const photoMeta = await orm.users.getPhotoMetadata('user-id');

// 获取用户照片内容
const photoBlob = await orm.users.getPhotoContent('user-id');
```

## 组管理

### 基本操作

```typescript
// 获取所有组
const groups = await orm.groups.findMany();

// 通过 ID 获取组
const group = await orm.groups.findById('group-id');

// 通过邮箱查找组
const groupByEmail = await orm.groups.findByEmail('group@example.com');

// 搜索组
const searchResults = await orm.groups.search('技术部');

// 创建组
const newGroup = await orm.groups.create({
  displayName: '技术部',
  mailNickname: 'techteam',
  mailEnabled: true,
  securityEnabled: true,
  groupTypes: ['Unified']
});

// 更新组
await orm.groups.update('group-id', {
  description: '技术部门团队'
});

// 删除组
await orm.groups.delete('group-id');
```

### 组成员管理

```typescript
// 获取组成员
const members = await orm.groups.members('group-id').findMany();

// 获取组所有者
const owners = await orm.groups.owners('group-id').findMany();

// 添加成员
await orm.groups.addMember('group-id', 'user-id');

// 移除成员
await orm.groups.removeMember('group-id', 'user-id');

// 添加所有者
await orm.groups.addOwner('group-id', 'user-id');

// 移除所有者
await orm.groups.removeOwner('group-id', 'user-id');

// 检查成员身份
const isMember = await orm.groups.isMember('group-id', 'user-id');
```

### 组资源

```typescript
// 获取组关联的团队
const team = await orm.groups.getTeam('group-id');

// 获取组驱动器
const drive = await orm.groups.getDrive('group-id');

// 获取组事件
const groupEvents = await orm.groups.events('group-id').findMany();

// 获取组日历事件
const calendarEvents = await orm.groups.calendarEvents('group-id').findMany();
```

## 文件操作

### 基本操作

```typescript
const files = orm.files;

// 获取用户根目录的文件
const rootItems = await files.getUserRootItems('user-id');

// 获取驱动器项（通过 ID）
const item = await files.getItem('user-id', 'item-id');

// 获取驱动器项（通过路径）
const itemByPath = await files.getItemByPath('user-id', '/Documents/report.xlsx');

// 列出文件夹的子项
const children = await files.listChildren('user-id', 'folder-id');
```

### 文件上传和下载

```typescript
// 创建文件夹
const folder = await files.createFolder('user-id', 'root', '工作文档');

// 上传小文件（< 4MB）
const uploadedFile = await files.uploadSmallFile(
  'user-id',
  '/Documents/report.pdf',
  fileBuffer,
  'replace'
);

// 下载文件
const downloadUrl = await files.downloadFile('user-id', 'file-id');

// 下载为特定格式
const pdfUrl = await files.downloadAsFormat('user-id', 'word-file-id', 'pdf');
```

### 文件管理

```typescript
// 更新文件元数据
await files.updateMetadata('user-id', 'item-id', {
  name: '新文件名.xlsx'
});

// 移动文件
await files.move('user-id', 'item-id', 'target-folder-id', '新名称.xlsx');

// 复制文件
const monitorUrl = await files.copy('user-id', 'item-id', 'target-folder-id');

// 删除文件（移到回收站）
await files.delete('user-id', 'item-id');

// 永久删除文件
await files.permanentDelete('user-id', 'item-id');
```

### 文件搜索

```typescript
// 搜索文件
const results = await files.search('user-id', '财务报告');

// 增量查询
const delta = await files.delta('user-id');
console.log('变化的文件:', delta.data);

// 下次查询使用 deltaLink
const nextDelta = await files.delta('user-id', delta.meta.deltaLink);
```

### 文件共享

```typescript
// 创建共享链接
const shareLink = await files.createShareLink(
  'user-id',
  'item-id',
  'view',
  'organization'
);
console.log('共享链接:', shareLink.link.webUrl);

// 邀请用户访问
await files.invite(
  'user-id',
  'item-id',
  ['colleague@example.com'],
  ['read'],
  true,
  '请查看这个文件'
);

// 列出权限
const permissions = await files.listPermissions('user-id', 'item-id');

// 删除权限
await files.deletePermission('user-id', 'item-id', 'permission-id');
```

### 其他操作

```typescript
// 获取缩略图
const thumbnails = await files.getThumbnails('user-id', 'item-id');

// 获取与我共享的文件
const sharedWithMe = await files.getSharedWithMe();

// 获取最近使用的文件
const recentFiles = await files.getRecentFiles();
```

## 邮件操作

> **💡 Repository vs Service**
> 
> - **Repository (`orm.users.messages()`)**: 用于 CRUD 操作和复杂查询，支持查询构建器
> - **Service (`orm.mail`)**: 用于业务逻辑操作，如发送邮件、管理文件夹

### 使用 Repository 查询邮件

```typescript
// 使用查询构建器进行复杂查询
const unreadMails = await orm.users.messages('user-id')
  .query()
  .where('isRead', 'eq', false)
  .orderBy('receivedDateTime', 'desc')
  .select(['id', 'subject', 'from', 'receivedDateTime'])
  .top(20)
  .get();

// 查询重要邮件
const importantMails = await orm.users.messages('user-id')
  .query()
  .where('importance', 'eq', 'high')
  .get();

// 查询带附件的邮件
const mailsWithAttachments = await orm.users.messages('user-id')
  .query()
  .where('hasAttachments', 'eq', true)
  .get();

// CRUD 操作
const message = await orm.users.messages('user-id').findById('message-id');
await orm.users.messages('user-id').update('message-id', { isRead: true });
await orm.users.messages('user-id').delete('message-id');

// 邮件特定操作
await orm.users.messages('user-id').reply('message-id', '收到，谢谢');
await orm.users.messages('user-id').markAsRead('message-id');
await orm.users.messages('user-id').markAsImportant('message-id');
```

### 使用 Service 进行业务操作

```typescript
const mail = orm.mail;

// 快捷访问常用文件夹
const inbox = await mail.getInbox('user-id', 20);
const sentItems = await mail.getSentItems('user-id', 20);
const drafts = await mail.getDrafts('user-id');
const folderMessages = await mail.getMessagesInFolder('user-id', 'folder-id');

// 获取单个邮件
const message = await mail.getMessage('user-id', 'message-id');
```

### 发送邮件

```typescript
// 发送邮件
await mail.send('user-id', {
  subject: '会议通知',
  body: {
    contentType: 'HTML',
    content: '<h1>下周一开会</h1><p>请准时参加。</p>'
  },
  toRecipients: [
    {
      emailAddress: {
        address: 'colleague@example.com',
        name: '同事'
      }
    }
  ]
});

// 创建草稿
const draft = await mail.createDraft('user-id', {
  subject: '草稿邮件',
  body: {
    contentType: 'Text',
    content: '这是一封草稿'
  }
});
```

### 邮件操作

```typescript
// 更新邮件（标记为已读）
await mail.updateMessage('user-id', 'message-id', {
  isRead: true
});

// 标记为已读
await mail.markAsRead('user-id', 'message-id');

// 标记为未读
await mail.markAsUnread('user-id', 'message-id');

// 标记为重要
await mail.markAsImportant('user-id', 'message-id');

// 删除邮件
await mail.deleteMessage('user-id', 'message-id');

// 移动邮件
await mail.moveMessage('user-id', 'message-id', 'destination-folder-id');

// 复制邮件
await mail.copyMessage('user-id', 'message-id', 'destination-folder-id');
```

### 回复和转发

```typescript
// 回复邮件
await mail.reply('user-id', 'message-id', '谢谢您的邮件');

// 全部回复
await mail.replyAll('user-id', 'message-id', '感谢各位');

// 转发邮件
await mail.forward('user-id', 'message-id', [
  {
    emailAddress: {
      address: 'another@example.com',
      name: '另一位同事'
    }
  }
], '请查看此邮件');
```

### 附件管理

```typescript
// 获取附件
const attachments = await mail.getAttachments('user-id', 'message-id');

// 添加附件
await mail.addAttachment('user-id', 'message-id', {
  '@odata.type': '#microsoft.graph.fileAttachment',
  name: 'document.pdf',
  contentBytes: base64Content
});

// 删除附件
await mail.deleteAttachment('user-id', 'message-id', 'attachment-id');
```

### 文件夹管理

```typescript
// 获取邮件文件夹
const folders = await mail.getMailFolders('user-id');

// 创建文件夹
const newFolder = await mail.createMailFolder('user-id', '工作邮件');

// 创建子文件夹
const subFolder = await mail.createMailFolder('user-id', '项目邮件', 'parent-folder-id');

// 删除文件夹
await mail.deleteMailFolder('user-id', 'folder-id');
```

### 搜索和增量查询

```typescript
// 搜索邮件
const searchResults = await mail.search('user-id', '项目会议', 10);

// 增量查询
const delta = await mail.delta('user-id');

// 下次查询
const nextDelta = await mail.delta('user-id', delta.meta.deltaLink);
```

## 日历操作

### 日历管理

```typescript
const calendar = orm.calendar;

// 获取用户主日历
const userCalendar = await calendar.getUserCalendar('user-id');

// 获取所有日历
const calendars = await calendar.getCalendars('user-id');

// 创建日历
const newCalendar = await calendar.createCalendar('user-id', {
  name: '工作日历'
});

// 获取日历组
const calendarGroups = await calendar.getCalendarGroups('user-id');
```

### 事件管理

```typescript
// 获取事件列表
const events = await calendar.getEvents('user-id', 20);

// 获取日历视图（指定时间范围）
const view = await calendar.getCalendarView(
  'user-id',
  '2024-01-01T00:00:00Z',
  '2024-01-31T23:59:59Z'
);

// 获取单个事件
const event = await calendar.getEvent('user-id', 'event-id');

// 创建事件
const newEvent = await calendar.createEvent('user-id', {
  subject: '团队会议',
  body: {
    contentType: 'HTML',
    content: '讨论项目进度'
  },
  start: {
    dateTime: '2024-01-15T10:00:00',
    timeZone: 'China Standard Time'
  },
  end: {
    dateTime: '2024-01-15T11:00:00',
    timeZone: 'China Standard Time'
  },
  location: {
    displayName: '会议室 A'
  },
  attendees: [
    {
      type: 'required',
      emailAddress: {
        address: 'colleague@example.com',
        name: '同事'
      }
    }
  ]
});

// 更新事件
await calendar.updateEvent('user-id', 'event-id', {
  subject: '更新的会议主题'
});

// 删除事件
await calendar.deleteEvent('user-id', 'event-id');
```

### 事件响应

```typescript
// 接受事件邀请
await calendar.acceptEvent('user-id', 'event-id', '我会参加');

// 暂时接受
await calendar.tentativelyAcceptEvent('user-id', 'event-id', '可能参加');

// 拒绝事件
await calendar.declineEvent('user-id', 'event-id', '时间冲突');

// 取消事件（组织者）
await calendar.cancelEvent('user-id', 'event-id', '会议取消');
```

### 会议室

```typescript
// 获取会议室列表（使用 places API）
const rooms = await calendar.getRooms();

// 获取会议室列表
const roomLists = await calendar.getRoomLists();

// 查找用户可访问的会议室
const userRooms = await calendar.findRooms('user-id');

// 查找会议室列表
const userRoomLists = await calendar.findRoomLists('user-id');

// 查找指定列表下的会议室
const roomsInList = await calendar.findRoomsInList(
  'user-id',
  'roomlist@contoso.com'
);
```

### 忙/闲和查找会议时间

```typescript
// 获取忙/闲时间表
const schedule = await calendar.getSchedule(
  ['user1@contoso.com', 'user2@contoso.com', 'room@contoso.com'],
  {
    dateTime: '2024-01-15T09:00:00',
    timeZone: 'China Standard Time'
  },
  {
    dateTime: '2024-01-15T18:00:00',
    timeZone: 'China Standard Time'
  },
  30
);

// 查找会议时间
const meetingTimes = await calendar.findMeetingTimes(
  [
    {
      type: 'required',
      emailAddress: {
        address: 'colleague@example.com'
      }
    }
  ],
  {
    timeslots: [
      {
        start: {
          dateTime: '2024-01-15T09:00:00',
          timeZone: 'China Standard Time'
        },
        end: {
          dateTime: '2024-01-15T18:00:00',
          timeZone: 'China Standard Time'
        }
      }
    ]
  },
  'PT1H',
  10
);
```

### 其他操作

```typescript
// 增量查询
const delta = await calendar.deltaEvents('user-id');

// 获取循环事件的实例
const instances = await calendar.getEventInstances(
  'user-id',
  'series-master-id',
  '2024-01-01T00:00:00Z',
  '2024-01-31T23:59:59Z'
);

// 获取事件附件
const eventAttachments = await calendar.getEventAttachments('user-id', 'event-id');

// 添加事件附件
await calendar.addEventAttachment('user-id', 'event-id', attachmentData);
```

## 高级查询

### 链式查询

```typescript
// 复杂查询
const users = await orm.users
  .query()
  .where('department', 'eq', 'IT')
  .and('city', 'Beijing')
  .select(['id', 'displayName', 'mail', 'jobTitle'])
  .orderBy('displayName', 'asc')
  .top(10)
  .get();

// 使用高级操作符
const searchUsers = await orm.users
  .query()
  .where('displayName', 'startswith', '张')
  .or('mail', 'contains', 'contoso.com')
  .get();

// 搜索
const results = await orm.users
  .query()
  .search('displayName', '工程师')
  .top(20)
  .get();
```

### 分页迭代

```typescript
// 自动处理分页
for await (const user of orm.users.query().pagination()) {
  console.log(user.displayName);
}

// 手动分页
let nextLink: string | undefined;
do {
  const page = await orm.users
    .query()
    .top(100)
    .skipToken(nextLink)
    .get();
  
  page.data.forEach(user => {
    console.log(user.displayName);
  });
  
  nextLink = page.meta.nextLink;
} while (nextLink);
```

## 增量查询

```typescript
// 首次查询
const initialDelta = await orm.deltaUsers();
console.log('初始用户数:', initialDelta.data.length);

// 保存 deltaLink
const deltaLink = initialDelta.meta.deltaLink;

// 后续查询（只返回变化的数据）
const changes = await orm.deltaUsers(deltaLink);
console.log('变化的用户:', changes.data);

// 其他资源的增量查询
const groupDelta = await orm.deltaGroups();
const messageDelta = await orm.deltaMessages('user-id');
const eventDelta = await orm.deltaEvents('user-id');
const fileDelta = await orm.deltaDriveItems('user-id');
```

## 批处理

```typescript
// 执行批处理请求（最多 20 个）
const batchResponse = await orm.batch([
  {
    id: '1',
    method: 'GET',
    url: '/users/user1@contoso.com'
  },
  {
    id: '2',
    method: 'GET',
    url: '/users/user2@contoso.com'
  },
  {
    id: '3',
    method: 'PATCH',
    url: '/users/user3@contoso.com',
    body: {
      jobTitle: '高级工程师'
    },
    headers: {
      'Content-Type': 'application/json'
    }
  }
]);

// 处理响应
batchResponse.responses.forEach(response => {
  console.log(`请求 ${response.id}: 状态 ${response.status}`);
  console.log('响应:', response.body);
});
```

## 调试和监控

### 请求调试

```typescript
// 获取查询构建器的调试信息
const query = orm.users.query()
  .where('department', 'eq', 'IT')
  .select(['id', 'displayName', 'mail'])
  .top(10);

// 执行前获取调试信息
const debugInfo = query.getDebug();
console.log('请求方法:', debugInfo.method);
console.log('请求 URL:', debugInfo.url);
console.log('查询参数:', debugInfo.queryParams);
console.log('请求头:', debugInfo.headers);

// 获取原始请求参数（不执行请求）
const rawParams = query.getRaw();
console.log('原始参数:', rawParams);

// 执行请求
const users = await query.get();

// 执行后获取最后请求的调试信息
const lastDebugInfo = query.getDebug();
console.log('最后请求信息:', lastDebugInfo);
```

### 服务层调试

```typescript
// 邮件服务调试
await orm.mail.send('user-id', message);
const mailDebugInfo = orm.mail.getDebug();
console.log('邮件发送调试信息:', mailDebugInfo);

// 文件服务调试
await orm.files.uploadSmallFile('user-id', path, buffer);
const fileDebugInfo = orm.files.getDebug();
console.log('文件上传调试信息:', fileDebugInfo);

// 日历服务调试
await orm.calendar.createEvent('user-id', event);
const calendarDebugInfo = orm.calendar.getDebug();
console.log('日历事件调试信息:', calendarDebugInfo);
```

### 自定义 Headers

```typescript
// 为查询添加自定义 headers
const users = await orm.users.query()
  .header('ConsistencyLevel', 'eventual')
  .header('Prefer', 'outlook.timezone="Asia/Shanghai"')
  .where('displayName', 'startswith', '张')
  .get();

// 为服务调用添加自定义 headers
await orm.mail.send('user-id', message, {
  'X-Custom-Header': 'value',
  'ConsistencyLevel': 'eventual'
});

// 检查应用的 headers
const debugInfo = orm.mail.getDebug();
console.log('应用的 headers:', debugInfo.headers);
```

## 完整示例

### 用户入职流程

```typescript
async function onboardUser(email: string, displayName: string) {
  // 1. 创建用户
  const user = await orm.users.create({
    displayName,
    userPrincipalName: email,
    mailNickname: email.split('@')[0],
    accountEnabled: true,
    passwordProfile: {
      password: 'TempPass123!',
      forceChangePasswordNextSignIn: true
    }
  });

  // 2. 分配许可证
  await orm.users.assignLicense(user.id, [
    { skuId: 'office365-sku-id' }
  ], []);

  // 3. 添加到组
  await orm.groups.addMember('new-employees-group-id', user.id);

  // 4. 设置邮箱设置
  await orm.users.updateMailboxSettings(user.id, {
    timeZone: 'China Standard Time',
    language: { locale: 'zh-CN' }
  });

  // 5. 发送欢迎邮件
  await orm.mail.send('admin@contoso.com', {
    subject: '欢迎加入公司',
    toRecipients: [
      {
        emailAddress: {
          address: email,
          name: displayName
        }
      }
    ],
    body: {
      contentType: 'HTML',
      content: '<h1>欢迎！</h1><p>请查看附件中的入职指南。</p>'
    }
  });

  console.log('用户入职完成:', user.id);
  return user;
}
```

### 组织会议

```typescript
async function scheduleMeeting(
  organizerId: string,
  subject: string,
  attendeeEmails: string[],
  duration: number
) {
  // 1. 查找合适的会议时间
  const attendees = attendeeEmails.map(email => ({
    type: 'required',
    emailAddress: { address: email }
  }));

  const suggestions = await orm.calendar.findMeetingTimes(
    attendees,
    {
      timeslots: [
        {
          start: {
            dateTime: '2024-01-15T09:00:00',
            timeZone: 'China Standard Time'
          },
          end: {
            dateTime: '2024-01-15T18:00:00',
            timeZone: 'China Standard Time'
          }
        }
      ]
    },
    `PT${duration}M`,
    5
  );

  // 2. 查找可用会议室
  const rooms = await orm.calendar.findRooms(organizerId);

  // 3. 创建事件
  const event = await orm.calendar.createEvent(organizerId, {
    subject,
    start: suggestions[0].meetingTimeSlot.start,
    end: suggestions[0].meetingTimeSlot.end,
    location: {
      displayName: rooms[0]?.displayName || '线上会议'
    },
    attendees
  });

  console.log('会议已创建:', event.id);
  return event;
}
```

这些示例展示了新版 Graph ORM 的主要功能和最佳实践。更多详细信息，请查看 [完整文档](./DOCS_INDEX.md)。

