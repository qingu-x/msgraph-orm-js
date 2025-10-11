# Microsoft Graph ORM 使用示例

本文档提供了各种常见场景的完整示例代码。

## 目录

- [初始化](#初始化)
- [用户管理](#用户管理)
- [日历管理](#日历管理)
- [邮件管理](#邮件管理)
- [文件管理](#文件管理)
- [Teams 协作](#teams-协作)
- [SharePoint](#sharepoint)
- [报告和分析](#报告和分析)
- [高级查询](#高级查询)

## 初始化

### 使用客户端凭据认证

```typescript
import { Client } from '@microsoft/microsoft-graph-client';
import { ClientCredentialsAuthProvider, createGraphORM } from '@qingu-x/msgraph-orm-js';

// 创建认证提供者
const authProvider = new ClientCredentialsAuthProvider({
  tenantId: 'your-tenant-id',
  clientId: 'your-client-id',
  clientSecret: 'your-client-secret',
  // 可选：指定国家云
  // cloudEndpoint: 'https://microsoftgraph.chinacloudapi.cn' // 中国版
});

// 创建 Graph 客户端
const client = Client.initWithMiddleware({
  authProvider
});

// 创建 ORM 实例
const orm = createGraphORM(client);
```

### 使用委托权限认证

```typescript
import { Client } from '@microsoft/microsoft-graph-client';
import { createGraphORM } from '@qingu-x/msgraph-orm-js';

// 使用已有的访问令牌
const client = Client.init({
  authProvider: (done) => {
    done(null, accessToken);
  }
});

const orm = createGraphORM(client);
```

## 用户管理

### 获取所有用户

```typescript
// 获取前 10 个用户
const users = await orm.users
  .query()
  .select(['id', 'displayName', 'mail'])
  .top(10)
  .execute();

console.log('用户列表:', users.data);
```

### 查询特定用户

```typescript
// 按条件查询
const managers = await orm.users
  .query()
  .where('jobTitle', 'contains', '经理')
  .and('accountEnabled', 'eq', true)
  .orderBy('displayName', 'asc')
  .execute();

// 获取单个用户
const user = await orm.users.findById('user-id');
console.log('用户信息:', user.displayName, user.mail);
```

### 创建用户

```typescript
const newUser = await orm.users.create({
  accountEnabled: true,
  displayName: '张三',
  mailNickname: 'zhangsan',
  userPrincipalName: 'zhangsan@contoso.com',
  passwordProfile: {
    forceChangePasswordNextSignIn: true,
    password: 'TempPassword123!'
  }
});

console.log('创建的用户 ID:', newUser.id);
```

### 更新用户

```typescript
const updatedUser = await orm.users.update('user-id', {
  jobTitle: '高级工程师',
  department: '技术部'
});
```

### 管理用户许可证

```typescript
// 获取用户许可证
const licenses = await orm.getUserLicenses('user-id');

// 为用户分配许可证
await orm.assignUserLicense('user-id', [
  {
    skuId: 'license-sku-id',
    disabledPlans: [] // 可选：禁用某些服务计划
  }
], []);

// 获取组织的所有 SKU
const skus = await orm.getSubscribedSkus();
console.log('可用许可证:', skus);
```

### 组管理

```typescript
// 获取所有组
const groups = await orm.groups.query().execute();

// 获取组成员
const members = await orm.groupMembers('group-id').query().execute();

// 获取组所有者
const owners = await orm.groupOwners('group-id').query().execute();

// 创建 Microsoft 365 组
const newGroup = await orm.groups.create({
  displayName: '项目团队',
  description: '项目协作组',
  mailNickname: 'projectteam',
  mailEnabled: true,
  securityEnabled: false,
  groupTypes: ['Unified'] // Microsoft 365 组
});
```

## 日历管理

### 获取用户日历事件

```typescript
// 获取用户的所有事件
const events = await orm.events('user@contoso.com')
  .query()
  .select(['subject', 'start', 'end', 'location'])
  .orderBy('start/dateTime', 'asc')
  .top(20)
  .execute();

console.log('日历事件:', events.data);
```

### 获取日历视图（指定时间范围）

```typescript
const now = new Date();
const nextWeek = new Date(now.getTime() + 7 * 24 * 60 * 60 * 1000);

const upcomingEvents = await orm.getCalendarView(
  'user@contoso.com',
  now.toISOString(),
  nextWeek.toISOString()
);

console.log('未来一周的事件:', upcomingEvents);
```

### 创建日历事件

```typescript
const newEvent = await orm.events('user@contoso.com').create({
  subject: '项目会议',
  body: {
    contentType: 'HTML',
    content: '讨论项目进展和下一步计划'
  },
  start: {
    dateTime: '2025-10-15T10:00:00',
    timeZone: 'China Standard Time'
  },
  end: {
    dateTime: '2025-10-15T11:00:00',
    timeZone: 'China Standard Time'
  },
  location: {
    displayName: '会议室 A'
  },
  attendees: [
    {
      emailAddress: {
        address: 'colleague@contoso.com',
        name: '李四'
      },
      type: 'required'
    }
  ],
  isOnlineMeeting: true,
  onlineMeetingProvider: 'teamsForBusiness'
});

console.log('创建的事件 ID:', newEvent.id);
console.log('在线会议链接:', newEvent.onlineMeetingUrl);
```

### 查询会议室可用性

```typescript
// 查找所有会议室
const rooms = await orm.findRooms('user@contoso.com');
console.log('可用会议室:', rooms);

// 获取忙/闲信息
const scheduleInfo = await orm.getSchedule(
  ['room1@contoso.com', 'room2@contoso.com'],
  {
    dateTime: '2025-10-15T09:00:00',
    timeZone: 'China Standard Time'
  },
  {
    dateTime: '2025-10-15T17:00:00',
    timeZone: 'China Standard Time'
  },
  30 // 30分钟间隔
);

console.log('会议室忙闲状态:', scheduleInfo);
```

## 邮件管理

### 读取收件箱

```typescript
// 获取收件箱邮件
const messages = await orm.inbox('user@contoso.com')
  .query()
  .select(['subject', 'from', 'receivedDateTime', 'isRead'])
  .orderBy('receivedDateTime', 'desc')
  .top(10)
  .execute();

console.log('收件箱邮件:', messages.data);

// 只获取未读邮件
const unreadMessages = await orm.inbox('user@contoso.com')
  .query()
  .where('isRead', 'eq', false)
  .execute();
```

### 发送邮件

```typescript
await orm.sendMail('user@contoso.com', {
  subject: '测试邮件',
  body: {
    contentType: 'HTML',
    content: '<h1>你好</h1><p>这是一封测试邮件。</p>'
  },
  toRecipients: [
    {
      emailAddress: {
        address: 'recipient@contoso.com',
        name: '收件人'
      }
    }
  ],
  ccRecipients: [
    {
      emailAddress: {
        address: 'cc@contoso.com'
      }
    }
  ]
}, true); // 保存到已发送项目

console.log('邮件发送成功');
```

### 管理邮件附件

```typescript
// 获取邮件附件
const attachments = await orm.getMessageAttachments('user@contoso.com', 'message-id');

// 添加附件
await orm.addMessageAttachment('user@contoso.com', 'message-id', {
  '@odata.type': '#microsoft.graph.fileAttachment',
  name: 'document.pdf',
  contentType: 'application/pdf',
  contentBytes: base64Content
});
```

### 设置邮箱规则

```typescript
// 创建邮件规则
const rule = await orm.messageRules('user@contoso.com').create({
  displayName: '重要邮件标记',
  sequence: 1,
  isEnabled: true,
  conditions: {
    fromAddresses: [
      {
        emailAddress: {
          address: 'boss@contoso.com'
        }
      }
    ],
    importance: 'high'
  },
  actions: {
    markImportance: 'high',
    markAsRead: false
  }
});
```

### 设置自动回复

```typescript
const settings = await orm.updateMailboxSettings('user@contoso.com', {
  automaticRepliesSetting: {
    status: 'scheduled',
    externalAudience: 'all',
    internalReplyMessage: '我正在休假，将在下周一回复。',
    externalReplyMessage: 'I am out of office. Will reply next Monday.',
    scheduledStartDateTime: {
      dateTime: '2025-10-20T00:00:00',
      timeZone: 'China Standard Time'
    },
    scheduledEndDateTime: {
      dateTime: '2025-10-27T23:59:59',
      timeZone: 'China Standard Time'
    }
  }
});
```

## 文件管理

### 列出文件

```typescript
// 获取根目录文件
const files = await orm.driveItems('user@contoso.com')
  .query()
  .select(['name', 'size', 'createdDateTime', 'webUrl'])
  .orderBy('lastModifiedDateTime', 'desc')
  .execute();

console.log('文件列表:', files.data);

// 获取指定路径的文件
const documents = await orm.driveItemsByPath('user@contoso.com', '/Documents')
  .query()
  .execute();
```

### 搜索文件

```typescript
const searchResults = await orm.searchDriveItems('user@contoso.com', '财务报告');
console.log('搜索结果:', searchResults);

// 使用 Microsoft Search API（更强大）
const results = await orm.searchFiles('季度报告', 25);
console.log('搜索结果:', results);
```

### 创建文件夹

```typescript
const newFolder = await orm.createFolder(
  'user@contoso.com',
  'root', // 父文件夹 ID
  '2025年项目',
  'rename' // 如果存在则重命名
);

console.log('创建的文件夹:', newFolder.name);
```

### 上传文件

```typescript
// 上传小文件（< 4MB）
const uploadedFile = await orm.uploadSmallFile(
  'user@contoso.com',
  '/Documents/report.pdf',
  fileBuffer,
  'rename'
);

console.log('上传成功:', uploadedFile.webUrl);
```

### 下载文件

```typescript
const downloadUrl = await orm.downloadFile('user@contoso.com', 'file-id');
console.log('下载链接:', downloadUrl);

// 使用 fetch 下载
const response = await fetch(downloadUrl);
const blob = await response.blob();
```

### 文件操作

```typescript
// 移动文件
await orm.moveDriveItem(
  'user@contoso.com',
  'file-id',
  'target-folder-id',
  '新文件名.pdf' // 可选
);

// 复制文件
const monitorUrl = await orm.copyDriveItem(
  'user@contoso.com',
  'file-id',
  'target-folder-id'
);
console.log('复制操作监控链接:', monitorUrl);

// 删除文件
await orm.deleteDriveItem('user@contoso.com', 'file-id');
```

### 共享文件

```typescript
// 创建共享链接
const shareLink = await orm.createDriveItemLink(
  'user@contoso.com',
  'file-id',
  'view', // 'view' | 'edit' | 'embed'
  'organization' // 'anonymous' | 'organization'
);

console.log('共享链接:', shareLink.link.webUrl);

// 邀请用户访问
await orm.inviteToDriveItem(
  'user@contoso.com',
  'file-id',
  ['colleague@contoso.com'],
  ['write'],
  true, // 发送邀请邮件
  '请查看这个文件'
);
```

### 监控文件变化（Delta Query）

```typescript
// 首次查询
let result = await orm.deltaDriveItems('user@contoso.com');
console.log('初始文件:', result.data);

// 保存 deltaLink
const deltaLink = result.meta.deltaLink;

// 后续查询（只返回变化）
result = await orm.deltaDriveItems('user@contoso.com', deltaLink);
console.log('变化的文件:', result.data);
```

## Teams 协作

### 获取团队信息

```typescript
// 获取所有团队
const teams = await orm.teams.query().execute();

// 通过组 ID 获取团队
const team = await orm.getTeamByGroup('group-id');

// 获取频道
const channels = await orm.channels('team-id').query().execute();
```

### 发送频道消息

```typescript
const message = await orm.sendChannelMessage('team-id', 'channel-id', {
  body: {
    contentType: 'html',
    content: '<h1>重要通知</h1><p>项目已成功交付！</p>'
  }
});

console.log('消息 ID:', message.id);
```

### 获取聊天消息

```typescript
// 获取频道消息
const messages = await orm.channelMessages('team-id', 'channel-id')
  .query()
  .orderBy('createdDateTime', 'desc')
  .top(50)
  .execute();

// 获取消息回复
const replies = await orm.channelMessageReplies('team-id', 'channel-id', 'message-id')
  .query()
  .execute();
```

### 管理团队成员

```typescript
// 获取团队成员
const members = await orm.teamMembers('team-id').query().execute();

// 添加成员
await orm.teamMembers('team-id').create({
  '@odata.type': '#microsoft.graph.aadUserConversationMember',
  roles: ['owner'], // 或 ['member']
  'user@odata.bind': `https://graph.microsoft.com/v1.0/users('user-id')`
});
```

## SharePoint

### 获取站点

```typescript
// 获取根站点
const rootSite = await orm.rootSite;

// 通过路径获取站点
const site = await orm.getSiteByPath(
  'contoso.sharepoint.com',
  '/sites/projectsite'
);

console.log('站点 ID:', site.id);
```

### 列表操作

```typescript
// 获取站点列表
const lists = await orm.siteLists('site-id').query().execute();

// 获取列表项
const items = await orm.listItems('site-id', 'list-id')
  .query()
  .expand('fields')
  .execute();

// 创建列表项
const newItem = await orm.listItems('site-id', 'list-id').create({
  fields: {
    Title: '新项目',
    Description: '项目描述',
    Status: '进行中'
  }
});
```

### 列定义

```typescript
// 获取列定义
const columns = await orm.listColumns('site-id', 'list-id').query().execute();

// 创建新列
const newColumn = await orm.listColumns('site-id', 'list-id').create({
  name: 'Priority',
  displayName: '优先级',
  choice: {
    choices: ['高', '中', '低'],
    allowTextEntry: false,
    displayAs: 'dropDownMenu'
  }
});
```

## 报告和分析

### 使用报告

```typescript
// 获取活跃用户报告（CSV 格式）
const activeUsers = await orm.getOffice365ActiveUserDetail('D30');
console.log('30天活跃用户报告:', activeUsers);

// 邮件活动报告
const emailActivity = await orm.getEmailActivityUserDetail('D7');

// Teams 使用报告
const teamsUsage = await orm.getTeamsUserActivityUserDetail('D30');
```

### 审计日志

```typescript
// 获取目录审计日志
const audits = await orm.directoryAudits
  .query()
  .orderBy('activityDateTime', 'desc')
  .top(100)
  .execute();

// 获取登录日志
const signIns = await orm.signInLogs
  .query()
  .where('createdDateTime', 'ge', '2025-10-01T00:00:00Z')
  .and('status/errorCode', 'eq', 0) // 成功登录
  .execute();
```

### 洞察分析

```typescript
// 获取趋势文件
const trending = await orm.trendingInsights('user@contoso.com')
  .query()
  .execute();

// 获取最近使用的文件
const used = await orm.usedInsights('user@contoso.com')
  .query()
  .execute();

// 获取共享的文件
const shared = await orm.sharedInsights('user@contoso.com')
  .query()
  .execute();
```

### 人员分析

```typescript
// 获取相关人员
const people = await orm.people('user@contoso.com')
  .query()
  .top(10)
  .execute();

console.log('相关人员:', people.data);
```

## 高级查询

### 复杂查询条件

```typescript
// 多条件查询
const results = await orm.users
  .query()
  .where('department', 'eq', '技术部')
  .and('jobTitle', 'contains', '工程师')
  .or('jobTitle', 'contains', '架构师')
  .orderBy('displayName', 'asc')
  .select(['id', 'displayName', 'jobTitle', 'mail'])
  .top(50)
  .execute();
```

### 展开关联数据

```typescript
// 展开用户的管理者信息
const usersWithManager = await orm.users
  .query()
  .expand('manager')
  .select(['displayName', 'mail', 'manager'])
  .execute();
```

### 搜索和筛选

```typescript
// 使用 $search
const searchResults = await orm.users
  .query()
  .search('displayName', '张')
  .execute();
```

### 分页查询

```typescript
// 使用分页
let page = 0;
let hasMore = true;

while (hasMore) {
  const result = await orm.users
    .query()
    .top(100)
    .skip(page * 100)
    .execute();
  
  console.log(`第 ${page + 1} 页，共 ${result.data.length} 条记录`);
  
  hasMore = !!result.meta.nextLink;
  page++;
}
```

### 批处理请求

```typescript
// 批量执行多个请求
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
    method: 'GET',
    url: '/groups'
  }
]);

// 处理响应
batchResponse.responses.forEach(response => {
  if (response.status === 200) {
    console.log(`请求 ${response.id} 成功:`, response.body);
  } else {
    console.error(`请求 ${response.id} 失败:`, response.status);
  }
});
```

### Delta Query（变化跟踪）

```typescript
// 跟踪用户变化
let deltaLink: string | undefined;

// 首次查询
const initialResult = await orm.deltaUsers();
console.log('初始用户数:', initialResult.data.length);
deltaLink = initialResult.meta.deltaLink;

// 定期查询变化
setInterval(async () => {
  const changes = await orm.deltaUsers(deltaLink);
  
  if (changes.data.length > 0) {
    console.log('检测到用户变化:', changes.data.length);
    
    // 处理变化
    changes.data.forEach(user => {
      if (user['@removed']) {
        console.log('用户被删除:', user.id);
      } else {
        console.log('用户更新:', user.displayName);
      }
    });
  }
  
  // 更新 deltaLink
  deltaLink = changes.meta.deltaLink;
}, 60000); // 每分钟检查一次
```

### Webhooks（订阅通知）

```typescript
// 创建订阅
const subscription = await orm.createSubscription({
  changeType: 'created,updated',
  notificationUrl: 'https://your-app.com/webhooks/graph',
  resource: 'users',
  expirationDateTime: new Date(Date.now() + 3600000).toISOString(), // 1小时后过期
  clientState: 'secret-state-value-12345'
});

console.log('订阅 ID:', subscription.id);

// 续订
await orm.renewSubscription(
  subscription.id,
  new Date(Date.now() + 3600000).toISOString()
);

// 删除订阅
await orm.deleteSubscription(subscription.id);
```

## 错误处理

```typescript
import { GraphOrmError, GraphErrorCode } from '@qingu-x/msgraph-orm-js';

try {
  const user = await orm.users.findById('invalid-id');
} catch (error) {
  if (error instanceof GraphOrmError) {
    console.error('错误代码:', error.code);
    console.error('错误消息:', error.message);
    console.error('HTTP 状态:', error.statusCode);
    console.error('请求 ID:', error.requestId);
    
    // 根据错误代码处理
    switch (error.code) {
      case GraphErrorCode.RESOURCE_NOT_FOUND:
        console.log('资源未找到');
        break;
      case GraphErrorCode.UNAUTHORIZED:
        console.log('未授权，请检查权限');
        break;
      case GraphErrorCode.THROTTLED:
        console.log('请求被限流，请稍后重试');
        break;
    }
  }
}
```

## 最佳实践

### 1. 使用 Select 减少数据传输

```typescript
// ❌ 不好：获取所有字段
const users = await orm.users.query().execute();

// ✅ 好：只获取需要的字段
const users = await orm.users
  .query()
  .select(['id', 'displayName', 'mail'])
  .execute();
```

### 2. 使用分页避免大量数据

```typescript
// ✅ 使用 top 限制结果数量
const users = await orm.users.query().top(100).execute();
```

### 3. 使用 Delta Query 跟踪变化

```typescript
// ✅ 使用 Delta Query 而不是每次全量查询
const changes = await orm.deltaUsers(savedDeltaLink);
```

### 4. 使用批处理减少请求次数

```typescript
// ✅ 批量请求多个资源
const batch = await orm.batch([/* 多个请求 */]);
```

### 5. 合理处理国家云差异

```typescript
// ✅ 对不可用的 API 进行降级处理
try {
  const notebooks = await orm.notebooks('user@contoso.com').query().execute();
} catch (error) {
  if (error.code === GraphErrorCode.RESOURCE_NOT_FOUND) {
    console.log('OneNote 在当前云环境不可用');
    // 使用替代方案
  }
}
```

## 更多资源

- [完整 API 参考](./GRAPH_API_REFERENCE.md)
- [ORM 使用指南](./GRAPH_ORM_GUIDE.md)
- [API 覆盖率](./API_COVERAGE.md)
- [Microsoft Graph 官方文档](https://learn.microsoft.com/graph/)
