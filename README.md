# @qingu-x/msgraph-orm-js

[![npm version](https://img.shields.io/npm/v/@qingu-x/msgraph-orm-js.svg)](https://www.npmjs.com/package/@qingu-x/msgraph-orm-js)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

Microsoft Graph 的 ORM 封装，提供简洁、类型安全、符合 ORM 规范的 API 调用方式。

> ⚠️ 本项目包括文档大部分由 AI 完成，如遇到问题请提 [Issue](https://github.com/qingu-x/msgraph-orm-js/issues)

## ✨ 特性

- 🏗️ **标准 ORM 架构**: 仓储模式 + 服务层，职责清晰
- 🎯 **类型安全**: 100% TypeScript，完整的类型推导
- 🔗 **关系映射**: 支持实体关系的访问和管理
- 🌏 **国家云支持**: 全球版/中国版/美国政府版
- 🚀 **95%+ API 覆盖**: 覆盖 Graph v1.0 绝大多数常用接口
- 🔍 **强大查询构建器**: 支持复杂查询、分页、排序
- ⚡ **高级功能**: Delta Query、批处理、Webhooks
- 📦 **模块化设计**: 按需使用仓储和服务

## 🆕 新版本亮点

### 清晰的三层架构

```text
┌─────────────────────────────────────┐
│         GraphORM (入口层)            │
│    统一访问入口，管理仓储和服务         │
└─────────────────────────────────────┘
          │                    │
          ▼                    ▼
┌───────────────────┐  ┌───────────────────┐
│ Repositories      │  │  Services         │
│ - UserRepository  │  │  - FileService    │
│ - GroupRepository │  │  - MailService    │
│ - DeviceRepository│  │  - CalendarService│
│ ...               │  │  ...              │
└───────────────────┘  └───────────────────┘
          │                    │
          └────────┬───────────┘
                   ▼
         ┌──────────────────┐
         │  QueryBuilder    │
         │  查询构建器        │
         └──────────────────┘
```

### 对比旧版本

| 特性 | 旧版本 | 新版本 |
|------|--------|--------|
| GraphORM 方法数 | 185+ | ~30 |
| 架构设计 | 混乱，职责不清 | 清晰的仓储+服务 |
| API 一致性 | 混合多种风格 | 统一的访问模式 |
| 关系访问 | 手动拼接路径 | 类型安全的关系方法 |
| 代码维护性 | 低 | 高 |

## 📦 安装

```bash
npm install @qingu-x/msgraph-orm-js @microsoft/microsoft-graph-client
```

## 🚀 快速开始

```typescript
import { GraphClient, defaultEndpoints, createGraphORM } from '@qingu-x/msgraph-orm-js';

// 1. 创建 Graph 客户端
const graphClient = new GraphClient(
  {
    tenantId: 'your-tenant-id',
    clientId: 'your-client-id',
    clientSecret: 'your-client-secret'
  },
  defaultEndpoints.global  // 全球版端点
);

// 2. 获取底层 Graph Client
const client = graphClient.getGraphClient();

// 3. 创建 ORM 实例
const orm = createGraphORM(client);

// 4. 开始使用
// 用户管理（仓储）
const users = await orm.users
  .query()
  .where('department', 'eq', 'IT')
  .select(['displayName', 'mail'])
  .top(10)
  .get();

// 文件操作（服务）
await orm.files.uploadSmallFile('user-id', '/path/file.pdf', buffer);

// 邮件操作（服务）
await orm.mail.send('user-id', {
  subject: '测试邮件',
  toRecipients: [{ emailAddress: { address: 'user@example.com' } }],
  body: { contentType: 'Text', content: '邮件内容' }
});
```

## 📚 核心概念

### 1. 仓储 (Repositories)

仓储负责实体的 CRUD 操作和关系访问：

```typescript
// 用户仓储
const user = await orm.users.findById('user-id');
const userByEmail = await orm.users.findByEmail('user@example.com');
const exists = await orm.users.exists('user-id');
const count = await orm.users.count({ department: 'IT' });

// 关系访问
const userGroups = await orm.users.groups('user-id').findMany();
const userMessages = await orm.users.messages('user-id').findMany();
const manager = await orm.users.getManager('user-id');

// 组仓储
const group = await orm.groups.findById('group-id');
const members = await orm.groups.members('group-id').findMany();
await orm.groups.addMember('group-id', 'user-id');
```

### 2. 服务 (Services)

服务负责特定领域的业务逻辑：

```typescript
// 文件服务
const files = orm.files;
await files.uploadSmallFile(userId, path, content);
await files.createFolder(userId, parentId, 'Documents');
const items = await files.search(userId, 'report');
await files.createShareLink(userId, itemId, 'view', 'organization');

// 邮件服务
const mail = orm.mail;
await mail.send(userId, message);
const inbox = await mail.getInbox(userId);
await mail.reply(userId, messageId, '谢谢');
await mail.markAsRead(userId, messageId);

// 日历服务
const calendar = orm.calendar;
await calendar.createEvent(userId, event);
const view = await calendar.getCalendarView(userId, start, end);
await calendar.acceptEvent(userId, eventId);
const rooms = await calendar.findRooms(userId);
```

### 3. 查询构建器

强大的链式查询：

```typescript
const results = await orm.users
  .query()
  .where('department', 'eq', 'IT')
  .and('city', 'Beijing')
  .or('jobTitle', 'contains', 'Manager')
  .select(['id', 'displayName', 'mail', 'jobTitle'])
  .orderBy('displayName', 'asc')
  .top(20)
  .get();

// 分页迭代
for await (const user of orm.users.query().pagination()) {
  console.log(user.displayName);
}

// 搜索
const searchResults = await orm.users
  .query()
  .search('displayName', 'John')
  .get();
```

## 🌟 常用场景

### 用户管理

```typescript
// 查询和过滤
const itUsers = await orm.users
  .query()
  .where('department', 'eq', 'IT')
  .orderBy('displayName')
  .get();

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
await orm.users.update('user-id', {
  jobTitle: '高级工程师',
  department: '技术部'
});

// 许可证管理
await orm.users.assignLicense('user-id', [
  { skuId: 'sku-id' }
], []);

// 用户关系
const groups = await orm.users.groups('user-id').findMany();
const messages = await orm.users.messages('user-id').findMany();
const events = await orm.users.events('user-id').findMany();
```

### 组管理

```typescript
// 创建组
const group = await orm.groups.create({
  displayName: '技术部',
  mailNickname: 'techteam',
  mailEnabled: true,
  securityEnabled: true,
  groupTypes: ['Unified']
});

// 成员管理
await orm.groups.addMember('group-id', 'user-id');
await orm.groups.removeMember('group-id', 'user-id');
const isMember = await orm.groups.isMember('group-id', 'user-id');

// 获取成员和所有者
const members = await orm.groups.members('group-id').findMany();
const owners = await orm.groups.owners('group-id').findMany();
```

### 文件操作

```typescript
// 上传和下载
await orm.files.uploadSmallFile('user-id', '/Documents/report.pdf', buffer);
const downloadUrl = await orm.files.downloadFile('user-id', 'file-id');

// 文件夹管理
await orm.files.createFolder('user-id', 'root', '工作文档');
const children = await orm.files.listChildren('user-id', 'folder-id');

// 文件操作
await orm.files.move('user-id', 'item-id', 'target-folder-id');
await orm.files.copy('user-id', 'item-id', 'target-folder-id');

// 共享和权限
const link = await orm.files.createShareLink('user-id', 'item-id', 'view', 'organization');
await orm.files.invite('user-id', 'item-id', ['user@example.com'], ['read']);

// 搜索
const results = await orm.files.search('user-id', '财务报告');
```

### 邮件操作

```typescript
// 发送邮件
await orm.mail.send('user-id', {
  subject: '会议通知',
  toRecipients: [
    { emailAddress: { address: 'colleague@example.com', name: '同事' } }
  ],
  body: {
    contentType: 'HTML',
    content: '<h1>会议通知</h1><p>明天下午2点开会</p>'
  }
});

// 读取邮件
const inbox = await orm.mail.getInbox('user-id', 20);
const message = await orm.mail.getMessage('user-id', 'message-id');

// 邮件操作
await orm.mail.reply('user-id', 'message-id', '收到');
await orm.mail.forward('user-id', 'message-id', recipients, '请查看');
await orm.mail.markAsRead('user-id', 'message-id');

// 附件管理
const attachments = await orm.mail.getAttachments('user-id', 'message-id');
await orm.mail.addAttachment('user-id', 'message-id', attachmentData);
```

### 日历操作

```typescript
// 创建事件
await orm.calendar.createEvent('user-id', {
  subject: '团队会议',
  start: {
    dateTime: '2024-01-15T10:00:00',
    timeZone: 'China Standard Time'
  },
  end: {
    dateTime: '2024-01-15T11:00:00',
    timeZone: 'China Standard Time'
  },
  attendees: [
    {
      type: 'required',
      emailAddress: { address: 'colleague@example.com' }
    }
  ]
});

// 日历视图
const events = await orm.calendar.getCalendarView(
  'user-id',
  '2024-01-01T00:00:00Z',
  '2024-01-31T23:59:59Z'
);

// 事件响应
await orm.calendar.acceptEvent('user-id', 'event-id', '我会参加');
await orm.calendar.declineEvent('user-id', 'event-id', '时间冲突');

// 会议室
const rooms = await orm.calendar.findRooms('user-id');
const schedule = await orm.calendar.getSchedule(
  ['user1@contoso.com', 'room@contoso.com'],
  startTime,
  endTime
);
```

## ⚡ 高级功能

### 增量查询 (Delta Query)

```typescript
// 首次查询
const delta = await orm.deltaUsers();
console.log('用户数:', delta.data.length);

// 保存 deltaLink
const deltaLink = delta.meta.deltaLink;

// 后续查询（只返回变化）
const changes = await orm.deltaUsers(deltaLink);
console.log('变化的用户:', changes.data);

// 其他资源
const groupDelta = await orm.deltaGroups();
const messageDelta = await orm.deltaMessages('user-id');
const eventDelta = await orm.deltaEvents('user-id');
const fileDelta = await orm.deltaDriveItems('user-id');
```

### 批处理请求

```typescript
const response = await orm.batch([
  { id: '1', method: 'GET', url: '/users/user1@contoso.com' },
  { id: '2', method: 'GET', url: '/users/user2@contoso.com' },
  {
    id: '3',
    method: 'PATCH',
    url: '/users/user3@contoso.com',
    body: { jobTitle: '工程师' }
  }
]);

response.responses.forEach(res => {
  console.log(`请求 ${res.id}: 状态 ${res.status}`);
});
```

### 搜索 API

```typescript
// 搜索用户
const userResults = await orm.searchUsers('John Doe', 10);

// 搜索邮件
const mailResults = await orm.searchMessages('project report', 20);

// 搜索文件
const fileResults = await orm.searchFiles('budget.xlsx', 10);

// 自定义搜索
const customSearch = await orm.search({
  entityTypes: ['driveItem', 'message'],
  query: { queryString: '财务报告' },
  from: 0,
  size: 25
});
```

## 🌏 国家云支持

```typescript
import { GraphClient, defaultEndpoints, createGraphORM } from '@qingu-x/msgraph-orm-js';

// 中国版（21Vianet）
const graphClientCN = new GraphClient(
  {
    tenantId: 'your-tenant-id',
    clientId: 'your-client-id',
    clientSecret: 'your-client-secret'
  },
  defaultEndpoints.cn  // 中国版端点
);

// 美国政府版（GCC）
const graphClientGCC = new GraphClient(
  {
    tenantId: 'your-tenant-id',
    clientId: 'your-client-id',
    clientSecret: 'your-client-secret'
  },
  defaultEndpoints.usGCC  // 美国政府版 GCC 端点
);

// 美国政府版（GCC High）
const graphClientGCCHigh = new GraphClient(
  {
    tenantId: 'your-tenant-id',
    clientId: 'your-client-id',
    clientSecret: 'your-client-secret'
  },
  defaultEndpoints.usGCCHigh  // 美国政府版 GCC High 端点
);

// 美国政府版（DoD）
const graphClientDoD = new GraphClient(
  {
    tenantId: 'your-tenant-id',
    clientId: 'your-client-id',
    clientSecret: 'your-client-secret'
  },
  defaultEndpoints.usDOD  // 美国政府版 DoD 端点
);

// 创建 ORM
const orm = createGraphORM(graphClientCN.getGraphClient());
```

## 📖 文档

- [完整 API 文档](./docs/DOCS_INDEX.md)
- [使用示例](./docs/EXAMPLES.md)
- [Graph API 参考](./docs/GRAPH_API_REFERENCE.md)
- [ORM 指南](./docs/GRAPH_ORM_GUIDE.md)
- [贡献指南](./docs/CONTRIBUTING.md)

## 🤝 贡献

欢迎贡献代码！请查看 [贡献指南](./docs/CONTRIBUTING.md)。

## 📄 许可证

MIT © [qingu-x](https://github.com/qingu-x)

## ⚠️ 免责声明

本项目不是微软官方项目，仅为 Microsoft Graph API 的封装。使用前请确保：

1. 具有适当的 API 权限
2. 遵守微软 Graph API 的使用条款
3. 了解国家云的功能限制

## 📮 反馈

遇到问题或有建议？请：

- 提交 [Issue](https://github.com/qingu-x/msgraph-orm-js/issues)
- 查看 [示例代码](./docs/EXAMPLES.md)
- 查看 [完整文档](./docs/DOCS_INDEX.md)

