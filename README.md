# @qingu-x/msgraph-orm-js

[![npm version](https://img.shields.io/npm/v/@qingu-x/msgraph-orm-js.svg)](https://www.npmjs.com/package/@qingu-x/msgraph-orm-js)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

Microsoft Graph 的类 ORM 封装，提供简洁、类型安全的 API 调用方式。

> ⚠️ 本项目包括文档大部分由 AI 完成，如遇到问题请提 [Issue](https://github.com/qingu-x/msgraph-orm-js/issues)

## ✨ 特性

- 🎯 **类 ORM 接口**: 直观的链式调用
- 📘 **完整类型支持**: 100% TypeScript 编写
- 🌏 **国家云支持**: 全球版/中国版/美国政府版
- 🚀 **95%+ API 覆盖**: 覆盖 Graph v1.0 绝大多数常用接口
- 🔍 **强大查询构建器**: 支持复杂查询、分页、排序
- ⚡ **高级功能**: Delta Query、批处理、Webhooks

## 📦 安装

```bash
npm install @qingu-x/msgraph-orm-js @microsoft/microsoft-graph-client
```

## 🚀 快速开始

```typescript
import { Client } from '@microsoft/microsoft-graph-client';
import { ClientCredentialsAuthProvider, createGraphORM } from '@qingu-x/msgraph-orm-js';

// 1. 创建认证提供者
const authProvider = new ClientCredentialsAuthProvider({
  tenantId: 'your-tenant-id',
  clientId: 'your-client-id',
  clientSecret: 'your-client-secret'
});

// 2. 创建 Graph 客户端
const client = Client.initWithMiddleware({ authProvider });

// 3. 创建 ORM 实例
const orm = createGraphORM(client);

// 4. 开始使用
const users = await orm.users
  .query()
  .select(['displayName', 'mail'])
  .top(10)
  .get();
```

### 国家云支持

```typescript
// 中国版
const authProvider = new ClientCredentialsAuthProvider({
  // ...其他配置
  cloudEndpoint: 'https://microsoftgraph.chinacloudapi.cn'
});

// 美国政府版
const authProvider = new ClientCredentialsAuthProvider({
  // ...其他配置
  cloudEndpoint: 'https://graph.microsoft.us'
});
```

## 📚 核心功能

### 用户管理

```typescript
// 查询用户
const users = await orm.users
  .query()
  .where('jobTitle', 'contains', '经理')
  .orderBy('displayName', 'asc')
  .get();

// CRUD 操作
const user = await orm.users.findById('user-id');
const newUser = await orm.users.create({ /* ... */ });
await orm.users.update('user-id', { jobTitle: '高级工程师' });
await orm.users.delete('user-id');
```

### 日历管理

```typescript
// 获取日历事件
const events = await orm.events('user@contoso.com')
  .query()
  .orderBy('start/dateTime', 'asc')
  .top(20)
  .get();

// 日历视图（时间范围）
const upcomingEvents = await orm.getCalendarView(
  'user@contoso.com',
  startDate,
  endDate
);

// 查找会议室
const rooms = await orm.findRooms('user@contoso.com');
```

### 邮件管理

```typescript
// 读取收件箱
const messages = await orm.inbox('user@contoso.com')
  .query()
  .top(10)
  .get();

// 发送邮件
await orm.sendMail('user@contoso.com', {
  subject: '测试邮件',
  body: { contentType: 'HTML', content: '<p>内容</p>' },
  toRecipients: [{ emailAddress: { address: 'recipient@contoso.com' } }]
});
```

### 文件管理

```typescript
// 列出文件
const files = await orm.driveItems('user@contoso.com').query().get();

// 搜索文件
const results = await orm.searchDriveItems('user@contoso.com', '报告');

// 上传文件
await orm.uploadSmallFile('user@contoso.com', '/path/file.pdf', buffer);

// 创建共享链接
const link = await orm.createDriveItemLink(
  'user@contoso.com',
  'file-id',
  'view',
  'organization'
);
```

### Teams 协作

```typescript
// 获取团队
const teams = await orm.teams.query().get();

// 发送频道消息
await orm.sendChannelMessage('team-id', 'channel-id', {
  body: { contentType: 'html', content: '<h1>通知</h1>' }
});

// 团队成员
const members = await orm.teamMembers('team-id').query().get();
```

### SharePoint

```typescript
// 获取站点
const site = await orm.getSiteByPath('contoso.sharepoint.com', '/sites/team');

// 列表操作
const items = await orm.listItems('site-id', 'list-id').query().get();
await orm.listItems('site-id', 'list-id').create({
  fields: { Title: '新项目', Status: '进行中' }
});
```

## 🔍 高级功能

### Delta Query（变化跟踪）

```typescript
// 首次查询
const result = await orm.deltaUsers();
const deltaLink = result.meta.deltaLink;

// 后续只获取变化
const changes = await orm.deltaUsers(deltaLink);
```

### 批处理请求

```typescript
const response = await orm.batch([
  { id: '1', method: 'GET', url: '/users/user1@contoso.com' },
  { id: '2', method: 'GET', url: '/groups' }
]);
```

### Webhooks 订阅

```typescript
const subscription = await orm.createSubscription({
  changeType: 'created,updated',
  notificationUrl: 'https://your-app.com/webhooks',
  resource: 'users',
  expirationDateTime: new Date(Date.now() + 3600000).toISOString()
});
```

## 📖 API 覆盖范围

| 功能模块 | 覆盖率 | 国家云支持 |
|---------|--------|-----------|
| 用户和身份 | ✅ 100% | 全球/中国/美国政府 |
| 日历 | ✅ 100% | 全球/中国/美国政府 |
| 邮件 | ✅ 100% | 全球/中国/美国政府 |
| 联系人 | ✅ 100% | 全球/中国/美国政府 |
| 文件存储 | ✅ 100% | 全球/中国/美国政府 |
| OneNote | ✅ 100% | 全球/美国政府 |
| 任务计划 | ✅ 100% | 全球/美国政府 |
| SharePoint | ✅ 100% | 全球/中国/美国政府 |
| Teams | ✅ 100% | 全球/美国政府 |
| 报告分析 | ✅ 95% | 全球/美国政府 |
| 安全合规 | ✅ 90% | 全球/美国政府 |

详细覆盖情况请查看 [API_COVERAGE.md](./docs/API_COVERAGE.md)

## 📝 文档

| 文档 | 描述 |
|------|------|
| [README.md](./README.md) | 项目主文档（当前） |
| [GRAPH_ORM_GUIDE.md](./docs/GRAPH_ORM_GUIDE.md) | ORM 使用教程 |
| [EXAMPLES.md](./docs/EXAMPLES.md) | 完整代码示例 |
| [API_COVERAGE.md](./docs/API_COVERAGE.md) | API 覆盖率 |
| [CONTRIBUTING.md](./docs/CONTRIBUTING.md) | 开发贡献指南 |
| [PUBLISHING.md](./docs/PUBLISHING.md) | NPM 发布指南 |
| [DOCS_INDEX.md](./docs/DOCS_INDEX.md) | 完整文档索引 |

## 🔐 权限说明

常用权限：

| 权限 | 说明 | 类型 |
|------|------|------|
| User.Read.All | 读取所有用户 | 应用 |
| Mail.Send | 发送邮件 | 委托 |
| Calendars.ReadWrite | 读写日历 | 委托/应用 |
| Files.ReadWrite.All | 读写文件 | 应用 |
| Group.ReadWrite.All | 读写组 | 应用 |

完整权限列表：[Microsoft Graph 权限参考](https://learn.microsoft.com/graph/permissions-reference)

## 🌟 最佳实践

```typescript
// ✅ 使用 Select 减少数据传输
const users = await orm.users
  .query()
  .select(['id', 'displayName', 'mail'])
  .get();

// ✅ 使用分页控制结果数量
const users = await orm.users.query().top(100).get();

// ✅ 使用 Delta Query 跟踪变化
const changes = await orm.deltaUsers(savedDeltaLink);

// ✅ 批处理减少请求次数
const batch = await orm.batch([/* 多个请求 */]);
```

## 🐛 错误处理

```typescript
import { GraphOrmError, GraphErrorCode } from '@qingu-x/msgraph-orm-js';

try {
  const user = await orm.users.findById('invalid-id');
} catch (error) {
  if (error instanceof GraphOrmError) {
    switch (error.code) {
      case GraphErrorCode.RESOURCE_NOT_FOUND:
        console.log('资源未找到');
        break;
      case GraphErrorCode.UNAUTHORIZED:
        console.log('权限不足');
        break;
      case GraphErrorCode.THROTTLED:
        console.log('请求被限流');
        break;
    }
  }
}
```

## 🤝 贡献

欢迎提交 Issue 和 Pull Request！

详细指南：

- [CONTRIBUTING.md](./docs/CONTRIBUTING.md) - 开发环境、构建、测试、代码规范
- [PUBLISHING.md](./docs/PUBLISHING.md) - 发布流程

## 📄 许可证

MIT License

## 🔗 相关链接

- [Microsoft Graph 文档](https://learn.microsoft.com/graph/)
- [Graph Explorer](https://developer.microsoft.com/graph/graph-explorer)
- [Graph SDK](https://github.com/microsoftgraph/msgraph-sdk-javascript)
- [国家云部署](https://learn.microsoft.com/graph/deployments)

## 💡 获取帮助

1. 查看 [文档](./docs/DOCS_INDEX.md)
2. 查看 [示例代码](./docs/EXAMPLES.md)
3. 提交 [Issue](https://github.com/qingu-x/msgraph-orm-js/issues)

---

Made with ❤️ by Qingu-X
