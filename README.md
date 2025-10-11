# @qingu-x/msgraph-orm-js

[![npm version](https://img.shields.io/npm/v/@qingu-x/msgraph-orm-js.svg)](https://www.npmjs.com/package/@qingu-x/msgraph-orm-js)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

Microsoft Graph JavaScript/TypeScript SDK 的类 ORM 封装，提供简洁、类型安全的 API 调用方式。

## ✨ 特性

- 🎯 **类 ORM 接口**: 直观的链式调用，类似 TypeORM/Sequelize 的开发体验
- 📘 **完整类型支持**: 100% TypeScript 编写，完整的类型定义和 IntelliSense
- 🌏 **国家云支持**: 完整支持全球版、中国版、美国政府版云环境
- 🔐 **多种认证方式**: 支持客户端凭据、委托权限等多种认证
- 🚀 **95%+ API 覆盖**: 覆盖 Microsoft Graph v1.0 的绝大多数常用接口
- 🔍 **强大查询构建器**: 支持复杂查询、分页、排序、展开等
- ⚡ **高级功能**: Delta Query、批处理、Webhooks 等
- 📝 **详细文档**: 完整的 API 文档和使用示例
- 🛡️ **权限提示**: 每个 API 都标注了所需权限

## 📦 安装

```bash
npm install @qingu-x/msgraph-orm-js @microsoft/microsoft-graph-client
```

## 🚀 快速开始

### 基础用法

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

// 4. 开始使用！
const users = await orm.users
  .query()
  .select(['displayName', 'mail'])
  .top(10)
  .execute();

console.log('用户列表:', users.data);
```

### 国家云支持

```typescript
// 中国版（由世纪互联运营）
const authProvider = new ClientCredentialsAuthProvider({
  tenantId: 'your-tenant-id',
  clientId: 'your-client-id',
  clientSecret: 'your-client-secret',
  cloudEndpoint: 'https://microsoftgraph.chinacloudapi.cn'
});

// 美国政府版
const authProvider = new ClientCredentialsAuthProvider({
  tenantId: 'your-tenant-id',
  clientId: 'your-client-id',
  clientSecret: 'your-client-secret',
  cloudEndpoint: 'https://graph.microsoft.us'
});
```

## 📚 核心功能

### 用户管理

```typescript
// 查询用户
const managers = await orm.users
  .query()
  .where('jobTitle', 'contains', '经理')
  .and('accountEnabled', 'eq', true)
  .orderBy('displayName', 'asc')
  .execute();

// 获取单个用户
const user = await orm.users.findById('user-id');

// 创建用户
const newUser = await orm.users.create({
  displayName: '张三',
  userPrincipalName: 'zhangsan@contoso.com',
  // ...
});

// 更新用户
await orm.users.update('user-id', {
  jobTitle: '高级工程师'
});
```

### 日历管理

```typescript
// 获取日历事件
const events = await orm.events('user@contoso.com')
  .query()
  .orderBy('start/dateTime', 'asc')
  .top(20)
  .execute();

// 获取日历视图（指定时间范围）
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
  .execute();

// 发送邮件
await orm.sendMail('user@contoso.com', {
  subject: '测试邮件',
  body: {
    contentType: 'HTML',
    content: '<p>邮件内容</p>'
  },
  toRecipients: [
    { emailAddress: { address: 'recipient@contoso.com' } }
  ]
});
```

### 文件管理

```typescript
// 列出文件
const files = await orm.driveItems('user@contoso.com').query().execute();

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
const teams = await orm.teams.query().execute();

// 发送频道消息
await orm.sendChannelMessage('team-id', 'channel-id', {
  body: {
    contentType: 'html',
    content: '<h1>通知</h1>'
  }
});

// 获取团队成员
const members = await orm.teamMembers('team-id').query().execute();
```

### SharePoint

```typescript
// 获取站点
const site = await orm.getSiteByPath('contoso.sharepoint.com', '/sites/team');

// 操作列表
const items = await orm.listItems('site-id', 'list-id').query().execute();

// 创建列表项
await orm.listItems('site-id', 'list-id').create({
  fields: {
    Title: '新项目',
    Status: '进行中'
  }
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
console.log('变化的用户:', changes.data);
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

### Microsoft Search

```typescript
// 搜索邮件
const results = await orm.searchMessages('项目会议', 25);

// 搜索文件
const files = await orm.searchFiles('财务报告', 10);
```

## 📖 API 覆盖范围

本 SDK 覆盖了 Microsoft Graph v1.0 的 95%+ 常用接口：

| 功能模块 | 覆盖率 | 国家云支持 |
|---------|--------|-----------|
| 用户和身份管理 | ✅ 100% | 全球版/中国版/美国政府版 |
| 日历 | ✅ 100% | 全球版/中国版/美国政府版 |
| 邮件 | ✅ 100% | 全球版/中国版/美国政府版 |
| 联系人 | ✅ 100% | 全球版/中国版/美国政府版 |
| 文件存储 (OneDrive) | ✅ 100% | 全球版/中国版/美国政府版 |
| OneNote | ✅ 100% | 全球版/美国政府版（中国版不可用） |
| 任务和计划 (Planner/To Do) | ✅ 100% | 全球版/美国政府版（中国版不可用） |
| SharePoint | ✅ 100% | 全球版/中国版/美国政府版 |
| Teams | ✅ 100% | 全球版/美国政府版（中国版功能受限） |
| 报告和分析 | ✅ 95% | 全球版/美国政府版（中国版部分功能受限） |
| 安全与合规 | ✅ 90% | 全球版/美国政府版（中国版不可用） |
| 扩展 | ✅ 100% | 全球版/中国版/美国政府版 |

详细覆盖情况请查看 [API_COVERAGE.md](./API_COVERAGE.md)

## 📝 文档

> 💡 **提示**: 查看 [文档索引](./DOCS_INDEX.md) 快速找到你需要的文档！

- [📑 文档索引](./DOCS_INDEX.md) - 所有文档的导航和说明
- [🎓 快速开始指南](./GRAPH_ORM_GUIDE.md) - ORM 使用教程
- [💡 完整示例](./EXAMPLES.md) - 各种场景的代码示例
- [📘 API 参考](./GRAPH_API_REFERENCE.md) - 所有 API 端点说明
- [✅ API 覆盖率](./API_COVERAGE.md) - 接口实现情况
- [📝 更新日志](./CHANGELOG.md) - 版本更新记录
- [🎯 实施总结](./IMPLEMENTATION_SUMMARY.md) - 完整实施报告

## 🔐 权限说明

每个 API 都需要相应的权限。常用权限包括：

| 权限 | 说明 | 类型 |
|------|------|------|
| User.Read.All | 读取所有用户完整配置文件 | 应用 |
| Mail.Send | 以用户身份发送邮件 | 委托 |
| Calendars.ReadWrite | 读写用户日历 | 委托/应用 |
| Files.ReadWrite.All | 读写所有文件 | 应用 |
| Group.ReadWrite.All | 读写所有组 | 应用 |

完整权限列表请参考 [Microsoft Graph 权限参考](https://learn.microsoft.com/graph/permissions-reference)

## 🌟 最佳实践

### 1. 使用 Select 减少数据传输

```typescript
// ✅ 好：只获取需要的字段
const users = await orm.users
  .query()
  .select(['id', 'displayName', 'mail'])
  .execute();

// ❌ 不好：获取所有字段
const users = await orm.users.query().execute();
```

### 2. 使用分页控制结果数量

```typescript
const users = await orm.users.query().top(100).execute();
```

### 3. 使用 Delta Query 跟踪变化

```typescript
// 定期检查变化而不是每次全量查询
const changes = await orm.deltaUsers(savedDeltaLink);
```

### 4. 批处理减少请求次数

```typescript
const batch = await orm.batch([/* 多个请求 */]);
```

### 5. 合理处理国家云差异

```typescript
try {
  const notebooks = await orm.notebooks('user@contoso.com').query().execute();
} catch (error) {
  // OneNote 在中国版不可用，使用替代方案
}
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

## 📄 许可证

MIT License

## 🔗 相关链接

- [Microsoft Graph 文档](https://learn.microsoft.com/graph/)
- [Microsoft Graph Explorer](https://developer.microsoft.com/graph/graph-explorer)
- [Microsoft Graph JavaScript SDK](https://github.com/microsoftgraph/msgraph-sdk-javascript)
- [国家云部署](https://learn.microsoft.com/graph/deployments)

## 💡 支持

如有问题，请：
1. 查看 [文档](./GRAPH_ORM_GUIDE.md)
2. 查看 [示例代码](./EXAMPLES.md)
3. 提交 [Issue](https://github.com/qingu-x/msgraph-orm-js/issues)

---

Made with ❤️ by Qingu-X

