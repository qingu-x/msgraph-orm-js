# Microsoft Graph ORM 使用指南

## 概述

本 SDK 提供了一个完整的 ORM（对象关系映射）层来简化 Microsoft Graph API 的使用，支持全球版和国家云部署（中国、美国政府版等）。

## 国家云部署支持

### 支持的云环境

- ✅ **全球版（Global）**：所有功能完全支持
- ⚠️ **中国版（21Vianet）**：部分功能不可用或受限
- ⚠️ **美国政府版（GCC, GCC High, DoD）**：大部分功能支持

**注意**：德国版（Microsoft Cloud Deutschland）已于 2021 年弃用，现已迁移至欧洲的全球云区域。

### 功能支持矩阵

| 功能 | 全球版 | 中国版 | 美国政府版 | 说明 |
|------|--------|--------|-----------|------|
| 用户管理 | ✅ | ✅ | ✅ | 完全支持 |
| 组管理 | ✅ | ✅ | ✅ | 完全支持 |
| 设备管理 | ✅ | ✅ | ✅ | 完全支持 |
| 日历事件 | ✅ | ✅ | ✅ | 完全支持 |
| 邮件管理 | ✅ | ✅ | ✅ | 完全支持 |
| 联系人管理 | ✅ | ✅ | ✅ | 完全支持 |
| OneDrive | ✅ | ✅ | ✅ | 完全支持 |
| SharePoint | ✅ | ✅ | ⚠️ | 部分支持 |
| Places API（会议室） | ✅ | ⚠️ | ⚠️ | 中国版功能受限 |
| findRooms | ✅ | ❌ | ⚠️ | 中国版不可用 |
| 批处理请求 | ✅ | ✅ | ✅ | 完全支持 |

## 快速开始

### 1. 初始化客户端

```typescript
import { ApplicationClient, defaultEndpoints, createGraphORM } from 'msgraph-orm-js';

// 全球版
const globalClient = new ApplicationClient({
  tenantId: 'your-tenant-id',
  clientId: 'your-client-id',
  clientSecret: 'your-client-secret'
}, defaultEndpoints.global);

// 中国版（21Vianet）
const cnClient = new ApplicationClient({
  tenantId: 'your-tenant-id',
  clientId: 'your-client-id',
  clientSecret: 'your-client-secret'
}, defaultEndpoints.cn);

const orm = createGraphORM(globalClient.getGraphClient());
```

## 核心功能

### 查询构建器（Query Builder）

#### 基本查询

```typescript
// 获取所有用户
const users = await orm.users.findMany();

// 条件查询
const result = await orm.users.query()
  .where('displayName', 'contains', '张')
  .top(10)
  .execute();

// 复杂查询
const result = await orm.users.query()
  .where('jobTitle', 'eq', '经理')
  .and('department', 'eq', '销售部')
  .orderBy('displayName', 'asc')
  .select(['id', 'displayName', 'mail'])
  .top(20)
  .execute();
```

#### 新增功能

##### 1. $expand - 展开导航属性

```typescript
// 获取用户及其管理者信息
const result = await orm.users.query()
  .select(['id', 'displayName', 'mail'])
  .expand('manager')
  .top(10)
  .execute();

// 展开多个属性
const result = await orm.users.query()
  .expand(['manager', 'memberOf'])
  .top(5)
  .execute();

// 嵌套展开并选择字段
const result = await orm.users.query()
  .expand('manager($select=displayName,mail)')
  .execute();
```

##### 2. $skiptoken - 分页令牌

```typescript
// 获取第一页
const firstPage = await orm.users.query()
  .top(10)
  .execute();

// 提取 skiptoken
if (firstPage.meta.nextLink) {
  const url = new URL(firstPage.meta.nextLink);
  const skiptoken = url.searchParams.get('$skiptoken');
  
  // 获取下一页
  const nextPage = await orm.users.query()
    .top(10)
    .skipToken(skiptoken!)
    .execute();
}
```

##### 3. executeWithPagination - 自动分页

```typescript
// 自动处理分页，迭代所有数据
for await (const user of orm.users.query().top(50).executeWithPagination()) {
  console.log(user.displayName);
  // 自动获取所有页，无需手动处理 nextLink
}

// 带条件的自动分页
for await (const user of orm.users.query()
  .where('department', 'eq', '销售部')
  .select(['displayName', 'mail'])
  .executeWithPagination()
) {
  console.log(user.displayName, user.mail);
}
```

##### 4. $format - 响应格式

```typescript
const result = await orm.users.query()
  .format('json')
  .top(10)
  .execute();
```

### 资源类型支持

#### 用户和身份

```typescript
// 用户管理
const users = await orm.users.findMany();
const user = await orm.users.findById('user@contoso.com');

// 组管理
const groups = await orm.groups.findMany();
const groupMembers = await orm.groupMembers('group-id').findMany();
const groupOwners = await orm.groupOwners('group-id').findMany();

// 设备管理
const devices = await orm.devices.findMany();

// 应用程序管理
const apps = await orm.applications.findMany();

// 服务主体管理
const servicePrincipals = await orm.servicePrincipals.findMany();
```

#### 日历和事件

```typescript
// 用户事件
const events = await orm.events('user@contoso.com').findMany();

// 日历事件
const calendarEvents = await orm.calendarEvents('user@contoso.com')
  .query()
  .where('start/dateTime', 'ge', new Date().toISOString())
  .orderBy('start/dateTime', 'asc')
  .execute();

// 指定日历的事件
const events = await orm.calendarEventsById('user@contoso.com', 'calendar-id')
  .findMany();

// 组事件
const groupEvents = await orm.groupEvents('group-id').findMany();
```

#### 会议室管理

```typescript
// Places API（全球版推荐，中国版可能不可用）
try {
  const rooms = await orm.rooms.findMany();
  const roomLists = await orm.roomLists.findMany();
} catch (error) {
  console.log('Places API 不可用，尝试使用 findRooms');
}

// findRooms 函数（权限要求较低，但中国版不可用）
const rooms = await orm.findRooms('user@contoso.com');
const roomLists = await orm.findRoomLists('user@contoso.com');
```

#### 邮件和联系人

```typescript
// 用户邮件
const messages = await orm.messages('user@contoso.com')
  .query()
  .where('isRead', false)
  .orderBy('receivedDateTime', 'desc')
  .top(20)
  .execute();

// 收件箱
const inbox = await orm.inbox('user@contoso.com').findMany();

// 已发送邮件
const sentItems = await orm.sentItems('user@contoso.com').findMany();

// 草稿箱
const drafts = await orm.drafts('user@contoso.com').findMany();

// 联系人
const contacts = await orm.contacts('user@contoso.com').findMany();

// 指定文件夹的联系人
const folderContacts = await orm.contactsInFolder('user@contoso.com', 'folder-id')
  .findMany();
```

#### OneDrive 和 SharePoint

```typescript
// 用户驱动器
const driveItems = await orm.driveItems('user@contoso.com').findMany();

// 指定路径的项
const documents = await orm.driveItemsByPath('user@contoso.com', '/Documents')
  .findMany();

// 组驱动器
const groupDrive = await orm.groupDriveItems('group-id').findMany();

// SharePoint 站点
const sites = await orm.sites.findMany();

// 根站点
const rootSite = await orm.rootSite;

// 站点驱动器项
const siteDocs = await orm.siteDriveItems('site-id').findMany();
```

#### 批处理请求

```typescript
// 在一个请求中获取多个用户
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

## 查询操作符

### 支持的操作符

```typescript
// 比较操作符
.where('field', 'eq', 'value')    // 等于
.where('field', 'ne', 'value')    // 不等于（User 和 Device 资源不支持）
.where('field', 'gt', 100)        // 大于
.where('field', 'ge', 100)        // 大于等于
.where('field', 'lt', 100)        // 小于
.where('field', 'le', 100)        // 小于等于

// 字符串操作符（需要 ConsistencyLevel: eventual，不支持排序）
.where('field', 'startswith', 'value')  // 开头匹配
.where('field', 'endswith', 'value')    // 结尾匹配
.where('field', 'contains', 'value')    // 包含

// 逻辑操作符
.where('field1', 'eq', 'value1')
.and('field2', 'eq', 'value2')
.or('field3', 'eq', 'value3')
```

### 操作符限制

#### User 和 Device 资源限制

```typescript
// ❌ 错误：User 资源不支持 ne 操作符
try {
  await orm.users.query()
    .where('mail', 'ne', null)
    .execute();
} catch (error) {
  // GraphOrmError: User 资源不支持 'ne' (不等于) 操作符
}

// ✅ 正确：使用客户端过滤
const users = await orm.users.query().execute();
const filteredUsers = users.data.filter(u => u.mail !== null);
```

#### 高级操作符与排序

```typescript
// ❌ 错误：使用 startswith 时不能排序
try {
  await orm.users.query()
    .where('displayName', 'startswith', '张')
    .orderBy('displayName', 'asc')  // 会抛出错误
    .execute();
} catch (error) {
  // GraphOrmError: 使用高级查询操作符时不支持排序
}

// ✅ 正确：移除排序或使用客户端排序
const result = await orm.users.query()
  .where('displayName', 'startswith', '张')
  .execute();
const sorted = result.data.sort((a, b) => 
  a.displayName.localeCompare(b.displayName)
);
```

## 错误处理

```typescript
import { GraphOrmError, GraphErrorCode } from 'msgraph-orm-js';

try {
  const user = await orm.users.findById('invalid-id');
} catch (error) {
  if (error instanceof GraphOrmError) {
    console.log('错误代码:', error.code);
    console.log('错误消息:', error.message);
    console.log('HTTP 状态:', error.statusCode);
    console.log('请求 ID:', error.requestId);
    
    // 根据错误代码处理
    switch (error.code) {
      case GraphErrorCode.RESOURCE_NOT_FOUND:
        console.log('资源未找到');
        break;
      case GraphErrorCode.AUTHENTICATION_FAILED:
        console.log('认证失败');
        break;
      case GraphErrorCode.INSUFFICIENT_PERMISSIONS:
        console.log('权限不足');
        break;
      default:
        console.log('其他错误');
    }
  }
}
```

## 最佳实践

### 1. 使用 $select 减少数据传输

```typescript
// ❌ 不推荐：获取所有字段
const users = await orm.users.findMany();

// ✅ 推荐：只获取需要的字段
const users = await orm.users.query()
  .select(['id', 'displayName', 'mail'])
  .execute();
```

### 2. 使用 $expand 减少请求次数

```typescript
// ❌ 不推荐：多次请求
const users = await orm.users.findMany();
for (const user of users.data) {
  const manager = await orm.users.findById(user.manager);
}

// ✅ 推荐：一次请求获取关联数据
const users = await orm.users.query()
  .expand('manager($select=displayName,mail)')
  .execute();
```

### 3. 使用 executeWithPagination 处理大量数据

```typescript
// ❌ 不推荐：可能超时或内存溢出
const allUsers = await orm.users.findMany();

// ✅ 推荐：流式处理
for await (const user of orm.users.query().executeWithPagination()) {
  // 逐个处理用户
  await processUser(user);
}
```

### 4. 使用批处理减少网络往返

```typescript
// ❌ 不推荐：多个单独请求
const user1 = await orm.users.findById('user1@contoso.com');
const user2 = await orm.users.findById('user2@contoso.com');
const user3 = await orm.users.findById('user3@contoso.com');

// ✅ 推荐：批处理请求
const results = await orm.batch([
  { id: '1', method: 'GET', url: '/users/user1@contoso.com' },
  { id: '2', method: 'GET', url: '/users/user2@contoso.com' },
  { id: '3', method: 'GET', url: '/users/user3@contoso.com' }
]);
```

### 5. 国家云兼容性处理

```typescript
// 为不同云环境提供降级方案
async function getRooms(userId: string) {
  try {
    // 尝试使用 Places API
    return await orm.rooms.findMany();
  } catch (error) {
    console.log('Places API 不可用，尝试 findRooms');
    try {
      // 降级到 findRooms
      return await orm.findRooms(userId);
    } catch (error2) {
      console.log('findRooms 也不可用（可能是中国版）');
      return [];
    }
  }
}
```

## 权限参考

| 资源类型 | 读取权限 | 写入权限 |
|---------|---------|---------|
| 用户 | User.Read.All | User.ReadWrite.All |
| 组 | Group.Read.All | Group.ReadWrite.All |
| 设备 | Device.Read.All | Device.ReadWrite.All |
| 日历 | Calendars.Read | Calendars.ReadWrite |
| 邮件 | Mail.Read | Mail.ReadWrite |
| 联系人 | Contacts.Read | Contacts.ReadWrite |
| 文件 | Files.Read.All | Files.ReadWrite.All |
| 站点 | Sites.Read.All | Sites.ReadWrite.All |
| 会议室 | Place.Read.All | - |
| 应用 | Application.Read.All | Application.ReadWrite.All |

## 参考链接

- [Microsoft Graph 文档](https://learn.microsoft.com/graph/)
- [OData 查询参数](https://learn.microsoft.com/graph/query-parameters)
- [国家云部署](https://learn.microsoft.com/graph/deployments)
- [权限参考](https://learn.microsoft.com/graph/permissions-reference)
- [批处理请求](https://learn.microsoft.com/graph/json-batching)
- [高级查询](https://learn.microsoft.com/graph/aad-advanced-queries)

