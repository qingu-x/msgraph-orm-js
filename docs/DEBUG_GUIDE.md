# ORM 调试指南

本指南展示如何在使用 Graph ORM 时获取调试信息。

## 基本用法

### 1. Repository 调试

#### 查询操作调试

```typescript
import { createGraphORM } from '@qingu-x/msgraph-orm-js';

const orm = createGraphORM(client);

// 构建查询
const query = orm.users.query()
  .where('department', 'eq', 'IT')
  .select(['id', 'displayName', 'mail'])
  .top(10);

// 获取调试信息（执行前）
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

#### CRUD 操作调试

```typescript
// 对于 orm.users.events($userid).create({}) 这种操作
const userId = 'user-id';
const eventData = {
  subject: '测试会议',
  start: { dateTime: '2024-01-15T10:00:00', timeZone: 'Asia/Shanghai' },
  end: { dateTime: '2024-01-15T11:00:00', timeZone: 'Asia/Shanghai' }
};

const eventsRepo = orm.users.events(userId);

// 方法一：通过 query() 获取基础调试信息
const query = eventsRepo.query();
const debugInfo = query.getDebug();
console.log('=== 事件创建调试信息 ===');
console.log('请求方法:', debugInfo.method);        // "GET" (query 的默认方法)
console.log('请求 URL:', debugInfo.url);           // "/users/user-id/events"
console.log('请求头:', debugInfo.headers);

// 执行创建操作
const event = await eventsRepo.create(eventData);
console.log('事件创建成功:', event.id);

// 方法二：调试包装器
async function debugRepositoryOperation<T>(
  operation: string,
  execute: () => Promise<T>
): Promise<T> {
  const query = eventsRepo.query();
  const debugInfo = query.getDebug();
  
  console.log(`=== ${operation} 调试信息 ===`);
  console.log('基础 URL:', debugInfo.url);
  console.log('请求方法:', debugInfo.method);
  console.log('请求头:', debugInfo.headers);
  console.log('时间戳:', new Date().toISOString());
  
  try {
    const result = await execute();
    console.log(`${operation} 成功`);
    return result;
  } catch (error) {
    console.error(`${operation} 失败:`, error);
    throw error;
  }
}

// 使用调试包装器
const event = await debugRepositoryOperation(
  '创建事件',
  () => eventsRepo.create(eventData)
);
```

### 2. Service 调试

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

### 3. GraphORM 调试

```typescript
// GraphORM 方法调试
await orm.me();
const debugInfo = orm.getDebug();
console.log('GraphORM 调试信息:', debugInfo);

// 增量查询调试
await orm.deltaUsers();
const deltaDebugInfo = orm.getDebug();
console.log('增量查询调试信息:', deltaDebugInfo);

// 批处理调试
await orm.batch(requests);
const batchDebugInfo = orm.getDebug();
console.log('批处理调试信息:', batchDebugInfo);
```

## 调试信息结构

```typescript
interface RequestDebugInfo {
  method: string;           // HTTP 方法 (GET, POST, PUT, PATCH, DELETE)
  url: string;             // 完整请求 URL
  queryParams?: Record<string, string | number>; // 查询参数
  headers: Record<string, string>; // 请求头
  body?: unknown;          // 请求体（POST/PUT/PATCH 请求）
  timestamp?: string;      // 请求时间戳
}
```

## 实际使用示例

### 用户管理调试

```typescript
// 查询用户
const usersQuery = orm.users.query()
  .where('department', 'eq', 'IT')
  .orderBy('displayName', 'asc')
  .top(20);

// 调试信息
const debugInfo = usersQuery.getDebug();
console.log('用户查询调试:', {
  method: debugInfo.method,        // "GET"
  url: debugInfo.url,             // "/users?$filter=department eq 'IT'&$orderby=displayName asc&$top=20"
  queryParams: debugInfo.queryParams, // { $filter: "department eq 'IT'", $orderby: "displayName asc", $top: 20 }
  headers: debugInfo.headers      // { "Authorization": "Bearer ...", "Content-Type": "application/json" }
});

const users = await usersQuery.get();
```

### 邮件操作调试

```typescript
// 发送邮件
const message = {
  subject: '测试邮件',
  toRecipients: [{ emailAddress: { address: 'test@example.com' } }],
  body: { contentType: 'Text', content: '测试内容' }
};

await orm.mail.send('user-id', message);

// 获取调试信息
const mailDebug = orm.mail.getDebug();
console.log('邮件发送调试:', {
  method: mailDebug.method,        // "POST"
  url: mailDebug.url,             // "/users/user-id/sendMail"
  body: mailDebug.body,           // { message: {...}, saveToSentItems: true }
  headers: mailDebug.headers      // 包含认证和内容类型头
});
```

### 文件操作调试

```typescript
// 上传文件
const buffer = Buffer.from('test content');
await orm.files.uploadSmallFile('user-id', '/test/file.txt', buffer);

// 获取调试信息
const fileDebug = orm.files.getDebug();
console.log('文件上传调试:', {
  method: fileDebug.method,        // "PUT"
  url: fileDebug.url,             // "/users/user-id/drive/root:/test/file.txt:/content"
  body: fileDebug.body,           // Buffer 内容
  headers: fileDebug.headers      // 包含内容类型头
});
```

### 日历操作调试

```typescript
// 创建事件
const event = {
  subject: '测试会议',
  start: { dateTime: '2024-01-15T10:00:00', timeZone: 'Asia/Shanghai' },
  end: { dateTime: '2024-01-15T11:00:00', timeZone: 'Asia/Shanghai' }
};

await orm.calendar.createEvent('user-id', event);

// 获取调试信息
const calendarDebug = orm.calendar.getDebug();
console.log('日历事件调试:', {
  method: calendarDebug.method,    // "POST"
  url: calendarDebug.url,         // "/users/user-id/events"
  body: calendarDebug.body,      // 事件对象
  headers: calendarDebug.headers  // 包含时区偏好头
});
```

## 调试最佳实践

### 1. 生产环境调试包装器

```typescript
async function debugRequest<T>(
  query: any, 
  operation: string,
  execute: () => Promise<T>
): Promise<T> {
  const debugInfo = query.getDebug();
  console.log(`=== ${operation} 调试信息 ===`);
  console.log('请求方法:', debugInfo.method);
  console.log('请求 URL:', debugInfo.url);
  console.log('时间戳:', new Date().toISOString());
  
  try {
    const result = await execute();
    console.log(`${operation} 成功`);
    return result;
  } catch (error) {
    console.error(`${operation} 失败:`, error);
    throw error;
  }
}

// 使用示例
const query = orm.users.query().where('department', 'eq', 'IT');
await debugRequest(query, '查询 IT 部门用户', () => query.get());
```

### 2. 调试信息记录

```typescript
function logDebugInfo(query: any, operation: string) {
  const debugInfo = query.getDebug();
  const logEntry = {
    timestamp: new Date().toISOString(),
    operation,
    method: debugInfo.method,
    url: debugInfo.url,
    queryParams: debugInfo.queryParams,
    headers: debugInfo.headers
  };
  
  // 记录到日志系统
  console.log('调试日志:', JSON.stringify(logEntry, null, 2));
}

// 使用示例
const query = orm.users.query().where('displayName', 'contains', '张');
logDebugInfo(query, 'user_search');
const users = await query.get();
```

### 3. 错误调试

```typescript
try {
  const users = await orm.users.query()
    .where('invalidField', 'eq', 'value')
    .get();
} catch (error) {
  // 获取失败请求的调试信息
  const debugInfo = orm.users.query().getDebug();
  console.error('请求失败调试信息:', {
    method: debugInfo.method,
    url: debugInfo.url,
    queryParams: debugInfo.queryParams,
    error: error.message
  });
}
```

## 注意事项

1. **执行前调试**: `getDebug()` 可以在执行请求前调用，获取请求参数
2. **执行后调试**: 执行请求后，`getDebug()` 返回最后一次请求的调试信息
3. **服务层调试**: 服务层方法执行后，可以通过对应的服务实例获取调试信息
4. **原始参数**: `getRaw()` 返回不包含完整 URL 的原始请求参数
5. **性能考虑**: 在生产环境中，建议只在需要时启用调试功能

## 完整示例

```typescript
import { createGraphORM } from '@qingu-x/msgraph-orm-js';

async function demonstrateDebugging() {
  const orm = createGraphORM(client);
  
  // 1. Repository 调试
  const usersQuery = orm.users.query()
    .where('department', 'eq', 'IT')
    .select(['id', 'displayName'])
    .top(5);
  
  console.log('用户查询调试:', usersQuery.getDebug());
  const users = await usersQuery.get();
  
  // 2. Service 调试
  await orm.mail.send('user-id', {
    subject: '测试',
    toRecipients: [{ emailAddress: { address: 'test@example.com' } }],
    body: { contentType: 'Text', content: '测试内容' }
  });
  console.log('邮件服务调试:', orm.mail.getDebug());
  
  // 3. GraphORM 调试
  await orm.me();
  console.log('GraphORM 调试:', orm.getDebug());
}

demonstrateDebugging();
```

这个调试系统让你能够：
- 在请求执行前查看请求参数
- 在请求执行后查看实际发送的请求信息
- 调试所有类型的操作（查询、创建、更新、删除）
- 监控服务层和 GraphORM 层的请求
- 在生产环境中进行问题排查
