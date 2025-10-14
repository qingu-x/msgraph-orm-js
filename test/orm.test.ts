import { GraphClient } from '../src/client-credentials';
import { defaultEndpoints } from '../src/types';
import { createGraphORM } from '../src/orm';

test('ORM 调试示例', () => {
  const client = new GraphClient({
    tenantId: 'your-tenant-id',
    clientId: 'your-client-id',
    clientSecret: 'your-client-secret'
  }, defaultEndpoints.global);
  
  const orm = createGraphORM(client.getGraphClient());
  
  // 1. Repository 查询调试
  const usersQuery = orm.users.query()
    .where('department', 'eq', 'IT')
    .select(['id', 'displayName', 'mail'])
    .top(10);
  
  // 获取调试信息
  const debugInfo = usersQuery.getDebug();
  console.log(debugInfo);
  
  // 2. 用户关系查询调试
  const messagesQuery = orm.users.messages('user-id')
    .query()
    .where('isRead', 'eq', false)
    .orderBy('receivedDateTime', 'desc')
    .top(5);
  
  const messagesDebugInfo = messagesQuery.getDebug();
  console.log('=== 用户邮件查询调试信息 ===');
  console.log(messagesDebugInfo);
  
  // 3. 带自定义 Headers 的查询调试
  const customQuery = orm.users.query()
    .header('ConsistencyLevel', 'eventual')
    .header('Prefer', 'outlook.timezone="Asia/Shanghai"')
    .where('displayName', 'startswith', '张');
  
  const customDebugInfo = customQuery.getDebug();
  console.log('=== 带自定义 Headers 的查询调试信息 ===');
  console.log(customDebugInfo);
  
  // 4. 服务层调试示例（不执行，只展示调试信息获取）
  console.log('=== 服务层调试示例 ===');
  
  // 邮件服务调试
  console.log('邮件服务调试方法: orm.mail.getDebug()');
  
  // 文件服务调试
  console.log('文件服务调试方法: orm.files.getDebug()');
  
  // 日历服务调试
  console.log('日历服务调试方法: orm.calendar.getDebug()');
  
  // 5. GraphORM 调试
  console.log('GraphORM 调试方法: orm.getDebug()');
  
  expect(1).toBe(1);
});