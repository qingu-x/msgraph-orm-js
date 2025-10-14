import { Client, GraphRequest } from '@microsoft/microsoft-graph-client';
import { GraphCollection, QueryBuilder, QueryCondition, OrderDirection, QueryOperator, RequestDebugInfo } from './types';
import { GraphOrmError, GraphErrorCode } from './errors';

/**
 * Microsoft Graph 查询构建器
 * 
 * 重要限制说明：
 * 1. 高级查询操作符 (startswith, endswith, contains) 需要 ConsistencyLevel: eventual
 * 2. 使用高级操作符时，排序功能通常不可用
 * 3. ne (不等于) 操作符在某些资源类型上不支持（如 User）
 * 4. $count 和 $search 需要 ConsistencyLevel: eventual
 * 5. 并非所有属性都支持所有操作符
 * 
 * @see https://learn.microsoft.com/graph/query-parameters
 * @see https://learn.microsoft.com/graph/aad-advanced-queries
 */
export class GraphQueryBuilder<T = unknown> implements QueryBuilder<T> {
  private conditions: QueryCondition[] = [];
  private orderByField?: string;
  private orderDirection: OrderDirection = 'asc';
  private selectFields?: string[];
  private topCount?: number;
  private skipCount?: number;
  private skipTokenValue?: string;
  private searchQuery?: string;
  private countOnly = false;
  private expandProperties?: string[];
  private responseFormat?: 'json' | 'atom';
  private customHeaders: Record<string, string> = {};
  
  private debugInfo?: RequestDebugInfo;

  // 高级查询操作符（需要 ConsistencyLevel: eventual，不支持排序）
  private readonly advancedOperators: QueryOperator[] = ['startswith', 'endswith', 'contains'];
  
  // 有问题的操作符（在某些资源上不支持或有限制）
  private readonly problematicOperators: QueryOperator[] = ['ne', 'not'];

  constructor(
    private client: Client,
    private endpoint: string
  ) {}

  /**
   * 检测当前端点是否为 User 资源
   */
  private isUserResource(): boolean {
    return this.endpoint.includes('/users');
  }

  /**
   * 检测当前端点是否为 Device 资源
   */
  private isDeviceResource(): boolean {
    return this.endpoint.includes('/devices');
  }

  /**
   * 检测当前资源是否支持 ne 操作符
   * User 和 Device 资源不支持 ne 操作符
   */
  private supportsNeOperator(): boolean {
    return !this.isUserResource() && !this.isDeviceResource();
  }

  private generateCondition(field: string, operator: QueryOperator | string | number | boolean | null, value?: string | number | boolean | null, logicalOperator?: 'and' | 'or') {
    if (value === undefined) {
      this.conditions.push({ field, operator: 'eq', value: operator as string | number | boolean | null, logicalOperator: logicalOperator });
    } else {
      const actualOperator = operator as QueryOperator;
      // 检查资源类型是否支持该操作符
      if (actualOperator === 'ne' && !this.supportsNeOperator()) {
        const resourceType = this.isUserResource() ? 'User' : 
                            this.isDeviceResource() ? 'Device' : '当前';
        throw GraphOrmError.create(
          GraphErrorCode.INVALID_PARAMETER,
          `${resourceType} 资源不支持 'ne' (不等于) 操作符。建议：使用其他操作符或在客户端过滤数据。`
        );
      }
      this.conditions.push({ field, operator: actualOperator, value, logicalOperator: logicalOperator });
    }
  }

  where(field: string, operator: QueryOperator | string | number | boolean | null, value?: string | number | boolean | null): QueryBuilder<T> {
    this.generateCondition(field, operator, value);
    return this;
  }

  and(field: string, operator: QueryOperator | string | number | boolean | null, value?: string | number | boolean | null): QueryBuilder<T> {
    this.generateCondition(field, operator, value, 'and');
    return this;
  }

  or(field: string, operator: QueryOperator | string | number | boolean | null, value?: string | number | boolean | null): QueryBuilder<T> {
    this.generateCondition(field, operator, value, 'or');
    return this;
  }

  orderBy(field: string, direction: OrderDirection = 'asc'): QueryBuilder<T> {
    // 检查当前查询是否支持排序
    // 1. 如果已经使用了高级操作符，不能排序
    if (this.hasAdvancedOperators()) {
      throw GraphOrmError.create(
        GraphErrorCode.INVALID_PARAMETER,
        '使用高级查询操作符 (startswith/endswith/contains) 时不支持排序。请移除 orderBy 或改用其他操作符。'
      );
    }
    
    // 2. 如果是 User/Device 资源且使用了 ne/not 操作符，不能排序
    if ((this.isUserResource() || this.isDeviceResource()) && this.hasProblematicOperators()) {
      const resourceType = this.isUserResource() ? 'User' : 'Device';
      throw GraphOrmError.create(
        GraphErrorCode.INVALID_PARAMETER,
        `${resourceType} 资源使用 'ne' 或 'not' 操作符时不支持排序。请移除 orderBy 或改用其他操作符。`
      );
    }
    
    this.orderByField = field;
    this.orderDirection = direction;
    return this;
  }

  select(fields: string[]): QueryBuilder<T> {
    this.selectFields = fields;
    return this;
  }

  top(count: number): QueryBuilder<T> {
    this.topCount = count;
    return this;
  }

  skip(count: number): QueryBuilder<T> {
    this.skipCount = count;
    return this;
  }

  /**
   * 设置分页令牌（用于获取下一页数据）
   * 通常从上一次查询响应的 @odata.nextLink 中提取
   * 
   * @param token - 分页令牌
   * @see https://learn.microsoft.com/graph/paging
   */
  skipToken(token: string): QueryBuilder<T> {
    this.skipTokenValue = token;
    return this;
  }

  search(field: string, value: string): QueryBuilder<T> {
    this.searchQuery = `${field}:${value}`;
    return this;
  }

  /**
   * 展开导航属性（$expand）
   * 允许在单个请求中检索相关资源
   * 
   * 示例：
   * - expand('manager') - 展开 manager 属性
   * - expand(['memberOf', 'directReports']) - 展开多个属性
   * - expand('manager($select=displayName,mail)') - 展开并选择特定字段
   * 
   * @param property - 要展开的属性名称或属性数组
   * @see https://learn.microsoft.com/graph/query-parameters#expand-parameter
   */
  expand(property: string | string[]): QueryBuilder<T> {
    if (Array.isArray(property)) {
      this.expandProperties = property;
    } else {
      this.expandProperties = [property];
    }
    return this;
  }

  /**
   * 设置响应格式（$format）
   * 默认为 json
   * 
   * @param format - 响应格式（json 或 atom）
   * @see https://learn.microsoft.com/graph/query-parameters#format-parameter
   */
  format(format: 'json' | 'atom'): QueryBuilder<T> {
    this.responseFormat = format;
    return this;
  }

  count(): QueryBuilder<T> {
    this.countOnly = true;
    return this;
  }

  /**
   * 设置自定义请求头
   * 
   * @param name - 请求头名称
   * @param value - 请求头值
   * @returns QueryBuilder 实例（支持链式调用）
   * @example
   * ```typescript
   * // 设置自定义请求头
   * queryBuilder.header('Prefer', 'outlook.timezone="Pacific Standard Time"')
   * ```
   */
  addHeader(name: string, value: string): QueryBuilder<T> {
    this.customHeaders[name] = value;
    return this;
  }

  setHeaders(headers: Record<string, string>) {
    this.customHeaders = headers;
    return this;
  }

  /**
   * 构建 OData $filter 查询字符串
   * @see https://learn.microsoft.com/graph/filter-query-parameter
   */
  private buildFilter(): string {
    // 过滤掉 contains（使用 $search）和有问题的操作符（如 ne）
    const filterConditions = this.conditions.filter(condition => {
      // contains 使用 $search 而不是 $filter
      if (condition.operator === 'contains') {
        return false;
      }
      
      // 注意：ne 操作符的检查已经在 where/and/or 方法中进行
      // 这里只是作为后备检查，理论上不应该执行到
      if (condition.operator === 'ne' && !this.supportsNeOperator()) {
        return false; // 跳过该条件
      }
      
      return true;
    });
    
    // 如果没有过滤条件，返回空字符串
    if (filterConditions.length === 0) {
      return "";
    }
    
    return filterConditions.map((condition, index) => {
      const { field, operator, value, logicalOperator } = condition;
      
      // 构建条件表达式
      let conditionStr = '';
      
      switch (operator) {
        case 'eq':
          // 相等比较，支持所有资源类型
          conditionStr = value === null ? `${field} eq null` : `${field} eq '${value}'`;
          break;
          
        case 'ne':
          // 不等于，在某些资源上不支持（如 User）
          // 注意：ne 操作符已在 buildFilter 开头被完全过滤
          // 这个分支理论上不应该被执行到
          if (value === null) {
            conditionStr = `${field} ne null`;
          } else {
            conditionStr = `${field} ne '${value}'`;
          }
          break;
          
        case 'gt':
          // 大于，适用于数值和日期
          conditionStr = `${field} gt ${value}`;
          break;
          
        case 'ge':
          // 大于等于，适用于数值和日期
          conditionStr = `${field} ge ${value}`;
          break;
          
        case 'lt':
          // 小于，适用于数值和日期
          conditionStr = `${field} lt ${value}`;
          break;
          
        case 'le':
          // 小于等于，适用于数值和日期
          conditionStr = `${field} le ${value}`;
          break;
          
        case 'startswith':
          // 字符串开头匹配（需要 ConsistencyLevel: eventual）
          conditionStr = `startswith(${field},'${value}')`;
          break;
          
        case 'endswith':
          // 字符串结尾匹配（需要 ConsistencyLevel: eventual）
          conditionStr = `endswith(${field},'${value}')`;
          break;
          
        case 'not':
          // 逻辑非
          conditionStr = `not ${field} eq '${value}'`;
          break;
          
        default:
          // 默认使用相等比较
          conditionStr = `${field} eq '${value}'`;
      }
      
      // 如果不是第一个条件且有逻辑操作符，添加逻辑操作符
      if (index > 0 && logicalOperator) {
        return `${logicalOperator} ${conditionStr}`;
      }
      
      return conditionStr;
    }).join(' ');
  }

  /**
   * 检查是否使用了高级查询操作符
   */
  private hasAdvancedOperators(): boolean {
    return this.conditions.some(condition => 
      this.advancedOperators.includes(condition.operator as QueryOperator)
    );
  }

  /**
   * 检查是否使用了有问题的操作符
   */
  private hasProblematicOperators(): boolean {
    return this.conditions.some(condition => 
      this.problematicOperators.includes(condition.operator as QueryOperator)
    );
  }

  /**
   * 构建 $search 查询字符串
   * @see https://learn.microsoft.com/graph/search-query-parameter
   */
  private buildSearchQuery(): string {
    const containsConditions = this.conditions.filter(condition => condition.operator === 'contains');
    if (containsConditions.length === 0) return '';
    
    return containsConditions.map(condition => {
      const { field, value } = condition;
      return `${field}:${value}`;
    }).join(' ');
  }

  /**
   * 获取查询的调试信息
   * 包括请求方法、端点、查询参数和请求头
   * 
   * 支持的场景：
   * 1. 查询操作（GET）- 返回查询参数和过滤条件
   * 2. CRUD 操作 - 返回最后一次执行的 CRUD 操作信息
   * 
   * @returns 包含完整请求信息的调试对象
   * @example
   * ```typescript
   * // 查询操作
   * const users = await queryBuilder.where('displayName', 'eq', 'John').get();
   * const debug = queryBuilder.getDebug();
   * // { method: 'GET', endpoint: '/users', params: {...}, headers: {...} }
   * 
   * // CRUD 操作
   * const user = await queryBuilder.create({ displayName: 'John' });
   * const debug = queryBuilder.getDebug();
   * // { method: 'POST', endpoint: '/users', body: {...}, headers: {...} }
   * ```
   */
  public getDebug(): RequestDebugInfo | undefined {
    // 否则返回查询操作的调试信息
    return this.debugInfo;
  }

  /**
   * 构建查询参数
   * 组合所有 OData 查询选项
   * @see https://learn.microsoft.com/graph/query-parameters
   */
  private buildQueryParams(): Record<string, string | number> {
    const params: Record<string, string | number> = {};
    
    // 1. 处理搜索查询 ($search)
    // 需要 ConsistencyLevel: eventual
    const searchQuery = this.searchQuery || this.buildSearchQuery();
    if (searchQuery) {
      params.$search = `"${searchQuery}"`;
    }
    
    // 2. 处理过滤条件 ($filter)
    const filter = this.buildFilter();
    if (filter) {
      params.$filter = filter;
    }
    
    // 3. 处理排序 ($orderby)
    // 注意：排序兼容性检查已在 orderBy() 方法中进行
    // 如果到这里还有 orderByField，说明是兼容的
    if (this.orderByField) {
      params.$orderby = `${this.orderByField} ${this.orderDirection}`;
    }
    
    // 4. 处理字段选择 ($select)
    // 建议：只选择需要的字段以提高性能
    if (this.selectFields && this.selectFields.length > 0) {
      params.$select = this.selectFields.join(',');
    }
    
    // 5. 处理分页 ($top 和 $skip)
    // $top: 返回结果的最大数量（默认最大值因资源而异）
    if (this.topCount) {
      params.$top = this.topCount;
    }
    
    // $skip: 跳过的结果数量（用于分页）
    if (this.skipCount) {
      params.$skip = this.skipCount;
    }
    
    // $skiptoken: 分页令牌（用于获取下一页）
    if (this.skipTokenValue) {
      params.$skiptoken = this.skipTokenValue;
    }
    
    // 6. 处理展开导航属性 ($expand)
    // 允许在单个请求中获取相关资源
    if (this.expandProperties && this.expandProperties.length > 0) {
      params.$expand = this.expandProperties.join(',');
    }
    
    // 7. 处理响应格式 ($format)
    // 默认为 json
    if (this.responseFormat) {
      params.$format = this.responseFormat;
    }
    
    // 8. 处理计数 ($count)
    // 需要 ConsistencyLevel: eventual
    if (this.countOnly) {
      params.$count = 'true';
    }
    
    return params;
  }
  private buildRequest(): GraphRequest {
    try {
      const params = this.buildQueryParams();
      const headers = { ...this.customHeaders };
      let request = this.client.api(this.endpoint);
      // 检查是否需要添加 ConsistencyLevel: eventual 头部
      // 以下情况需要：
      // 1. 使用 $search
      // 2. 使用 $count
      // 3. 使用高级查询操作符（startswith, endswith, contains）
      const needsEventualConsistency = 
        this.searchQuery ||
        this.buildSearchQuery() ||
        this.countOnly ||
        this.hasAdvancedOperators();
      
      if (needsEventualConsistency) {
        headers['ConsistencyLevel'] = 'eventual';
      }
      
      // 添加自定义请求头
      request.headers(headers);
      
      // 添加查询参数
      if (Object.keys(params).length > 0) {
        request = request.query(params);
      }
      this.debugInfo = {
        method: 'GET',
        params,
        headers,
        url: this.endpoint
      };
      
      return request;
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }
  /**
   * 执行查询并返回结果集合
   * 
   * 重要说明：
   * 1. 使用高级查询功能（$search, $count, 高级操作符）时会自动添加 ConsistencyLevel: eventual
   * 2. ConsistencyLevel: eventual 意味着最终一致性，可能会有轻微延迟
   * 
   * @returns 包含数据和元信息的集合
   * @throws GraphOrmError 查询失败时抛出
   * @see https://learn.microsoft.com/graph/aad-advanced-queries
   */
  async get(): Promise<GraphCollection<T>> {
    const request = this.buildRequest();
    this.debugInfo!.method = 'GET';
    const response = await request.get();
    this.debugInfo!.response = response;
    return {
      meta: {
        nextLink: response['@odata.nextLink'],
        count: response['@odata.count'],
      },
      data: response.value || []
    };
  }

  /**
   * 执行查询并返回单个实体
   * 如果没有找到结果，会抛出错误
   * 
   * @returns 第一个匹配的实体
   * @throws GraphOrmError 没有找到结果时抛出 NO_ENTITY_FOUND 错误
   */
  async first(): Promise<T | null> {
    const response = await this.get();
    const result =  response.data.length > 0 ? response.data[0] : null;
    this.debugInfo!.response = result;
    return result;
  } 

  /**
   * 执行查询并自动处理分页，返回异步迭代器
   * 
   * 使用示例：
   * ```typescript
   * for await (const user of orm.users.query().pagination()) {
   *   console.log(user.displayName);
   * }
   * ```
   * 
   * 此方法会自动处理 @odata.nextLink，无需手动管理分页
   * 
   * @returns 异步迭代器，逐个返回实体
   * @see https://learn.microsoft.com/graph/paging
   */
  async *pagination(): AsyncIterableIterator<T> {
    let nextLink: string | undefined;
    let isFirstPage = true;
    
    do {
      try {
        let result: GraphCollection<T>;
        
        if (isFirstPage) {
          // 第一页：使用当前的查询参数
          result = await this.get() as GraphCollection<T>;
          isFirstPage = false;
        } else if (nextLink) {
          // 后续页：使用 nextLink
          this.endpoint = nextLink;
          result = await this.get() as GraphCollection<T>;
        } else {
          // 没有更多数据
          break;
        }
        
        // 逐个返回当前页的数据
        for (const item of result.data) {
          yield item;
        }
        
        // 更新 nextLink 以获取下一页
        nextLink = result.meta.nextLink;
        
      } catch (error) {
        throw GraphOrmError.fromGraphError(error);
      }
    } while (nextLink);
  }

  /**
   * 通过 ID 查找单个实体
   * 
   * @param id - 实体 ID
   * @returns 实体对象
   * @throws GraphOrmError 查询失败时抛出
   * @example
   * ```typescript
   * const user = await queryBuilder.findById('user-id-123');
   * const debug = queryBuilder.getDebug();
   * console.log(debug); // { method: 'GET', endpoint: '/users/user-id-123', ... }
   * ```
   */
  async findById(id: string): Promise<T | null> {
    try {
      this.endpoint = `${this.endpoint}/${encodeURIComponent(id)}`;
      const request = this.buildRequest();
      this.debugInfo!.method = 'GET';
      const response = await request.get();
      this.debugInfo!.response = response;
      return response;
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 创建新实体
   * 
   * @param entity - 要创建的实体数据（不含 id）
   * @returns 创建后的实体对象（包含服务器生成的 id）
   * @throws GraphOrmError 创建失败时抛出
   * @example
   * ```typescript
   * const newUser = await queryBuilder.create({
   *   displayName: 'John Doe',
   *   userPrincipalName: 'john@contoso.com'
   * });
   * const debug = queryBuilder.getDebug();
   * console.log(debug); // { method: 'POST', endpoint: '/users', body: {...}, ... }
   * ```
   */
  async create(entity: Omit<T, 'id'>): Promise<T> {
    try {
      const request = this.buildRequest();
      this.debugInfo!.method = 'POST';
      this.debugInfo!.payload = entity;
      const response = await request.post(entity);
      this.debugInfo!.response = response;
      return response;
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 更新实体
   * 
   * @param id - 实体 ID
   * @param entity - 要更新的字段（部分更新）
   * @returns 更新后的实体对象
   * @throws GraphOrmError 更新失败时抛出
   * @example
   * ```typescript
   * const updatedUser = await queryBuilder.update('user-id-123', {
   *   displayName: 'Jane Doe'
   * });
   * const debug = queryBuilder.getDebug();
   * console.log(debug); // { method: 'PATCH', endpoint: '/users/user-id-123', body: {...}, ... }
   * ```
   */
  async update(id: string | null, entity: Partial<T>): Promise<T> {
    try {
      if(id) {
        this.endpoint = `${this.endpoint}/${encodeURIComponent(id)}`;
      }
      const request = this.buildRequest();
      this.debugInfo!.method = 'PATCH';
      this.debugInfo!.payload = entity;
      const response = await request.patch(entity);
      this.debugInfo!.response = response;
      return response;
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 删除实体
   * 
   * @param id - 实体 ID
   * @throws GraphOrmError 删除失败时抛出
   * @example
   * ```typescript
   * await queryBuilder.delete('user-id-123');
   * const debug = queryBuilder.getDebug();
   * console.log(debug); // { method: 'DELETE', endpoint: '/users/user-id-123', ... }
   * ```
   */
  async delete(id: string): Promise<void> {
    try {
      this.endpoint = `${this.endpoint}/${encodeURIComponent(id)}`;
      const request = this.buildRequest();
      this.debugInfo!.method = 'DELETE';
      const response = await request.delete();
      this.debugInfo!.response = response;
      return response;
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 执行 PUT 请求（用于文件上传等场景）
   * 
   * @param path - 相对于当前 endpoint 的路径（可以为空字符串）
   * @param body - 请求体（可以是对象、ArrayBuffer、Blob、string 等）
   * @returns 响应结果
   * @throws GraphOrmError 请求失败时抛出
   * @example
   * ```typescript
   * // 上传文件
   * const result = await queryBuilder.put('root:/test.txt:/content', fileContent);
   * 
   * // 更新实体（带 ID）
   * const result = await queryBuilder.put('user-id-123', { displayName: 'New Name' });
   * ```
   */
  async put<TResult = unknown>(path: string, body: unknown): Promise<TResult> {
    try {
      this.endpoint = path ? `${this.endpoint}/${path}` : this.endpoint;
      const request = this.buildRequest();
      this.debugInfo!.method = 'PUT';
      this.debugInfo!.payload = body;
      const response = await request.put(body);
      this.debugInfo!.response = response;
      return response;
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 执行 POST 请求（用于创建资源、执行操作等场景）
   * 
   * @param path - 相对于当前 endpoint 的路径
   * @param body - 请求体（可以是对象、ArrayBuffer、Blob、string 等）
   * @param id - 实体 ID（可选，用于构建完整端点路径）
   * @returns 响应结果
   * @throws GraphOrmError 请求失败时抛出
   * @example
   * ```typescript
   * // 执行操作
   * await queryBuilder.post('sendMail', { message: {...} });
   * 
   * // 在特定实体上执行操作
   * await queryBuilder.post('move', { destinationId: 'folder-id' }, 'message-id');
   * 
   * // 创建子资源
   * const attachment = await queryBuilder.post<Attachment>('attachments', attachmentData, 'message-id');
   * ```
   */
  async post<TResult = unknown>(path: string, body?: unknown, id?: string): Promise<TResult> {
    try {
      // 构建完整端点路径
      if (id) {
        this.endpoint = `${this.endpoint}/${encodeURIComponent(id)}/${path}`;
      } else {
        this.endpoint = `${this.endpoint}/${path}`;
      }
      
      const request = this.buildRequest();
      this.debugInfo!.method = 'POST';
      this.debugInfo!.payload = body;
      const response = await request.post(body);
      this.debugInfo!.response = response;
      return response;
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

}
