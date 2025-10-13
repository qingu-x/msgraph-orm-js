import { Client } from '@microsoft/microsoft-graph-client';
import { GraphEntity, EntityManager, GraphCollection, CrudRawRequest } from './types';
import { GraphQueryBuilder } from './query-builder';
import { GraphOrmError } from './errors';

export class GraphEntityManager<T extends GraphEntity> implements EntityManager<T> {
  constructor(
    private client: Client,
    private endpoint: string
  ) {}

  async findById(id: string): Promise<T> {
    try {
      return await this.client.api(`${this.endpoint}/${encodeURIComponent(id)}`).get();
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  async findOne(query: Partial<T>): Promise<T | null> {
    const conditions = this.buildConditionsFromQuery(query);
    const queryBuilder = this.query();
    
    conditions.forEach(condition => {
      queryBuilder.where(condition.field, condition.operator, condition.value);
    });
    
    try {
      const result = await queryBuilder.first();
      return result;
    } catch (error) {
      // 如果是 NO_ENTITY_FOUND 错误，返回 null
      if (error instanceof GraphOrmError && error.code === 'NO_ENTITY_FOUND') {
        return null;
      }
      // 其他错误向上抛出
      throw error;
    }
  }

  async findMany(query?: Partial<T>): Promise<GraphCollection<T>> {
    try {
      const queryBuilder = this.query();
      
      if (query) {
        const conditions = this.buildConditionsFromQuery(query);
        conditions.forEach(condition => {
          queryBuilder.where(condition.field, condition.operator, condition.value);
        });
      }
      const result = await queryBuilder.count().get();
      return result;
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  async create(entity: Omit<T, 'id'>): Promise<T> {
    try {
      return await this.client.api(this.endpoint).post(entity);
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  async update(id: string, entity: Partial<T>): Promise<T> {
    try {
      return await this.client.api(`${this.endpoint}/${encodeURIComponent(id)}`).patch(entity);
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  async delete(id: string): Promise<void> {
    try {
      await this.client.api(`${this.endpoint}/${encodeURIComponent(id)}`).delete();
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  query(): GraphQueryBuilder<T> {
    return new GraphQueryBuilder<T>(this.client, this.endpoint);
  }

  private buildConditionsFromQuery(query: Partial<T>) {
    const conditions: Array<{ field: string; operator: 'eq'; value: string | number | boolean }> = [];
    
    for (const [key, value] of Object.entries(query)) {
      if (value !== undefined && value !== null && key !== 'id') {
        conditions.push({
          field: key,
          operator: 'eq',
          value: value as string | number | boolean
        });
      }
    }
    
    return conditions;
  }

  /**
   * 获取 findById 操作的原始请求信息（用于调试）
   * 不执行实际请求，仅返回请求详情
   */
  getRawFindById(id: string): CrudRawRequest {
    const url = `${this.endpoint}/${encodeURIComponent(id)}`;
    return {
      method: 'GET',
      endpoint: this.endpoint,
      url,
    };
  }

  /**
   * 获取 create 操作的原始请求信息（用于调试）
   * 不执行实际请求，仅返回请求详情
   */
  getRawCreate(entity: Omit<T, 'id'>): CrudRawRequest {
    return {
      method: 'POST',
      endpoint: this.endpoint,
      url: this.endpoint,
      body: entity,
      headers: {
        'Content-Type': 'application/json'
      }
    };
  }

  /**
   * 获取 update 操作的原始请求信息（用于调试）
   * 不执行实际请求，仅返回请求详情
   */
  getRawUpdate(id: string, entity: Partial<T>): CrudRawRequest {
    const url = `${this.endpoint}/${encodeURIComponent(id)}`;
    return {
      method: 'PATCH',
      endpoint: this.endpoint,
      url,
      body: entity,
      headers: {
        'Content-Type': 'application/json'
      }
    };
  }

  /**
   * 获取 delete 操作的原始请求信息（用于调试）
   * 不执行实际请求，仅返回请求详情
   */
  getRawDelete(id: string): CrudRawRequest {
    const url = `${this.endpoint}/${encodeURIComponent(id)}`;
    return {
      method: 'DELETE',
      endpoint: this.endpoint,
      url,
    };
  }
}
