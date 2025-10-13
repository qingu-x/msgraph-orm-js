import { Client } from '@microsoft/microsoft-graph-client';
import { GraphEntity, EntityManager, GraphCollection, QueryBuilder } from './types';
import { GraphQueryBuilder } from './query-builder';
import { GraphOrmError } from './errors';

/**
 * Graph 仓储基类
 * 
 * 提供标准的 CRUD 操作和查询构建能力
 * 所有资源仓储都应继承此类
 */
export abstract class GraphRepository<T extends GraphEntity> implements EntityManager<T> {
  constructor(
    protected client: Client,
    protected endpoint: string
  ) {}

  /**
   * 通过 ID 查找实体
   */
  async findById(id: string): Promise<T> {
    try {
      return await this.client.api(`${this.endpoint}/${encodeURIComponent(id)}`).get();
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 查找单个实体
   */
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
      if (error instanceof GraphOrmError && error.code === 'NO_ENTITY_FOUND') {
        return null;
      }
      throw error;
    }
  }

  /**
   * 查找多个实体
   */
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

  /**
   * 获取所有实体（带分页）
   */
  async all(): Promise<GraphCollection<T>> {
    return this.findMany();
  }

  /**
   * 创建实体
   */
  async create(entity: Omit<T, 'id'>): Promise<T> {
    try {
      return await this.client.api(this.endpoint).post(entity);
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 更新实体
   */
  async update(id: string, entity: Partial<T>): Promise<T> {
    try {
      return await this.client.api(`${this.endpoint}/${encodeURIComponent(id)}`).patch(entity);
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 删除实体
   */
  async delete(id: string): Promise<void> {
    try {
      await this.client.api(`${this.endpoint}/${encodeURIComponent(id)}`).delete();
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 检查实体是否存在
   */
  async exists(id: string): Promise<boolean> {
    try {
      await this.findById(id);
      return true;
    } catch (error) {
      if (error instanceof GraphOrmError && error.code === 'RESOURCE_NOT_FOUND') {
        return false;
      }
      throw error;
    }
  }

  /**
   * 计数
   */
  async count(query?: Partial<T>): Promise<number> {
    const result = await this.findMany(query);
    return result.meta.count || result.data.length;
  }

  /**
   * 创建查询构建器
   */
  query(): GraphQueryBuilder<T> {
    return new GraphQueryBuilder<T>(this.client, this.endpoint);
  }

  /**
   * 构建查询条件
   */
  protected buildConditionsFromQuery(query: Partial<T>) {
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
   * 获取子资源仓储
   * 用于导航属性访问
   */
  protected subRepository<R extends GraphEntity>(
    id: string,
    navigationProperty: string
  ): GraphRepository<R> {
    const subEndpoint = `${this.endpoint}/${encodeURIComponent(id)}/${navigationProperty}`;
    return new GenericRepository<R>(this.client, subEndpoint);
  }
}

/**
 * 通用仓储实现
 * 用于没有特殊逻辑的资源
 */
export class GenericRepository<T extends GraphEntity> extends GraphRepository<T> {
  constructor(client: Client, endpoint: string) {
    super(client, endpoint);
  }
}

