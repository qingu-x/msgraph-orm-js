/**
 * 增强的类型定义
 * 
 * 支持关系映射和延迟加载
 */

import { GraphEntity, User, Group, Message, Event, DriveItem } from './types';

/**
 * 关系加载选项
 */
export interface RelationLoadOptions {
  /**
   * 要加载的关系名称
   */
  relations: string[];
  
  /**
   * 是否递归加载嵌套关系
   */
  recursive?: boolean;
  
  /**
   * 最大递归深度
   */
  maxDepth?: number;
}

/**
 * 带关系的实体基类
 */
export interface EntityWithRelations<T extends GraphEntity> {
  /**
   * 实体数据
   */
  data: T;
  
  /**
   * 加载的关系
   */
  relations?: Record<string, unknown>;
  
  /**
   * 是否已加载指定关系
   */
  hasRelation(name: string): boolean;
  
  /**
   * 获取关系数据
   */
  getRelation<R>(name: string): R | undefined;
}

/**
 * 增强的用户类型（包含关系）
 */
export interface UserWithRelations extends User {
  /**
   * 用户所属的组
   * 需要显式加载：orm.users.query().with('groups').get()
   */
  groups?: Group[];
  
  /**
   * 用户的经理
   * 需要显式加载：orm.users.query().with('manager').get()
   */
  manager?: User;
  
  /**
   * 用户的直接下属
   * 需要显式加载：orm.users.query().with('directReports').get()
   */
  directReports?: User[];
  
  /**
   * 用户的邮件
   * 需要显式加载：orm.users.query().with('messages').get()
   */
  messages?: Message[];
  
  /**
   * 用户的事件
   * 需要显式加载：orm.users.query().with('events').get()
   */
  events?: Event[];
  
  /**
   * 用户的驱动器项
   * 需要显式加载：orm.users.query().with('driveItems').get()
   */
  driveItems?: DriveItem[];
}

/**
 * 增强的组类型（包含关系）
 */
export interface GroupWithRelations extends Group {
  /**
   * 组成员
   * 需要显式加载：orm.groups.query().with('members').get()
   */
  members?: User[];
  
  /**
   * 组所有者
   * 需要显式加载：orm.groups.query().with('owners').get()
   */
  owners?: User[];
  
  /**
   * 组事件
   * 需要显式加载：orm.groups.query().with('events').get()
   */
  events?: Event[];
  
  /**
   * 组驱动器项
   * 需要显式加载：orm.groups.query().with('driveItems').get()
   */
  driveItems?: DriveItem[];
}

/**
 * 关系映射配置
 */
export interface RelationMapping {
  /**
   * 关系名称
   */
  name: string;
  
  /**
   * 关系类型
   */
  type: 'one-to-one' | 'one-to-many' | 'many-to-many';
  
  /**
   * 目标实体类型
   */
  target: string;
  
  /**
   * API 端点路径
   */
  endpoint: string;
  
  /**
   * 是否自动加载
   */
  eager?: boolean;
}

/**
 * 实体元数据
 */
export interface EntityMetadata {
  /**
   * 实体名称
   */
  name: string;
  
  /**
   * API 端点
   */
  endpoint: string;
  
  /**
   * 关系映射
   */
  relations: RelationMapping[];
  
  /**
   * 可选字段
   */
  selectableFields?: string[];
  
  /**
   * 可过滤字段
   */
  filterableFields?: string[];
  
  /**
   * 可排序字段
   */
  orderableFields?: string[];
  
  /**
   * 可搜索字段
   */
  searchableFields?: string[];
}

/**
 * 用户实体元数据
 */
export const UserMetadata: EntityMetadata = {
  name: 'User',
  endpoint: '/users',
  relations: [
    {
      name: 'groups',
      type: 'many-to-many',
      target: 'Group',
      endpoint: 'memberOf'
    },
    {
      name: 'manager',
      type: 'one-to-one',
      target: 'User',
      endpoint: 'manager'
    },
    {
      name: 'directReports',
      type: 'one-to-many',
      target: 'User',
      endpoint: 'directReports'
    },
    {
      name: 'messages',
      type: 'one-to-many',
      target: 'Message',
      endpoint: 'messages'
    },
    {
      name: 'events',
      type: 'one-to-many',
      target: 'Event',
      endpoint: 'events'
    },
    {
      name: 'driveItems',
      type: 'one-to-many',
      target: 'DriveItem',
      endpoint: 'drive/root/children'
    }
  ],
  selectableFields: ['id', 'displayName', 'mail', 'userPrincipalName', 'jobTitle', 'department'],
  filterableFields: ['displayName', 'mail', 'userPrincipalName', 'department', 'jobTitle'],
  orderableFields: ['displayName', 'mail', 'userPrincipalName', 'createdDateTime'],
  searchableFields: ['displayName', 'mail', 'userPrincipalName']
};

/**
 * 组实体元数据
 */
export const GroupMetadata: EntityMetadata = {
  name: 'Group',
  endpoint: '/groups',
  relations: [
    {
      name: 'members',
      type: 'many-to-many',
      target: 'User',
      endpoint: 'members'
    },
    {
      name: 'owners',
      type: 'many-to-many',
      target: 'User',
      endpoint: 'owners'
    },
    {
      name: 'events',
      type: 'one-to-many',
      target: 'Event',
      endpoint: 'events'
    },
    {
      name: 'driveItems',
      type: 'one-to-many',
      target: 'DriveItem',
      endpoint: 'drive/root/children'
    }
  ],
  selectableFields: ['id', 'displayName', 'mail', 'description', 'groupTypes'],
  filterableFields: ['displayName', 'mail', 'mailNickname', 'securityEnabled'],
  orderableFields: ['displayName', 'mail', 'createdDateTime'],
  searchableFields: ['displayName', 'mail', 'description']
};

/**
 * 查询选项（支持关系加载）
 */
export interface QueryOptions {
  /**
   * 选择字段
   */
  select?: string[];
  
  /**
   * 过滤条件
   */
  filter?: string;
  
  /**
   * 排序
   */
  orderBy?: string;
  
  /**
   * 限制数量
   */
  top?: number;
  
  /**
   * 跳过数量
   */
  skip?: number;
  
  /**
   * 扩展属性（Graph API 的 $expand）
   */
  expand?: string[];
  
  /**
   * 关系加载（ORM 层面的关系加载）
   */
  with?: string[];
}

/**
 * 仓储接口（支持关系加载）
 */
export interface RepositoryWithRelations<T extends GraphEntity> {
  /**
   * 查询并加载关系
   */
  findWithRelations(id: string, relations: string[]): Promise<EntityWithRelations<T>>;
  
  /**
   * 批量查询并加载关系
   */
  findManyWithRelations(query: Partial<T>, relations: string[]): Promise<Array<EntityWithRelations<T>>>;
}

