import { Client } from '@microsoft/microsoft-graph-client';
import { GraphRepository } from '../repository';
import { Group, User, Team, Drive, Event } from '../types';
import { GraphOrmError } from '../errors';

/**
 * 组仓储
 * 
 * 权限要求：Group.Read.All, Group.ReadWrite.All, Directory.Read.All
 * 国家云支持：✓ 全球版 ✓ 中国版 ✓ 美国政府版
 * 
 * @see https://learn.microsoft.com/graph/api/resources/group
 */
export class GroupRepository extends GraphRepository<Group> {
  constructor(client: Client) {
    super(client, '/groups');
  }

  /**
   * 通过邮箱查找组
   */
  async findByEmail(email: string): Promise<Group | null> {
    return this.query()
      .where('mail', 'eq', email)
      .first()
      .catch(() => null);
  }

  /**
   * 搜索组
   */
  async search(keyword: string) {
    return this.query()
      .search('displayName', keyword)
      .get();
  }

  /**
   * 获取组成员
   */
  members(groupId: string) {
    return this.subRepository<User>(groupId, 'members');
  }

  /**
   * 获取组所有者
   */
  owners(groupId: string) {
    return this.subRepository<User>(groupId, 'owners');
  }

  /**
   * 添加成员到组
   * 
   * 权限要求：GroupMember.ReadWrite.All, Group.ReadWrite.All, Directory.ReadWrite.All
   * 使用 query-builder 支持调试和自定义 header
   */
  async addMember(groupId: string, userId: string): Promise<void> {
    return this.query().post('members/$ref', {
      '@odata.id': `https://graph.microsoft.com/v1.0/directoryObjects/${userId}`
    }, groupId);
  }

  /**
   * 从组中移除成员
   * 
   * 权限要求：GroupMember.ReadWrite.All, Group.ReadWrite.All, Directory.ReadWrite.All
   * 使用 query-builder 支持调试和自定义 header
   */
  async removeMember(groupId: string, userId: string): Promise<void> {
    return this.query().delete(`${encodeURIComponent(groupId)}/members/${encodeURIComponent(userId)}/$ref`);
  }

  /**
   * 添加所有者到组
   * 
   * 权限要求：Group.ReadWrite.All, Directory.ReadWrite.All
   * 使用 query-builder 支持调试和自定义 header
   */
  async addOwner(groupId: string, userId: string): Promise<void> {
    return this.query().post('owners/$ref', {
      '@odata.id': `https://graph.microsoft.com/v1.0/directoryObjects/${userId}`
    }, groupId);
  }

  /**
   * 从组中移除所有者
   * 
   * 权限要求：Group.ReadWrite.All, Directory.ReadWrite.All
   * 使用 query-builder 支持调试和自定义 header
   */
  async removeOwner(groupId: string, userId: string): Promise<void> {
    return this.query().delete(
      `${encodeURIComponent(groupId)}/owners/${encodeURIComponent(userId)}/$ref`
    );
  }

  /**
   * 获取组关联的团队
   * 
   * 权限要求：Team.ReadBasic.All, TeamSettings.Read.All
   * 使用 query-builder 支持调试和自定义 header
   */
  async getTeam(groupId: string): Promise<Team | null> {
    try {
      return await this.query(`${encodeURIComponent(groupId)}/team`).first();
    } catch (error) {
      if (error instanceof GraphOrmError && error.code === 'RESOURCE_NOT_FOUND') {
        return null;
      }
      throw error;
    }
  }

  /**
   * 获取组驱动器
   * 
   * 使用 query-builder 支持调试和自定义 header
   */
  async getDrive(groupId: string): Promise<Drive | null> {
    return await this.query(`${encodeURIComponent(groupId)}/drive`).first();
  }

  /**
   * 获取组事件
   */
  events(groupId: string) {
    return this.subRepository<Event>(groupId, 'events');
  }

  /**
   * 获取组日历事件
   */
  calendarEvents(groupId: string) {
    return this.subRepository<Event>(groupId, 'calendar/events');
  }

  /**
   * 检查成员身份
   * 
   * 使用 query-builder 支持调试和自定义 header
   */
  async isMember(groupId: string, userId: string): Promise<boolean> {
    try {
      await this.query(`${encodeURIComponent(groupId)}/members/${encodeURIComponent(userId)}`).get();
      return true;
    } catch (error) {
      if (error instanceof GraphOrmError && error.code === 'RESOURCE_NOT_FOUND') {
        return false;
      }
      throw error;
    }
  }

  /**
   * 获取组照片（元数据）
   * 
   * 使用 query-builder 支持调试和自定义 header
   */
  async getPhotoMetadata(groupId: string): Promise<{ width: number; height: number; id: string } | null> {
    return await this.query<{ width: number; height: number; id: string }>(`${encodeURIComponent(groupId)}/photo`).first();
  }

  /**
   * 获取组照片内容
   * 
   * 使用 query-builder 支持调试和自定义 header
   */
  async getPhotoContent(groupId: string): Promise<Blob | null> {
    return await this.query<Blob>(`${encodeURIComponent(groupId)}/photo/$value`).first();
  }
}

