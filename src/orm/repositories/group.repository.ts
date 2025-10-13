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
   */
  async addMember(groupId: string, userId: string): Promise<void> {
    try {
      await this.client
        .api(`/groups/${encodeURIComponent(groupId)}/members/$ref`)
        .post({
          '@odata.id': `https://graph.microsoft.com/v1.0/directoryObjects/${userId}`
        });
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 从组中移除成员
   * 
   * 权限要求：GroupMember.ReadWrite.All, Group.ReadWrite.All, Directory.ReadWrite.All
   */
  async removeMember(groupId: string, userId: string): Promise<void> {
    try {
      await this.client
        .api(`/groups/${encodeURIComponent(groupId)}/members/${encodeURIComponent(userId)}/$ref`)
        .delete();
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 添加所有者到组
   * 
   * 权限要求：Group.ReadWrite.All, Directory.ReadWrite.All
   */
  async addOwner(groupId: string, userId: string): Promise<void> {
    try {
      await this.client
        .api(`/groups/${encodeURIComponent(groupId)}/owners/$ref`)
        .post({
          '@odata.id': `https://graph.microsoft.com/v1.0/directoryObjects/${userId}`
        });
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 从组中移除所有者
   * 
   * 权限要求：Group.ReadWrite.All, Directory.ReadWrite.All
   */
  async removeOwner(groupId: string, userId: string): Promise<void> {
    try {
      await this.client
        .api(`/groups/${encodeURIComponent(groupId)}/owners/${encodeURIComponent(userId)}/$ref`)
        .delete();
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 获取组关联的团队
   * 
   * 权限要求：Team.ReadBasic.All, TeamSettings.Read.All
   */
  async getTeam(groupId: string): Promise<Team | null> {
    try {
      return await this.client
        .api(`/groups/${encodeURIComponent(groupId)}/team`)
        .get();
    } catch (error) {
      if (error instanceof GraphOrmError && error.code === 'RESOURCE_NOT_FOUND') {
        return null;
      }
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 获取组驱动器
   */
  async getDrive(groupId: string): Promise<Drive> {
    try {
      return await this.client
        .api(`/groups/${encodeURIComponent(groupId)}/drive`)
        .get();
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
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
   */
  async isMember(groupId: string, userId: string): Promise<boolean> {
    try {
      await this.client
        .api(`/groups/${encodeURIComponent(groupId)}/members/${encodeURIComponent(userId)}`)
        .get();
      return true;
    } catch (error) {
      if (error instanceof GraphOrmError && error.code === 'RESOURCE_NOT_FOUND') {
        return false;
      }
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 获取组照片（元数据）
   */
  async getPhotoMetadata(groupId: string): Promise<{ width: number; height: number; id: string }> {
    try {
      return await this.client
        .api(`/groups/${encodeURIComponent(groupId)}/photo`)
        .get();
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 获取组照片内容
   */
  async getPhotoContent(groupId: string): Promise<Blob> {
    try {
      return await this.client
        .api(`/groups/${encodeURIComponent(groupId)}/photo/$value`)
        .get();
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }
}

