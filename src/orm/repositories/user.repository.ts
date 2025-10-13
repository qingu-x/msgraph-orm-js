import { Client } from '@microsoft/microsoft-graph-client';
import { GraphRepository } from '../repository';
import { User, Group, Message, Event, Drive, LicenseDetails, MailboxSettings } from '../types';
import { GraphOrmError } from '../errors';

/**
 * 用户仓储
 * 
 * 权限要求：User.Read.All, User.ReadWrite.All, Directory.Read.All
 * 国家云支持：✓ 全球版 ✓ 中国版 ✓ 美国政府版
 * 
 * @see https://learn.microsoft.com/graph/api/resources/user
 */
export class UserRepository extends GraphRepository<User> {
  constructor(client: Client) {
    super(client, '/users');
  }

  /**
   * 通过邮箱查找用户
   */
  async findByEmail(email: string): Promise<User | null> {
    return this.query()
      .where('mail', 'eq', email)
      .first()
      .catch(() => null);
  }

  /**
   * 通过 UPN 查找用户
   */
  async findByUserPrincipalName(upn: string): Promise<User | null> {
    return this.query()
      .where('userPrincipalName', 'eq', upn)
      .first()
      .catch(() => null);
  }

  /**
   * 搜索用户（按显示名称或邮箱）
   */
  async search(keyword: string) {
    return this.query()
      .search('displayName', keyword)
      .get();
  }

  /**
   * 获取用户所属的组
   */
  groups(userId: string) {
    return this.subRepository<Group>(userId, 'memberOf');
  }

  /**
   * 获取用户的直接下属
   */
  directReports(userId: string) {
    return this.subRepository<User>(userId, 'directReports');
  }

  /**
   * 获取用户的邮件
   */
  messages(userId: string) {
    return this.subRepository<Message>(userId, 'messages');
  }

  /**
   * 获取用户的事件
   */
  events(userId: string) {
    return this.subRepository<Event>(userId, 'events');
  }

  /**
   * 获取用户的日历事件
   */
  calendarEvents(userId: string) {
    return this.subRepository<Event>(userId, 'calendar/events');
  }

  /**
   * 获取用户的驱动器
   */
  async getDrive(userId: string): Promise<Drive> {
    try {
      return await this.client
        .api(`/users/${encodeURIComponent(userId)}/drive`)
        .get();
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 获取用户许可证详情
   * 
   * 权限要求：User.Read.All, Directory.Read.All
   */
  async getLicenses(userId: string): Promise<LicenseDetails[]> {
    try {
      const response = await this.client
        .api(`/users/${encodeURIComponent(userId)}/licenseDetails`)
        .get();
      return response.value || [];
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 分配许可证给用户
   * 
   * 权限要求：User.ReadWrite.All, Directory.ReadWrite.All
   */
  async assignLicense(
    userId: string,
    addLicenses: Array<{ skuId: string; disabledPlans?: string[] }>,
    removeLicenses: string[] = []
  ): Promise<User> {
    try {
      return await this.client
        .api(`/users/${encodeURIComponent(userId)}/assignLicense`)
        .post({
          addLicenses,
          removeLicenses
        });
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 获取用户邮箱设置
   * 
   * 权限要求：MailboxSettings.Read, MailboxSettings.ReadWrite
   */
  async getMailboxSettings(userId: string): Promise<MailboxSettings> {
    try {
      return await this.client
        .api(`/users/${encodeURIComponent(userId)}/mailboxSettings`)
        .get();
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 更新用户邮箱设置
   * 
   * 权限要求：MailboxSettings.ReadWrite
   */
  async updateMailboxSettings(userId: string, settings: Partial<MailboxSettings>): Promise<MailboxSettings> {
    try {
      return await this.client
        .api(`/users/${encodeURIComponent(userId)}/mailboxSettings`)
        .patch(settings);
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 获取用户的经理
   */
  async getManager(userId: string): Promise<User | null> {
    try {
      return await this.client
        .api(`/users/${encodeURIComponent(userId)}/manager`)
        .get();
    } catch (error) {
      if (error instanceof GraphOrmError && error.code === 'RESOURCE_NOT_FOUND') {
        return null;
      }
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 获取用户照片（元数据）
   */
  async getPhotoMetadata(userId: string): Promise<{ width: number; height: number; id: string }> {
    try {
      return await this.client
        .api(`/users/${encodeURIComponent(userId)}/photo`)
        .get();
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 获取用户照片内容
   */
  async getPhotoContent(userId: string): Promise<Blob> {
    try {
      return await this.client
        .api(`/users/${encodeURIComponent(userId)}/photo/$value`)
        .get();
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }
}

