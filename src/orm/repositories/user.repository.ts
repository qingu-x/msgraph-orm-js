import { Client } from '@microsoft/microsoft-graph-client';
import { GraphRepository } from '../repository';
import { User, Group, Drive, LicenseDetails, MailboxSettings } from '../types';
import { GraphOrmError } from '../errors';
import { MessageRepository } from './message.repository';
import { EventRepository } from './event.repository';
import { DriveItemRepository } from './drive-item.repository';

/**
 * 用户仓储
 * 
 * 权限要求：User.Read.All, User.ReadWrite.All, Directory.Read.All
 * 国家云支持：✓ 全球版 ✓ 中国版 ✓ 美国政府版
 * 
 * @see https://learn.microsoft.com/graph/api/resources/user
 */
export class UserRepository extends GraphRepository<User> {
  private defaultTimeZone: string = 'UTC';

  constructor(client: Client, defaultTimeZone?: string) {
    super(client, '/users');
    if (defaultTimeZone) {
      this.defaultTimeZone = defaultTimeZone;
    }
  }

  /**
   * 设置默认时区
   * 
   * @param timeZone IANA 时区名称，如 'Asia/Shanghai', 'America/New_York', 'UTC'
   */
  setDefaultTimeZone(timeZone: string): void {
    this.defaultTimeZone = timeZone;
  }

  /**
   * 获取当前默认时区
   */
  getDefaultTimeZone(): string {
    return this.defaultTimeZone;
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
   * 获取用户的邮件仓储
   * 
   * 返回 MessageRepository，支持查询构建器和邮件特定操作
   */
  messages(userId: string): MessageRepository {
    const endpoint = `/users/${encodeURIComponent(userId)}/messages`;
    return new MessageRepository(this.client, endpoint);
  }

  /**
   * 获取用户的事件仓储（默认日历）
   * 
   * 返回 EventRepository，支持查询构建器和事件特定操作
   * 
   * 注意：这是访问用户默认日历的快捷方式
   * 如果需要访问特定日历，请使用 CalendarService
   */
  events(userId: string): EventRepository {
    const endpoint = `/users/${encodeURIComponent(userId)}/events`;
    return new EventRepository(this.client, endpoint, this.defaultTimeZone);
  }

  /**
   * 获取用户的日历事件仓储
   */
  calendarEvents(userId: string): EventRepository {
    const endpoint = `/users/${encodeURIComponent(userId)}/calendar/events`;
    return new EventRepository(this.client, endpoint, this.defaultTimeZone);
  }

  /**
   * 获取用户的驱动器项仓储
   * 
   * 返回 DriveItemRepository，支持查询构建器和文件特定操作
   */
  driveItems(userId: string): DriveItemRepository {
    const endpoint = `/users/${encodeURIComponent(userId)}/drive/items`;
    return new DriveItemRepository(this.client, endpoint);
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

