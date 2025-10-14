import { Client } from '@microsoft/microsoft-graph-client';
import { UserRepository } from './repositories/user.repository';
import { GroupRepository } from './repositories/group.repository';
import { FileService } from './services/file.service';
import { MailService } from './services/mail.service';
import { CalendarService } from './services/calendar.service';
import { GenericRepository } from './repository';
import { GraphQueryBuilder } from './query-builder';
import {
  Device,
  Application,
  ServicePrincipal,
  Team,
  Chat,
  Plan,
  Subscription,
  DirectoryRole,
  DirectoryRoleTemplate,
  AdministrativeUnit,
  Site,
  DeltaCollection,
  SearchRequest,
  SearchResponse,
  User,
  Group,
  SubscribedSku,
  Invitation,
  DirectoryObject,
  Message,
  Event,
  DriveItem,
  GraphCollection
} from './types';
import { GraphOrmError, GraphErrorCode } from './errors';

/**
 * Microsoft Graph ORM
 * 
 * 轻量级入口类，提供对各种资源仓储和服务的访问
 * 
 * 国家云部署支持说明：
 * - 全球版（Global）：所有功能完全支持
 * - 中国版（21Vianet）：部分功能不可用或受限
 * - 美国政府版（GCC, GCC High, DoD）：大部分功能支持，部分高级功能受限
 * 
 * 注意：德国版（Microsoft Cloud Deutschland）已于 2021 年弃用并迁移至全球版
 * 
 * @see https://learn.microsoft.com/graph/deployments
 * @see https://learn.microsoft.com/graph/permissions-reference
 */
export class GraphORM {
  // ==================== 仓储 (Repositories) ====================
  
  /**
   * 用户仓储
   * 
   * 权限要求：User.Read.All, User.ReadWrite.All, Directory.Read.All
   * 国家云支持：✓ 全球版 ✓ 中国版 ✓ 美国政府版
   * 
   * @see https://learn.microsoft.com/graph/api/resources/user
   */
  public readonly users: UserRepository;

  /**
   * 组仓储
   * 
   * 权限要求：Group.Read.All, Group.ReadWrite.All, Directory.Read.All
   * 国家云支持：✓ 全球版 ✓ 中国版 ✓ 美国政府版
   * 
   * @see https://learn.microsoft.com/graph/api/resources/group
   */
  public readonly groups: GroupRepository;

  /**
   * 设备仓储
   * 
   * 权限要求：Device.Read.All, Device.ReadWrite.All, Directory.Read.All
   * 国家云支持：✓ 全球版 ✓ 中国版 ✓ 美国政府版
   * 
   * @see https://learn.microsoft.com/graph/api/resources/device
   */
  public readonly devices: GenericRepository<Device>;

  /**
   * 应用程序仓储
   * 
   * 权限要求：Application.Read.All, Application.ReadWrite.All, Directory.Read.All
   * 国家云支持：✓ 全球版 ✓ 中国版 ✓ 美国政府版
   * 
   * @see https://learn.microsoft.com/graph/api/resources/application
   */
  public readonly applications: GenericRepository<Application>;

  /**
   * 服务主体仓储
   * 
   * 权限要求：Application.Read.All, Application.ReadWrite.All, Directory.Read.All
   * 国家云支持：✓ 全球版 ✓ 中国版 ✓ 美国政府版
   * 
   * @see https://learn.microsoft.com/graph/api/resources/serviceprincipal
   */
  public readonly servicePrincipals: GenericRepository<ServicePrincipal>;

  /**
   * Teams 团队仓储
   * 
   * 权限要求：Team.ReadBasic.All, TeamSettings.Read.All, TeamSettings.ReadWrite.All
   * 国家云支持：✓ 全球版 ⚠️ 中国版（部分功能受限）✓ 美国政府版（GCC）
   * 
   * @see https://learn.microsoft.com/graph/api/resources/team
   */
  public readonly teams: GenericRepository<Team>;

  /**
   * SharePoint 站点仓储
   * 
   * 权限要求：Sites.Read.All, Sites.ReadWrite.All
   * 国家云支持：✓ 全球版 ✓ 中国版 ✓ 美国政府版
   * 
   * @see https://learn.microsoft.com/graph/api/resources/site
   */
  public readonly sites: GenericRepository<Site>;

  /**
   * Planner 计划仓储
   * 
   * 权限要求：Tasks.Read, Tasks.ReadWrite
   * 国家云支持：✓ 全球版 ⚠️ 中国版（不可用）✓ 美国政府版（GCC）
   * 
   * @see https://learn.microsoft.com/graph/api/resources/planner-overview
   */
  public readonly plans: GenericRepository<Plan>;

  /**
   * 订阅仓储（Webhooks）
   * 
   * 权限要求：根据订阅的资源类型而定
   * 国家云支持：✓ 全球版 ✓ 中国版 ✓ 美国政府版
   * 
   * @see https://learn.microsoft.com/graph/api/resources/subscription
   */
  public readonly subscriptions: GenericRepository<Subscription>;

  /**
   * 目录角色仓储
   * 
   * 权限要求：RoleManagement.Read.Directory, RoleManagement.ReadWrite.Directory
   * 国家云支持：✓ 全球版 ✓ 中国版 ✓ 美国政府版
   * 
   * @see https://learn.microsoft.com/graph/api/resources/directoryrole
   */
  public readonly directoryRoles: GenericRepository<DirectoryRole>;

  /**
   * 目录角色模板仓储
   * 
   * 权限要求：RoleManagement.Read.Directory
   * 国家云支持：✓ 全球版 ✓ 中国版 ✓ 美国政府版
   * 
   * @see https://learn.microsoft.com/graph/api/resources/directoryroletemplate
   */
  public readonly directoryRoleTemplates: GenericRepository<DirectoryRoleTemplate>;

  /**
   * 管理单元仓储
   * 
   * 权限要求：AdministrativeUnit.Read.All, AdministrativeUnit.ReadWrite.All
   * 国家云支持：✓ 全球版 ✓ 中国版 ✓ 美国政府版
   * 
   * @see https://learn.microsoft.com/graph/api/resources/administrativeunit
   */
  public readonly administrativeUnits: GenericRepository<AdministrativeUnit>;

  /**
   * 聊天仓储
   * 
   * 权限要求：Chat.Read, Chat.ReadWrite
   * 国家云支持：✓ 全球版 ⚠️ 中国版（部分功能受限）✓ 美国政府版（GCC）
   * 
   * @see https://learn.microsoft.com/graph/api/resources/chat
   */
  public readonly chats: GenericRepository<Chat>;

  // ==================== 服务 (Services) ====================

  /**
   * 文件服务
   * 
   * 提供 OneDrive 和 SharePoint 文件操作
   */
  public readonly files: FileService;

  /**
   * 邮件服务
   * 
   * 提供邮件发送、接收和管理功能
   */
  public readonly mail: MailService;

  /**
   * 日历服务
   * 
   * 提供日历和事件管理功能
   */
  public readonly calendar: CalendarService;

  constructor(private client: Client, defaultTimeZone?: string) {
    // 初始化仓储
    this.users = new UserRepository(client, defaultTimeZone);
    this.groups = new GroupRepository(client);
    this.devices = new GenericRepository<Device>(client, '/devices');
    this.applications = new GenericRepository<Application>(client, '/applications');
    this.servicePrincipals = new GenericRepository<ServicePrincipal>(client, '/servicePrincipals');
    this.teams = new GenericRepository<Team>(client, '/teams');
    this.sites = new GenericRepository<Site>(client, '/sites');
    this.plans = new GenericRepository<Plan>(client, '/planner/plans');
    this.subscriptions = new GenericRepository<Subscription>(client, '/subscriptions');
    this.directoryRoles = new GenericRepository<DirectoryRole>(client, '/directoryRoles');
    this.directoryRoleTemplates = new GenericRepository<DirectoryRoleTemplate>(client, '/directoryRoleTemplates');
    this.administrativeUnits = new GenericRepository<AdministrativeUnit>(client, '/directory/administrativeUnits');
    this.chats = new GenericRepository<Chat>(client, '/chats');

    // 初始化服务
    this.files = new FileService(client);
    this.mail = new MailService(client);
    this.calendar = new CalendarService(client);
  }

  /**
   * 创建 query-builder 实例（私有辅助方法）
   */
  private query<T>(endpoint: string) {
    return new GraphQueryBuilder<T>(this.client, endpoint);
  }

  // ==================== 便捷访问方法 ====================

  /**
   * 设置默认时区（用于事件创建和更新）
   * 
   * @param timeZone IANA 时区名称，如 'Asia/Shanghai', 'America/New_York', 'UTC', 'China Standard Time'
   * 
   * @example
   * ```typescript
   * orm.setDefaultTimeZone('Asia/Shanghai');
   * // 之后创建的事件会自动使用此时区
   * await orm.users.events('user-id').create({
   *   subject: '会议',
   *   start: { dateTime: '2024-01-15T10:00:00' },  // 自动应用 Asia/Shanghai
   *   end: { dateTime: '2024-01-15T11:00:00' }     // 自动应用 Asia/Shanghai
   * });
   * ```
   */
  setDefaultTimeZone(timeZone: string): void {
    this.users.setDefaultTimeZone(timeZone);
  }

  /**
   * 获取当前默认时区
   */
  getDefaultTimeZone(): string {
    return this.users.getDefaultTimeZone();
  }

  /**
   * 获取当前认证用户信息
   * 
   * 注意：此方法需要委托权限，不适用于纯应用权限场景
   * 使用 query-builder 支持调试和自定义 header
   */
  async me(): Promise<User | null> {
    return await this.query<User>('me').first();
  }

  /**
   * 获取组织信息
   * 
   * 权限要求：Organization.Read.All
   * 使用 query-builder 支持调试和自定义 header
   */
  async organization(): Promise<unknown | null> {
    return await this.query<unknown>('organization').first();
  }

  /**
   * 获取订阅的 SKU
   * 
   * 权限要求：Organization.Read.All, Directory.Read.All
   * 使用 query-builder 支持调试和自定义 header
   */
  async getSubscribedSkus(): Promise<GraphCollection<SubscribedSku>> {
    return await this.query<SubscribedSku>('subscribedSkus').get();
  }

  // ==================== Delta Query（增量查询）====================

  /**
   * 用户增量查询
   * 
   * 权限要求：User.Read.All
   * 使用 query-builder 支持调试和自定义 header
   */
  async deltaUsers(deltaLink?: string): Promise<GraphCollection<User>> {
    const endpoint = deltaLink || 'users/delta';
    return await this.query<User>(endpoint).get();
  }

  /**
   * 组增量查询
   * 
   * 使用 query-builder 支持调试和自定义 header
   */
  async deltaGroups(deltaLink?: string): Promise<GraphCollection<Group>> {
    const endpoint = deltaLink || 'groups/delta';
    return await this.query<Group>(endpoint).get();
  }

  /**
   * 邮件增量查询
   */
  async deltaMessages(userId: string, deltaLink?: string): Promise<GraphCollection<Message>> {
    return this.mail.delta(userId, deltaLink);
  }

  /**
   * 事件增量查询
   */
  async deltaEvents(userId: string, deltaLink?: string): Promise<DeltaCollection<Event>> {
    return this.calendar.deltaEvents(userId, deltaLink);
  }

  /**
   * 驱动器项增量查询
   */
  async deltaDriveItems(userId: string, deltaLink?: string): Promise<DeltaCollection<DriveItem>> {
    return this.files.delta(userId, deltaLink);
  }

  // ==================== Search API ====================

  /**
   * Microsoft Search API
   * 
   * 权限要求：根据搜索的资源类型而定
   * 国家云支持：✓ 全球版 ⚠️ 中国版（不可用）✓ 美国政府版（GCC）
   * 使用 query-builder 支持调试和自定义 header
   */
  async search(searchRequest: SearchRequest): Promise<SearchResponse> {
    const response = await this.query('').post<{ value: SearchResponse[] }>(
      'query',
      { requests: [searchRequest] },
      'search'
    );
    return response.value[0];
  }

  /**
   * 快速搜索用户
   */
  async searchUsers(queryString: string, size: number = 25): Promise<SearchResponse> {
    return this.search({
      entityTypes: ['person'],
      query: { queryString },
      size
    });
  }

  /**
   * 快速搜索邮件
   */
  async searchMessages(queryString: string, size: number = 25): Promise<SearchResponse> {
    return this.search({
      entityTypes: ['message'],
      query: { queryString },
      size
    });
  }

  /**
   * 快速搜索文件
   */
  async searchFiles(queryString: string, size: number = 25): Promise<SearchResponse> {
    return this.search({
      entityTypes: ['driveItem'],
      query: { queryString },
      size
    });
  }

  // ==================== 批处理请求 ====================

  /**
   * 执行批处理请求
   * 
   * 允许在单个 HTTP 请求中发送多个 API 调用（最多 20 个）
   * 
   * @example
   * ```typescript
   * const batchResponse = await orm.batch([
   *   { id: '1', method: 'GET', url: '/users/user1@contoso.com' },
   *   { id: '2', method: 'GET', url: '/users/user2@contoso.com' }
   * ]);
   * ```
   * 使用 query-builder 支持调试和自定义 header
   */
  async batch(requests: Array<{
    id: string;
    method: string;
    url: string;
    body?: unknown;
    headers?: Record<string, string>;
  }>): Promise<{
    responses: Array<{
      id: string;
      status: number;
      headers?: Record<string, string>;
      body: unknown;
    }>;
  }> {
    if (requests.length > 20) {
      throw GraphOrmError.create(
        GraphErrorCode.INVALID_PARAMETER,
        '批处理请求最多支持 20 个请求'
      );
    }

    return this.query('').post<{
      responses: Array<{
        id: string;
        status: number;
        headers?: Record<string, string>;
        body: unknown;
      }>;
    }>('$batch', { requests });
  }

  // ==================== 邀请管理 ====================

  /**
   * 创建邀请
   * 
   * 邀请外部用户加入组织
   * 
   * 权限要求：User.Invite.All
   * 使用 query-builder 支持调试和自定义 header
   */
  async createInvitation(invitation: Omit<Invitation, 'id'>): Promise<Invitation> {
    return this.query('').post<Invitation>('invitations', invitation);
  }

  // ==================== 目录对象 ====================

  /**
   * 通过 ID 获取目录对象
   * 
   * 权限要求：Directory.Read.All
   * 使用 query-builder 支持调试和自定义 header
   */
  async getDirectoryObject(id: string): Promise<DirectoryObject | null> {
    return this.query<DirectoryObject>(`directoryObjects/${encodeURIComponent(id)}`).first();
  }

  /**
   * 通过 ID 列表获取目录对象（最多 1000 个）
   * 
   * 使用 query-builder 支持调试和自定义 header
   */
  async getDirectoryObjectsByIds(ids: string[], types?: string[]): Promise<DirectoryObject[]> {
    const body: { ids: string[]; types?: string[] } = { ids };
    if (types && types.length > 0) {
      body.types = types;
    }
    const response = await this.query('').post<{ value: DirectoryObject[] }>(
      'getByIds',
      body,
      'directoryObjects'
    );
    return response.value || [];
  }

  /**
   * 检查成员组
   * 
   * 使用 query-builder 支持调试和自定义 header
   */
  async checkMemberGroups(memberId: string, groupIds: string[]): Promise<string[]> {
    const response = await this.query('').post<{ value: string[] }>(
      'checkMemberGroups',
      { groupIds },
      `directoryObjects/${encodeURIComponent(memberId)}`
    );
    return response.value || [];
  }

  /**
   * 获取成员对象
   * 
   * 使用 query-builder 支持调试和自定义 header
   */
  async getMemberObjects(directoryObjectId: string, securityEnabledOnly: boolean = false): Promise<string[]> {
    const response = await this.query('').post<{ value: string[] }>(
      'getMemberObjects',
      { securityEnabledOnly },
      `directoryObjects/${encodeURIComponent(directoryObjectId)}`
    );
    return response.value || [];
  }

  // ==================== 报告 API ====================

  /**
   * 获取 Office 365 活动用户详情
   * 
   * 权限要求：Reports.Read.All
   * 使用 query-builder 支持调试和自定义 header
   */
  async getOffice365ActiveUserDetail(period: 'D7' | 'D30' | 'D90' | 'D180'): Promise<string | null> {
    return this.query<string>(`reports/getOffice365ActiveUserDetail(period='${period}')`).first();
  }
}

/**
 * 便捷的 ORM 工厂函数
 * 
 * @param client Microsoft Graph 客户端实例
 * @param defaultTimeZone 默认时区，用于事件创建和更新（如 'Asia/Shanghai', 'UTC'）
 */
export function createGraphORM(client: Client, defaultTimeZone?: string): GraphORM {
  return new GraphORM(client, defaultTimeZone);
}

