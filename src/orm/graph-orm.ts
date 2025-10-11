import { Client } from '@microsoft/microsoft-graph-client';
import { GraphEntityManager } from './entity-manager';
import { 
  User, 
  Device, 
  Event, 
  Room, 
  Group, 
  Contact, 
  Message, 
  DriveItem, 
  Site, 
  Application, 
  ServicePrincipal,
  Team,
  Channel,
  Chat,
  Plan,
  Task,
  Subscription,
  DeltaCollection,
  SearchRequest,
  SearchResponse,
  DirectoryRole,
  DirectoryRoleTemplate,
  AdministrativeUnit,
  LicenseDetails,
  SubscribedSku,
  Invitation,
  Calendar,
  CalendarGroup,
  MailFolder,
  MessageRule,
  MailboxSettings,
  InferenceClassification,
  Attachment,
  FileAttachment,
  ContactFolder,
  OrgContact,
  Drive,
  Notebook,
  Section,
  OneNotePage,
  Bucket,
  TodoTaskList,
  TodoTask,
  List,
  ListItem,
  ColumnDefinition,
  ChatMessage,
  ConversationMember,
  TeamsAppInstallation,
  TeamworkTag,
  DirectoryAudit,
  SignIn,
  Trending,
  UsedInsight,
  SharedInsight,
  Person,
  Alert,
  SecureScore,
  RiskDetection,
  RiskyUser,
  SchemaExtension,
  Extension,
  DirectoryObject
} from './types';
import { GraphOrmError, GraphErrorCode } from './errors';

/**
 * Microsoft Graph ORM
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
  constructor(private client: Client) {}

  // ==================== 用户和身份 ====================
  
  /**
   * 用户管理
   * 
   * 权限要求：User.Read.All, User.ReadWrite.All, Directory.Read.All
   * 国家云支持：✓ 全球版 ✓ 中国版 ✓ 美国政府版
   * 
   * @see https://learn.microsoft.com/graph/api/resources/user
   */
  get users() {
    return new GraphEntityManager<User>(this.client, '/users');
  }

  /**
   * 组管理
   * 
   * 权限要求：Group.Read.All, Group.ReadWrite.All, Directory.Read.All
   * 国家云支持：✓ 全球版 ✓ 中国版 ✓ 美国政府版
   * 
   * @see https://learn.microsoft.com/graph/api/resources/group
   */
  get groups() {
    return new GraphEntityManager<Group>(this.client, '/groups');
  }

  /**
   * 获取组成员
   * 
   * @param groupId - 组 ID
   */
  groupMembers(groupId: string) {
    return new GraphEntityManager<User>(this.client, `/groups/${encodeURIComponent(groupId)}/members`);
  }

  /**
   * 获取组所有者
   * 
   * @param groupId - 组 ID
   */
  groupOwners(groupId: string) {
    return new GraphEntityManager<User>(this.client, `/groups/${encodeURIComponent(groupId)}/owners`);
  }

  /**
   * 设备管理
   * 
   * 权限要求：Device.Read.All, Device.ReadWrite.All, Directory.Read.All
   * 国家云支持：✓ 全球版 ✓ 中国版 ✓ 美国政府版
   * 
   * @see https://learn.microsoft.com/graph/api/resources/device
   */
  get devices() {
    return new GraphEntityManager<Device>(this.client, '/devices');
  }

  /**
   * 应用程序管理
   * 
   * 权限要求：Application.Read.All, Application.ReadWrite.All, Directory.Read.All
   * 国家云支持：✓ 全球版 ✓ 中国版 ✓ 美国政府版
   * 
   * @see https://learn.microsoft.com/graph/api/resources/application
   */
  get applications() {
    return new GraphEntityManager<Application>(this.client, '/applications');
  }

  /**
   * 服务主体管理
   * 
   * 权限要求：Application.Read.All, Application.ReadWrite.All, Directory.Read.All
   * 国家云支持：✓ 全球版 ✓ 中国版 ✓ 美国政府版
   * 
   * @see https://learn.microsoft.com/graph/api/resources/serviceprincipal
   */
  get servicePrincipals() {
    return new GraphEntityManager<ServicePrincipal>(this.client, '/servicePrincipals');
  }

  // ==================== 日历和事件 ====================

  /**
   * 用户事件管理（需要指定用户）
   * 
   * 权限要求：Calendars.Read, Calendars.ReadWrite
   * 国家云支持：✓ 全球版 ✓ 中国版 ✓ 美国政府版
   * 
   * @param userId - 用户 ID 或 UPN
   * @see https://learn.microsoft.com/graph/api/resources/event
   */
  events(userId: string) {
    return new GraphEntityManager<Event>(this.client, `/users/${encodeURIComponent(userId)}/events`);
  }

  /**
   * 用户日历事件管理
   * 
   * 与 events() 的区别：此方法访问主日历的事件
   * 
   * @param userId - 用户 ID 或 UPN
   */
  calendarEvents(userId: string) {
    return new GraphEntityManager<Event>(this.client, `/users/${encodeURIComponent(userId)}/calendar/events`);
  }

  /**
   * 指定日历的事件
   * 
   * @param userId - 用户 ID 或 UPN
   * @param calendarId - 日历 ID
   */
  calendarEventsById(userId: string, calendarId: string) {
    return new GraphEntityManager<Event>(
      this.client, 
      `/users/${encodeURIComponent(userId)}/calendars/${encodeURIComponent(calendarId)}/events`
    );
  }

  /**
   * 组日历事件
   * 
   * @param groupId - 组 ID
   */
  groupEvents(groupId: string) {
    return new GraphEntityManager<Event>(this.client, `/groups/${encodeURIComponent(groupId)}/events`);
  }

  // ==================== 会议室 ====================
  
  /**
   * 会议室管理 - 使用 places API
   * 
   * 权限要求：Place.Read.All
   * 国家云支持：✓ 全球版 ✓ 美国政府版（GCC）⚠️ 中国版（功能受限）⚠️ 美国政府版（GCC High, DoD 不支持）
   * 
   * 注意：
   * - 中国版（21Vianet）：places API 可能不可用或功能受限
   * - 建议在中国版使用 findRooms() 替代方案
   * 
   * @see https://learn.microsoft.com/graph/api/resources/place
   */
  get rooms() {
    return new GraphEntityManager<Room>(this.client, '/places/microsoft.graph.room');
  }

  /**
   * 会议室列表管理 - 使用 places API
   * 
   * 权限要求：Place.Read.All
   * 国家云支持：✓ 全球版 ✓ 美国政府版（GCC）⚠️ 中国版（功能受限）⚠️ 美国政府版（GCC High, DoD 不支持）
   */
  get roomLists() {
    return new GraphEntityManager<Room>(this.client, '/places/microsoft.graph.roomList');
  }

  /**
   * 获取指定会议室列表下的会议室
   * 
   * @param roomListId - 会议室列表 ID
   */
  roomsInList(roomListId: string) {
    return new GraphEntityManager<Room>(
      this.client, 
      `/places/${encodeURIComponent(roomListId)}/microsoft.graph.roomList/rooms`
    );
  }

  /**
   * 查找用户可访问的会议室 - 使用 findRooms 函数
   * 
   * 权限要求：Calendars.Read（比 Place.Read.All 权限要求更低）
   * 国家云支持：✓ 全球版 ⚠️ 美国政府版（部分支持）❌ 中国版（不可用）
   * 
   * 注意：
   * - 此 API 是函数调用，使用 POST 方法
   * - 中国版（21Vianet）：此 API 不可用，返回空数组
   * - 如果用户没有配置会议室，返回空数组而不抛出错误
   * 
   * @param userId - 用户 ID 或 UPN
   * @returns 会议室列表，如果不可用则返回空数组
   * @see https://learn.microsoft.com/graph/api/user-findrooms
   */
  async findRooms(userId: string): Promise<Room[]> {
    try {
      const response = await this.client
        .api(`/users/${encodeURIComponent(userId)}/findRooms`)
        .post(null);
      return response.value || response || [];
    } catch (error) {
      const ormError = GraphOrmError.fromGraphError(error);
      // 404 或 400 可能意味着该用户没有配置会议室或 API 不可用（如中国版）
      if (ormError.code === GraphErrorCode.RESOURCE_NOT_FOUND || 
          ormError.code === GraphErrorCode.BAD_REQUEST) {
        return [];
      }
      throw ormError;
    }
  }

  /**
   * 查找会议室列表 - 使用 findRoomLists 函数
   * 
   * 权限要求：Calendars.Read
   * 国家云支持：✓ 全球版 ⚠️ 美国政府版（部分支持）❌ 中国版（不可用）
   * 
   * @param userId - 用户 ID 或 UPN
   * @returns 会议室列表，如果不可用则返回空数组
   * @see https://learn.microsoft.com/graph/api/user-findroomlists
   */
  async findRoomLists(userId: string): Promise<Room[]> {
    try {
      const response = await this.client
        .api(`/users/${encodeURIComponent(userId)}/findRoomLists`)
        .post(null);
      return response.value || response || [];
    } catch (error) {
      const ormError = GraphOrmError.fromGraphError(error);
      // 404 或 400 可能意味着该用户没有配置会议室或 API 不可用（如中国版）
      if (ormError.code === GraphErrorCode.RESOURCE_NOT_FOUND || 
          ormError.code === GraphErrorCode.BAD_REQUEST) {
        return [];
      }
      throw ormError;
    }
  }

  /**
   * 查找指定会议室列表下的会议室
   * 
   * 权限要求：Calendars.Read
   * 国家云支持：✓ 全球版 ⚠️ 美国政府版（部分支持）❌ 中国版（不可用）
   * 
   * @param userId - 用户 ID 或 UPN
   * @param roomListEmail - 会议室列表的电子邮件地址
   * @returns 会议室列表，如果不可用则返回空数组
   */
  async findRoomsInList(userId: string, roomListEmail: string): Promise<Room[]> {
    try {
      const response = await this.client
        .api(`/users/${encodeURIComponent(userId)}/findRooms(RoomList='${roomListEmail}')`)
        .post(null);
      return response.value || response || [];
    } catch (error) {
      const ormError = GraphOrmError.fromGraphError(error);
      // 404 或 400 可能意味着该用户没有配置会议室或 API 不可用（如中国版）
      if (ormError.code === GraphErrorCode.RESOURCE_NOT_FOUND || 
          ormError.code === GraphErrorCode.BAD_REQUEST) {
        return [];
      }
      throw ormError;
    }
  }

  // ==================== 邮件和联系人 ====================

  /**
   * 用户邮件管理
   * 
   * 权限要求：Mail.Read, Mail.ReadWrite
   * 国家云支持：✓ 全球版 ✓ 中国版 ✓ 美国政府版
   * 
   * @param userId - 用户 ID 或 UPN
   * @see https://learn.microsoft.com/graph/api/resources/message
   */
  messages(userId: string) {
    return new GraphEntityManager<Message>(this.client, `/users/${encodeURIComponent(userId)}/messages`);
  }

  /**
   * 用户收件箱邮件
   * 
   * @param userId - 用户 ID 或 UPN
   */
  inbox(userId: string) {
    return new GraphEntityManager<Message>(
      this.client, 
      `/users/${encodeURIComponent(userId)}/mailFolders/inbox/messages`
    );
  }

  /**
   * 用户已发送邮件
   * 
   * @param userId - 用户 ID 或 UPN
   */
  sentItems(userId: string) {
    return new GraphEntityManager<Message>(
      this.client, 
      `/users/${encodeURIComponent(userId)}/mailFolders/sentitems/messages`
    );
  }

  /**
   * 用户草稿箱邮件
   * 
   * @param userId - 用户 ID 或 UPN
   */
  drafts(userId: string) {
    return new GraphEntityManager<Message>(
      this.client, 
      `/users/${encodeURIComponent(userId)}/mailFolders/drafts/messages`
    );
  }

  /**
   * 用户联系人管理
   * 
   * 权限要求：Contacts.Read, Contacts.ReadWrite
   * 国家云支持：✓ 全球版 ✓ 中国版 ✓ 美国政府版
   * 
   * @param userId - 用户 ID 或 UPN
   * @see https://learn.microsoft.com/graph/api/resources/contact
   */
  contacts(userId: string) {
    return new GraphEntityManager<Contact>(this.client, `/users/${encodeURIComponent(userId)}/contacts`);
  }

  /**
   * 指定联系人文件夹的联系人
   * 
   * @param userId - 用户 ID 或 UPN
   * @param folderId - 联系人文件夹 ID
   */
  contactsInFolder(userId: string, folderId: string) {
    return new GraphEntityManager<Contact>(
      this.client, 
      `/users/${encodeURIComponent(userId)}/contactFolders/${encodeURIComponent(folderId)}/contacts`
    );
  }

  // ==================== OneDrive 和 SharePoint ====================

  /**
   * 用户驱动器项管理
   * 
   * 权限要求：Files.Read, Files.ReadWrite, Files.Read.All, Files.ReadWrite.All
   * 国家云支持：✓ 全球版 ✓ 中国版 ✓ 美国政府版
   * 
   * @param userId - 用户 ID 或 UPN
   * @see https://learn.microsoft.com/graph/api/resources/driveitem
   */
  driveItems(userId: string) {
    return new GraphEntityManager<DriveItem>(
      this.client, 
      `/users/${encodeURIComponent(userId)}/drive/root/children`
    );
  }

  /**
   * 用户驱动器指定路径的项
   * 
   * @param userId - 用户 ID 或 UPN
   * @param itemPath - 项目路径（例如：'/Documents'）
   */
  driveItemsByPath(userId: string, itemPath: string) {
    const encodedPath = itemPath.split('/').map(p => encodeURIComponent(p)).join('/');
    return new GraphEntityManager<DriveItem>(
      this.client, 
      `/users/${encodeURIComponent(userId)}/drive/root:${encodedPath}:/children`
    );
  }

  /**
   * 组驱动器项
   * 
   * @param groupId - 组 ID
   */
  groupDriveItems(groupId: string) {
    return new GraphEntityManager<DriveItem>(
      this.client, 
      `/groups/${encodeURIComponent(groupId)}/drive/root/children`
    );
  }

  /**
   * 获取驱动器项（通过 ID）
   * 
   * 权限要求：Files.Read, Files.ReadWrite, Files.Read.All, Files.ReadWrite.All
   * 
   * @param userId - 用户 ID 或 UPN
   * @param itemId - 驱动器项 ID
   * @returns 驱动器项
   * @see https://learn.microsoft.com/graph/api/driveitem-get
   */
  async getDriveItem(userId: string, itemId: string): Promise<DriveItem> {
    try {
      return await this.client
        .api(`/users/${encodeURIComponent(userId)}/drive/items/${encodeURIComponent(itemId)}`)
        .get();
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 获取驱动器项（通过路径）
   * 
   * @param userId - 用户 ID 或 UPN
   * @param itemPath - 项目路径（例如：'/Documents/report.xlsx'）
   * @returns 驱动器项
   */
  async getDriveItemByPath(userId: string, itemPath: string): Promise<DriveItem> {
    try {
      return await this.client
        .api(`/users/${encodeURIComponent(userId)}/drive/root:${itemPath}`)
        .get();
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 列出驱动器项的子项
   * 
   * @param userId - 用户 ID 或 UPN
   * @param itemId - 父项 ID
   * @returns 子项列表
   * @see https://learn.microsoft.com/graph/api/driveitem-list-children
   */
  async listDriveItemChildren(userId: string, itemId: string): Promise<DriveItem[]> {
    try {
      const response = await this.client
        .api(`/users/${encodeURIComponent(userId)}/drive/items/${encodeURIComponent(itemId)}/children`)
        .get();
      return response.value || [];
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 创建文件夹
   * 
   * 权限要求：Files.ReadWrite, Files.ReadWrite.All
   * 
   * @param userId - 用户 ID 或 UPN
   * @param parentId - 父文件夹 ID（如果为 'root' 则在根目录创建）
   * @param folderName - 文件夹名称
   * @param conflictBehavior - 冲突处理：'fail' | 'replace' | 'rename'
   * @returns 创建的文件夹
   * @see https://learn.microsoft.com/graph/api/driveitem-post-children
   */
  async createFolder(
    userId: string, 
    parentId: string, 
    folderName: string,
    conflictBehavior: 'fail' | 'replace' | 'rename' = 'fail'
  ): Promise<DriveItem> {
    try {
      const endpoint = parentId === 'root' 
        ? `/users/${encodeURIComponent(userId)}/drive/root/children`
        : `/users/${encodeURIComponent(userId)}/drive/items/${encodeURIComponent(parentId)}/children`;

      return await this.client
        .api(endpoint)
        .header('Content-Type', 'application/json')
        .post({
          name: folderName,
          folder: {},
          '@microsoft.graph.conflictBehavior': conflictBehavior
        });
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 上传小文件（< 4MB）
   * 
   * 权限要求：Files.ReadWrite, Files.ReadWrite.All
   * 
   * @param userId - 用户 ID 或 UPN
   * @param itemPath - 文件路径（例如：'/Documents/report.pdf'）
   * @param content - 文件内容（Buffer 或 Blob）
   * @param conflictBehavior - 冲突处理：'fail' | 'replace' | 'rename'
   * @returns 上传的文件
   * @see https://learn.microsoft.com/graph/api/driveitem-put-content
   */
  async uploadSmallFile(
    userId: string,
    itemPath: string,
    content: Buffer | Blob,
    conflictBehavior: 'fail' | 'replace' | 'rename' = 'replace'
  ): Promise<DriveItem> {
    try {
      const url = `/users/${encodeURIComponent(userId)}/drive/root:${itemPath}:/content?@microsoft.graph.conflictBehavior=${conflictBehavior}`;
      return await this.client
        .api(url)
        .header('Content-Type', 'application/octet-stream')
        .put(content);
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 下载文件内容
   * 
   * 权限要求：Files.Read, Files.ReadWrite, Files.Read.All, Files.ReadWrite.All
   * 
   * @param userId - 用户 ID 或 UPN
   * @param itemId - 文件 ID
   * @returns 下载 URL（短期有效，约1小时）
   * @see https://learn.microsoft.com/graph/api/driveitem-get-content
   */
  async downloadFile(userId: string, itemId: string): Promise<string> {
    try {
      const item = await this.client
        .api(`/users/${encodeURIComponent(userId)}/drive/items/${encodeURIComponent(itemId)}`)
        .select('@microsoft.graph.downloadUrl')
        .get();
      return item['@microsoft.graph.downloadUrl'];
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 下载特定格式的文件（例如：将 Word 转换为 PDF）
   * 
   * 权限要求：Files.Read, Files.ReadWrite, Files.Read.All, Files.ReadWrite.All
   * 
   * @param userId - 用户 ID 或 UPN
   * @param itemId - 文件 ID
   * @param format - 目标格式（如 'pdf'）
   * @returns 下载 URL
   * @see https://learn.microsoft.com/graph/api/driveitem-get-content-format
   */
  async downloadFileAsFormat(userId: string, itemId: string, format: string): Promise<string> {
    try {
      const response = await this.client
        .api(`/users/${encodeURIComponent(userId)}/drive/items/${encodeURIComponent(itemId)}/content?format=${format}`)
        .get();
      return response;
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 更新驱动器项元数据
   * 
   * 权限要求：Files.ReadWrite, Files.ReadWrite.All
   * 
   * @param userId - 用户 ID 或 UPN
   * @param itemId - 项 ID
   * @param updates - 要更新的属性
   * @returns 更新后的项
   * @see https://learn.microsoft.com/graph/api/driveitem-update
   */
  async updateDriveItem(userId: string, itemId: string, updates: Partial<DriveItem>): Promise<DriveItem> {
    try {
      return await this.client
        .api(`/users/${encodeURIComponent(userId)}/drive/items/${encodeURIComponent(itemId)}`)
        .patch(updates);
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 删除驱动器项（移到回收站）
   * 
   * 权限要求：Files.ReadWrite, Files.ReadWrite.All
   * 
   * @param userId - 用户 ID 或 UPN
   * @param itemId - 项 ID
   * @see https://learn.microsoft.com/graph/api/driveitem-delete
   */
  async deleteDriveItem(userId: string, itemId: string): Promise<void> {
    try {
      await this.client
        .api(`/users/${encodeURIComponent(userId)}/drive/items/${encodeURIComponent(itemId)}`)
        .delete();
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 永久删除驱动器项（跳过回收站）
   * 
   * 权限要求：Files.ReadWrite, Files.ReadWrite.All
   * 注意：此操作不可恢复
   * 
   * @param userId - 用户 ID 或 UPN
   * @param itemId - 项 ID
   * @see https://learn.microsoft.com/graph/api/driveitem-permanentdelete
   */
  async permanentlyDeleteDriveItem(userId: string, itemId: string): Promise<void> {
    try {
      await this.client
        .api(`/users/${encodeURIComponent(userId)}/drive/items/${encodeURIComponent(itemId)}/permanentDelete`)
        .post({});
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 移动驱动器项
   * 
   * 权限要求：Files.ReadWrite, Files.ReadWrite.All
   * 
   * @param userId - 用户 ID 或 UPN
   * @param itemId - 项 ID
   * @param targetParentId - 目标父文件夹 ID
   * @param newName - 新名称（可选）
   * @returns 移动后的项
   * @see https://learn.microsoft.com/graph/api/driveitem-move
   */
  async moveDriveItem(
    userId: string,
    itemId: string,
    targetParentId: string,
    newName?: string
  ): Promise<DriveItem> {
    try {
      const body: { parentReference: { id: string }; name?: string } = {
        parentReference: { id: targetParentId }
      };
      if (newName) {
        body.name = newName;
      }
      return await this.client
        .api(`/users/${encodeURIComponent(userId)}/drive/items/${encodeURIComponent(itemId)}`)
        .patch(body);
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 复制驱动器项
   * 
   * 权限要求：Files.ReadWrite, Files.ReadWrite.All
   * 注意：复制操作是异步的，返回监控 URL
   * 
   * @param userId - 用户 ID 或 UPN
   * @param itemId - 项 ID
   * @param targetParentId - 目标父文件夹 ID
   * @param newName - 新名称（可选）
   * @returns 监控 URL
   * @see https://learn.microsoft.com/graph/api/driveitem-copy
   */
  async copyDriveItem(
    userId: string,
    itemId: string,
    targetParentId: string,
    newName?: string
  ): Promise<string> {
    try {
      const body: { parentReference: { id: string }; name?: string } = {
        parentReference: { id: targetParentId }
      };
      if (newName) {
        body.name = newName;
      }
      const response = await this.client
        .api(`/users/${encodeURIComponent(userId)}/drive/items/${encodeURIComponent(itemId)}/copy`)
        .post(body);
      // 返回 Location 头中的监控 URL
      return response;
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 搜索驱动器项
   * 
   * 权限要求：Files.Read, Files.ReadWrite, Files.Read.All, Files.ReadWrite.All
   * 
   * @param userId - 用户 ID 或 UPN
   * @param query - 搜索查询
   * @returns 匹配的项列表
   * @see https://learn.microsoft.com/graph/api/driveitem-search
   */
  async searchDriveItems(userId: string, query: string): Promise<DriveItem[]> {
    try {
      const response = await this.client
        .api(`/users/${encodeURIComponent(userId)}/drive/root/search(q='${encodeURIComponent(query)}')`)
        .get();
      return response.value || [];
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 驱动器项增量查询（Delta Query）
   * 
   * 跟踪驱动器项的变化
   * 
   * 权限要求：Files.Read, Files.ReadWrite, Files.Read.All, Files.ReadWrite.All
   * 
   * @param userId - 用户 ID 或 UPN
   * @param deltaLink - 上次查询返回的 deltaLink
   * @returns Delta 查询结果
   * @see https://learn.microsoft.com/graph/api/driveitem-delta
   */
  async deltaDriveItems(userId: string, deltaLink?: string): Promise<DeltaCollection<DriveItem>> {
    try {
      const endpoint = deltaLink || `/users/${encodeURIComponent(userId)}/drive/root/delta`;
      const response = await this.client.api(endpoint).get();
      
      return {
        meta: {
          deltaLink: response['@odata.deltaLink'],
          nextLink: response['@odata.nextLink'],
          count: response['@odata.count'],
        },
        data: response.value || []
      };
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 关注驱动器项（添加到"我的文件"）
   * 
   * 权限要求：Files.ReadWrite, Files.ReadWrite.All
   * 
   * @param userId - 用户 ID 或 UPN
   * @param itemId - 项 ID
   * @returns 关注后的项
   * @see https://learn.microsoft.com/graph/api/driveitem-follow
   */
  async followDriveItem(userId: string, itemId: string): Promise<DriveItem> {
    try {
      return await this.client
        .api(`/users/${encodeURIComponent(userId)}/drive/items/${encodeURIComponent(itemId)}/follow`)
        .post({});
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 取消关注驱动器项
   * 
   * 权限要求：Files.ReadWrite, Files.ReadWrite.All
   * 
   * @param userId - 用户 ID 或 UPN
   * @param itemId - 项 ID
   * @see https://learn.microsoft.com/graph/api/driveitem-unfollow
   */
  async unfollowDriveItem(userId: string, itemId: string): Promise<void> {
    try {
      await this.client
        .api(`/users/${encodeURIComponent(userId)}/drive/items/${encodeURIComponent(itemId)}/unfollow`)
        .post({});
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 获取驱动器项缩略图
   * 
   * 权限要求：Files.Read, Files.ReadWrite, Files.Read.All, Files.ReadWrite.All
   * 
   * @param userId - 用户 ID 或 UPN
   * @param itemId - 项 ID
   * @returns 缩略图集合
   * @see https://learn.microsoft.com/graph/api/driveitem-list-thumbnails
   */
  async getDriveItemThumbnails(userId: string, itemId: string): Promise<unknown[]> {
    try {
      const response = await this.client
        .api(`/users/${encodeURIComponent(userId)}/drive/items/${encodeURIComponent(itemId)}/thumbnails`)
        .get();
      return response.value || [];
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 创建共享链接
   * 
   * 权限要求：Files.ReadWrite, Files.ReadWrite.All
   * 
   * @param userId - 用户 ID 或 UPN
   * @param itemId - 项 ID
   * @param type - 链接类型：'view' | 'edit' | 'embed'
   * @param scope - 范围：'anonymous' | 'organization'
   * @returns 共享链接信息
   * @see https://learn.microsoft.com/graph/api/driveitem-createlink
   */
  async createDriveItemLink(
    userId: string,
    itemId: string,
    type: 'view' | 'edit' | 'embed',
    scope: 'anonymous' | 'organization' = 'anonymous'
  ): Promise<{
    id: string;
    link: {
      type: string;
      scope: string;
      webUrl: string;
    };
  }> {
    try {
      return await this.client
        .api(`/users/${encodeURIComponent(userId)}/drive/items/${encodeURIComponent(itemId)}/createLink`)
        .post({
          type,
          scope
        });
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 邀请用户访问驱动器项（添加权限）
   * 
   * 权限要求：Files.ReadWrite, Files.ReadWrite.All
   * 
   * @param userId - 用户 ID 或 UPN
   * @param itemId - 项 ID
   * @param recipients - 收件人邮箱列表
   * @param roles - 角色：['read'] 或 ['write']
   * @param sendInvitation - 是否发送邀请邮件
   * @param message - 邀请消息
   * @returns 权限列表
   * @see https://learn.microsoft.com/graph/api/driveitem-invite
   */
  async inviteToDriveItem(
    userId: string,
    itemId: string,
    recipients: string[],
    roles: ('read' | 'write')[],
    sendInvitation: boolean = true,
    message?: string
  ): Promise<unknown[]> {
    try {
      const body: {
        recipients: Array<{ email: string }>;
        roles: string[];
        sendInvitation: boolean;
        message?: string;
      } = {
        recipients: recipients.map(email => ({ email })),
        roles,
        sendInvitation
      };
      if (message) {
        body.message = message;
      }
      const response = await this.client
        .api(`/users/${encodeURIComponent(userId)}/drive/items/${encodeURIComponent(itemId)}/invite`)
        .post(body);
      return response.value || [];
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 列出驱动器项权限
   * 
   * 权限要求：Files.Read, Files.ReadWrite, Files.Read.All, Files.ReadWrite.All
   * 
   * @param userId - 用户 ID 或 UPN
   * @param itemId - 项 ID
   * @returns 权限列表
   * @see https://learn.microsoft.com/graph/api/driveitem-list-permissions
   */
  async listDriveItemPermissions(userId: string, itemId: string): Promise<unknown[]> {
    try {
      const response = await this.client
        .api(`/users/${encodeURIComponent(userId)}/drive/items/${encodeURIComponent(itemId)}/permissions`)
        .get();
      return response.value || [];
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 删除驱动器项权限
   * 
   * 权限要求：Files.ReadWrite, Files.ReadWrite.All
   * 
   * @param userId - 用户 ID 或 UPN
   * @param itemId - 项 ID
   * @param permissionId - 权限 ID
   * @see https://learn.microsoft.com/graph/api/permission-delete
   */
  async deleteDriveItemPermission(userId: string, itemId: string, permissionId: string): Promise<void> {
    try {
      await this.client
        .api(`/users/${encodeURIComponent(userId)}/drive/items/${encodeURIComponent(itemId)}/permissions/${encodeURIComponent(permissionId)}`)
        .delete();
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 预览驱动器项
   * 
   * 获取项的可嵌入预览 URL
   * 
   * 权限要求：Files.Read, Files.ReadWrite, Files.Read.All, Files.ReadWrite.All
   * 
   * @param userId - 用户 ID 或 UPN
   * @param itemId - 项 ID
   * @returns 预览信息
   * @see https://learn.microsoft.com/graph/api/driveitem-preview
   */
  async previewDriveItem(userId: string, itemId: string): Promise<{
    getUrl: string;
    postUrl: string;
    postParameters: string;
  }> {
    try {
      return await this.client
        .api(`/users/${encodeURIComponent(userId)}/drive/items/${encodeURIComponent(itemId)}/preview`)
        .post({});
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 列出驱动器项版本
   * 
   * 权限要求：Files.Read, Files.ReadWrite, Files.Read.All, Files.ReadWrite.All
   * 
   * @param userId - 用户 ID 或 UPN
   * @param itemId - 项 ID
   * @returns 版本列表
   * @see https://learn.microsoft.com/graph/api/driveitem-list-versions
   */
  async listDriveItemVersions(userId: string, itemId: string): Promise<unknown[]> {
    try {
      const response = await this.client
        .api(`/users/${encodeURIComponent(userId)}/drive/items/${encodeURIComponent(itemId)}/versions`)
        .get();
      return response.value || [];
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 获取驱动器项分析数据
   * 
   * 权限要求：Files.Read, Files.ReadWrite, Files.Read.All, Files.ReadWrite.All
   * 
   * @param userId - 用户 ID 或 UPN
   * @param itemId - 项 ID
   * @returns 分析数据
   * @see https://learn.microsoft.com/graph/api/itemanalytics-get
   */
  async getDriveItemAnalytics(userId: string, itemId: string): Promise<unknown> {
    try {
      return await this.client
        .api(`/users/${encodeURIComponent(userId)}/drive/items/${encodeURIComponent(itemId)}/analytics`)
        .get();
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 获取驱动器项活动（按时间间隔）
   * 
   * 权限要求：Files.Read, Files.ReadWrite, Files.Read.All, Files.ReadWrite.All
   * 
   * @param userId - 用户 ID 或 UPN
   * @param itemId - 项 ID
   * @param startDateTime - 开始时间
   * @param endDateTime - 结束时间
   * @param interval - 时间间隔
   * @returns 活动统计
   * @see https://learn.microsoft.com/graph/api/itemactivitystat-getactivitybyinterval
   */
  async getDriveItemActivities(
    userId: string,
    itemId: string,
    startDateTime: string,
    endDateTime: string,
    interval: string = 'day'
  ): Promise<unknown[]> {
    try {
      const response = await this.client
        .api(`/users/${encodeURIComponent(userId)}/drive/items/${encodeURIComponent(itemId)}/analytics/allTime/getActivityByInterval(startDateTime='${startDateTime}',endDateTime='${endDateTime}',interval='${interval}')`)
        .get();
      return response.value || [];
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * SharePoint 站点管理
   * 
   * 权限要求：Sites.Read.All, Sites.ReadWrite.All
   * 国家云支持：✓ 全球版 ✓ 中国版 ✓ 美国政府版
   * 
   * @see https://learn.microsoft.com/graph/api/resources/site
   */
  get sites() {
    return new GraphEntityManager<Site>(this.client, '/sites');
  }

  /**
   * 根站点
   * 
   * 获取 SharePoint 租户的根站点
   */
  get rootSite() {
    return this.client.api('/sites/root').get();
  }

  /**
   * 通过站点 ID 或路径获取站点的驱动器项
   * 
   * @param siteId - 站点 ID
   */
  siteDriveItems(siteId: string) {
    return new GraphEntityManager<DriveItem>(
      this.client, 
      `/sites/${encodeURIComponent(siteId)}/drive/root/children`
    );
  }

  // ==================== 批处理请求 ====================

  /**
   * 执行批处理请求
   * 
   * 允许在单个 HTTP 请求中发送多个 API 调用
   * 最多支持 20 个请求
   * 
   * 权限要求：根据批处理中的各个请求而定
   * 国家云支持：✓ 全球版 ✓ 中国版 ✓ 美国政府版
   * 
   * @param requests - 批处理请求数组
   * @returns 批处理响应
   * @see https://learn.microsoft.com/graph/json-batching
   * 
   * @example
   * ```typescript
   * const batchResponse = await orm.batch([
   *   { id: '1', method: 'GET', url: '/users/user1@contoso.com' },
   *   { id: '2', method: 'GET', url: '/users/user2@contoso.com' }
   * ]);
   * ```
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
    try {
      if (requests.length > 20) {
        throw GraphOrmError.create(
          GraphErrorCode.INVALID_PARAMETER,
          '批处理请求最多支持 20 个请求'
        );
      }

      const response = await this.client
        .api('/$batch')
        .post({ requests });

      return response;
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  // ==================== Microsoft Teams ====================

  /**
   * Teams 团队管理
   * 
   * 权限要求：Team.ReadBasic.All, TeamSettings.Read.All, TeamSettings.ReadWrite.All
   * 国家云支持：✓ 全球版 ⚠️ 中国版（部分功能受限）✓ 美国政府版（GCC）
   * 
   * 注意：
   * - 中国版：Teams 功能可能受限
   * - GCC High 和 DoD：部分 Teams 功能不可用
   * 
   * @see https://learn.microsoft.com/graph/api/resources/team
   */
  get teams() {
    return new GraphEntityManager<Team>(this.client, '/teams');
  }

  /**
   * 获取组关联的团队
   * 
   * @param groupId - 组 ID
   */
  async getTeamByGroup(groupId: string): Promise<Team> {
    try {
      return await this.client
        .api(`/groups/${encodeURIComponent(groupId)}/team`)
        .get();
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 团队频道管理
   * 
   * @param teamId - 团队 ID
   */
  channels(teamId: string) {
    return new GraphEntityManager<Channel>(
      this.client,
      `/teams/${encodeURIComponent(teamId)}/channels`
    );
  }

  /**
   * 用户聊天管理
   * 
   * 权限要求：Chat.Read, Chat.ReadWrite
   * 
   * @param userId - 用户 ID
   */
  chats(userId: string) {
    return new GraphEntityManager<Chat>(
      this.client,
      `/users/${encodeURIComponent(userId)}/chats`
    );
  }

  /**
   * 获取所有聊天（需要应用权限）
   */
  get allChats() {
    return new GraphEntityManager<Chat>(this.client, '/chats');
  }

  // ==================== Planner ====================

  /**
   * Planner 计划管理
   * 
   * 权限要求：Tasks.Read, Tasks.ReadWrite
   * 国家云支持：✓ 全球版 ⚠️ 中国版（不可用）✓ 美国政府版（GCC）
   * 
   * 注意：
   * - 中国版：Planner API 不可用
   * - GCC High 和 DoD：部分功能受限
   * 
   * @see https://learn.microsoft.com/graph/api/resources/planner-overview
   */
  get plans() {
    return new GraphEntityManager<Plan>(this.client, '/planner/plans');
  }

  /**
   * 组的计划
   * 
   * @param groupId - 组 ID
   */
  groupPlans(groupId: string) {
    return new GraphEntityManager<Plan>(
      this.client,
      `/groups/${encodeURIComponent(groupId)}/planner/plans`
    );
  }

  /**
   * Planner 任务管理
   * 
   * @param planId - 计划 ID
   */
  planTasks(planId: string) {
    return new GraphEntityManager<Task>(
      this.client,
      `/planner/plans/${encodeURIComponent(planId)}/tasks`
    );
  }

  /**
   * 用户的所有任务
   * 
   * @param userId - 用户 ID
   */
  userTasks(userId: string) {
    return new GraphEntityManager<Task>(
      this.client,
      `/users/${encodeURIComponent(userId)}/planner/tasks`
    );
  }

  // ==================== Delta Query（增量查询）====================

  /**
   * 用户增量查询
   * 
   * 跟踪用户的变化（创建、更新、删除）
   * 
   * 权限要求：User.Read.All
   * 国家云支持：✓ 全球版 ✓ 中国版 ✓ 美国政府版
   * 
   * @param deltaLink - 上次查询返回的 deltaLink（可选，首次查询传空）
   * @returns Delta 查询结果，包含 deltaLink 用于下次查询
   * 
   * @see https://learn.microsoft.com/graph/delta-query-users
   * 
   * @example
   * ```typescript
   * // 首次查询
   * const result1 = await orm.deltaUsers();
   * console.log('变化的用户:', result1.data);
   * 
   * // 保存 deltaLink
   * const deltaLink = result1.meta.deltaLink;
   * 
   * // 后续查询（只返回变化的数据）
   * const result2 = await orm.deltaUsers(deltaLink);
   * console.log('新的变化:', result2.data);
   * ```
   */
  async deltaUsers(deltaLink?: string): Promise<DeltaCollection<User>> {
    try {
      const endpoint = deltaLink || '/users/delta';
      const response = await this.client.api(endpoint).get();
      
      return {
        meta: {
          deltaLink: response['@odata.deltaLink'],
          nextLink: response['@odata.nextLink'],
          count: response['@odata.count'],
        },
        data: response.value || []
      };
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 组增量查询
   * 
   * @param deltaLink - 上次查询返回的 deltaLink
   * @see https://learn.microsoft.com/graph/delta-query-groups
   */
  async deltaGroups(deltaLink?: string): Promise<DeltaCollection<Group>> {
    try {
      const endpoint = deltaLink || '/groups/delta';
      const response = await this.client.api(endpoint).get();
      
      return {
        meta: {
          deltaLink: response['@odata.deltaLink'],
          nextLink: response['@odata.nextLink'],
          count: response['@odata.count'],
        },
        data: response.value || []
      };
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 邮件增量查询
   * 
   * @param userId - 用户 ID
   * @param deltaLink - 上次查询返回的 deltaLink
   * @see https://learn.microsoft.com/graph/delta-query-messages
   */
  async deltaMessages(userId: string, deltaLink?: string): Promise<DeltaCollection<Message>> {
    try {
      const endpoint = deltaLink || `/users/${encodeURIComponent(userId)}/messages/delta`;
      const response = await this.client.api(endpoint).get();
      
      return {
        meta: {
          deltaLink: response['@odata.deltaLink'],
          nextLink: response['@odata.nextLink'],
          count: response['@odata.count'],
        },
        data: response.value || []
      };
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 事件增量查询
   * 
   * @param userId - 用户 ID
   * @param deltaLink - 上次查询返回的 deltaLink
   * @see https://learn.microsoft.com/graph/delta-query-events
   */
  async deltaEvents(userId: string, deltaLink?: string): Promise<DeltaCollection<Event>> {
    try {
      const endpoint = deltaLink || `/users/${encodeURIComponent(userId)}/events/delta`;
      const response = await this.client.api(endpoint).get();
      
      return {
        meta: {
          deltaLink: response['@odata.deltaLink'],
          nextLink: response['@odata.nextLink'],
          count: response['@odata.count'],
        },
        data: response.value || []
      };
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  // ==================== Search API ====================

  /**
   * Microsoft Search API
   * 
   * 跨多个数据源（邮件、事件、文件、Teams消息等）进行搜索
   * 
   * 权限要求：根据搜索的资源类型而定
   * - 邮件：Mail.Read
   * - 事件：Calendars.Read
   * - 文件：Files.Read.All
   * - Teams消息：ChannelMessage.Read.All
   * 
   * 国家云支持：✓ 全球版 ⚠️ 中国版（不可用）✓ 美国政府版（GCC）
   * 
   * @param searchRequest - 搜索请求配置
   * @returns 搜索结果
   * 
   * @see https://learn.microsoft.com/graph/api/search-query
   * @see https://learn.microsoft.com/graph/search-concept-overview
   * 
   * @example
   * ```typescript
   * // 搜索邮件和事件
   * const results = await orm.search({
   *   entityTypes: ['message', 'event'],
   *   query: { queryString: '项目会议' },
   *   from: 0,
   *   size: 25
   * });
   * 
   * // 搜索文件
   * const fileResults = await orm.search({
   *   entityTypes: ['driveItem'],
   *   query: { queryString: '财务报告' },
   *   from: 0,
   *   size: 10
   * });
   * ```
   */
  async search(searchRequest: SearchRequest): Promise<SearchResponse> {
    try {
      const response = await this.client
        .api('/search/query')
        .post({
          requests: [searchRequest]
        });
      
      return response.value[0];
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 快速搜索用户
   * 
   * @param queryString - 搜索关键词
   * @param size - 返回结果数量
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
   * 
   * @param queryString - 搜索关键词
   * @param size - 返回结果数量
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
   * 
   * @param queryString - 搜索关键词
   * @param size - 返回结果数量
   */
  async searchFiles(queryString: string, size: number = 25): Promise<SearchResponse> {
    return this.search({
      entityTypes: ['driveItem'],
      query: { queryString },
      size
    });
  }

  // ==================== Subscriptions（Webhooks）====================

  /**
   * 订阅管理（Webhooks）
   * 
   * 创建订阅以接收资源变化通知
   * 
   * 权限要求：根据订阅的资源类型而定
   * 国家云支持：✓ 全球版 ✓ 中国版 ✓ 美国政府版
   * 
   * @see https://learn.microsoft.com/graph/api/resources/subscription
   * @see https://learn.microsoft.com/graph/webhooks
   */
  get subscriptions() {
    return new GraphEntityManager<Subscription>(this.client, '/subscriptions');
  }

  /**
   * 创建订阅
   * 
   * @param subscription - 订阅配置
   * @returns 创建的订阅
   * 
   * @example
   * ```typescript
   * // 订阅用户邮件变化
   * const subscription = await orm.createSubscription({
   *   changeType: 'created,updated',
   *   notificationUrl: 'https://your-app.com/webhooks',
   *   resource: 'users/user@contoso.com/messages',
   *   expirationDateTime: new Date(Date.now() + 3600000).toISOString(), // 1小时后过期
   *   clientState: 'secret-state-value'
   * });
   * ```
   */
  async createSubscription(subscription: Omit<Subscription, 'id'>): Promise<Subscription> {
    try {
      return await this.client
        .api('/subscriptions')
        .post(subscription);
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 更新订阅（通常用于延长过期时间）
   * 
   * @param subscriptionId - 订阅 ID
   * @param expirationDateTime - 新的过期时间
   */
  async renewSubscription(subscriptionId: string, expirationDateTime: string): Promise<Subscription> {
    try {
      return await this.client
        .api(`/subscriptions/${encodeURIComponent(subscriptionId)}`)
        .patch({ expirationDateTime });
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 删除订阅
   * 
   * @param subscriptionId - 订阅 ID
   */
  async deleteSubscription(subscriptionId: string): Promise<void> {
    try {
      await this.client
        .api(`/subscriptions/${encodeURIComponent(subscriptionId)}`)
        .delete();
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  // ==================== 目录角色和管理单元 ====================

  /**
   * 目录角色管理
   * 
   * 权限要求：RoleManagement.Read.Directory, RoleManagement.ReadWrite.Directory
   * 国家云支持：✓ 全球版 ✓ 中国版 ✓ 美国政府版
   * 
   * @see https://learn.microsoft.com/graph/api/resources/directoryrole
   */
  get directoryRoles() {
    return new GraphEntityManager<DirectoryRole>(this.client, '/directoryRoles');
  }

  /**
   * 目录角色模板管理
   * 
   * 权限要求：RoleManagement.Read.Directory
   * 国家云支持：✓ 全球版 ✓ 中国版 ✓ 美国政府版
   * 
   * @see https://learn.microsoft.com/graph/api/resources/directoryroletemplate
   */
  get directoryRoleTemplates() {
    return new GraphEntityManager<DirectoryRoleTemplate>(this.client, '/directoryRoleTemplates');
  }

  /**
   * 获取目录角色成员
   * 
   * @param roleId - 角色 ID
   */
  directoryRoleMembers(roleId: string) {
    return new GraphEntityManager<User>(
      this.client,
      `/directoryRoles/${encodeURIComponent(roleId)}/members`
    );
  }

  /**
   * 管理单元管理
   * 
   * 权限要求：AdministrativeUnit.Read.All, AdministrativeUnit.ReadWrite.All
   * 国家云支持：✓ 全球版 ✓ 中国版 ✓ 美国政府版
   * 
   * @see https://learn.microsoft.com/graph/api/resources/administrativeunit
   */
  get administrativeUnits() {
    return new GraphEntityManager<AdministrativeUnit>(this.client, '/directory/administrativeUnits');
  }

  /**
   * 获取管理单元成员
   * 
   * @param unitId - 管理单元 ID
   */
  administrativeUnitMembers(unitId: string) {
    return new GraphEntityManager<User>(
      this.client,
      `/directory/administrativeUnits/${encodeURIComponent(unitId)}/members`
    );
  }

  // ==================== 许可证管理 ====================

  /**
   * 获取用户许可证详情
   * 
   * 权限要求：User.Read.All, Directory.Read.All
   * 国家云支持：✓ 全球版 ✓ 中国版 ✓ 美国政府版
   * 
   * @param userId - 用户 ID
   * @returns 许可证详情列表
   * @see https://learn.microsoft.com/graph/api/user-list-licensedetails
   */
  async getUserLicenses(userId: string): Promise<LicenseDetails[]> {
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
   * 获取订阅的 SKU
   * 
   * 权限要求：Organization.Read.All, Directory.Read.All
   * 国家云支持：✓ 全球版 ✓ 中国版 ✓ 美国政府版
   * 
   * @returns 订阅的 SKU 列表
   * @see https://learn.microsoft.com/graph/api/subscribedsku-list
   */
  async getSubscribedSkus(): Promise<SubscribedSku[]> {
    try {
      const response = await this.client.api('/subscribedSkus').get();
      return response.value || [];
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 为用户分配许可证
   * 
   * 权限要求：User.ReadWrite.All, Directory.ReadWrite.All
   * 
   * @param userId - 用户 ID
   * @param addLicenses - 要添加的许可证
   * @param removeLicenses - 要移除的许可证（SKU ID）
   * @returns 更新后的用户
   * @see https://learn.microsoft.com/graph/api/user-assignlicense
   */
  async assignUserLicense(
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

  // ==================== 邀请管理 ====================

  /**
   * 创建邀请
   * 
   * 邀请外部用户加入组织
   * 
   * 权限要求：User.Invite.All
   * 国家云支持：✓ 全球版 ⚠️ 中国版（功能受限）✓ 美国政府版
   * 
   * @param invitation - 邀请信息
   * @returns 创建的邀请
   * @see https://learn.microsoft.com/graph/api/invitation-post
   * 
   * @example
   * ```typescript
   * const invitation = await orm.createInvitation({
   *   invitedUserEmailAddress: 'external@example.com',
   *   inviteRedirectUrl: 'https://myapp.com',
   *   sendInvitationMessage: true
   * });
   * ```
   */
  async createInvitation(invitation: Omit<Invitation, 'id'>): Promise<Invitation> {
    try {
      return await this.client
        .api('/invitations')
        .post(invitation);
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  // ==================== 日历增强 ====================

  /**
   * 用户日历管理
   * 
   * 权限要求：Calendars.Read, Calendars.ReadWrite
   * 国家云支持：✓ 全球版 ✓ 中国版 ✓ 美国政府版
   * 
   * @param userId - 用户 ID 或 UPN
   * @see https://learn.microsoft.com/graph/api/resources/calendar
   */
  calendars(userId: string) {
    return new GraphEntityManager<Calendar>(
      this.client,
      `/users/${encodeURIComponent(userId)}/calendars`
    );
  }

  /**
   * 获取用户主日历
   * 
   * @param userId - 用户 ID 或 UPN
   * @returns 主日历
   */
  async getUserCalendar(userId: string): Promise<Calendar> {
    try {
      return await this.client
        .api(`/users/${encodeURIComponent(userId)}/calendar`)
        .get();
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 用户日历组管理
   * 
   * 权限要求：Calendars.Read, Calendars.ReadWrite
   * 国家云支持：✓ 全球版 ✓ 中国版 ✓ 美国政府版
   * 
   * @param userId - 用户 ID 或 UPN
   * @see https://learn.microsoft.com/graph/api/resources/calendargroup
   */
  calendarGroups(userId: string) {
    return new GraphEntityManager<CalendarGroup>(
      this.client,
      `/users/${encodeURIComponent(userId)}/calendarGroups`
    );
  }

  /**
   * 日历组的日历
   * 
   * @param userId - 用户 ID 或 UPN
   * @param groupId - 日历组 ID
   */
  calendarsInGroup(userId: string, groupId: string) {
    return new GraphEntityManager<Calendar>(
      this.client,
      `/users/${encodeURIComponent(userId)}/calendarGroups/${encodeURIComponent(groupId)}/calendars`
    );
  }

  /**
   * 获取日历视图（指定时间范围的事件）
   * 
   * @param userId - 用户 ID 或 UPN
   * @param startDateTime - 开始时间（ISO 8601）
   * @param endDateTime - 结束时间（ISO 8601）
   * @returns 事件列表
   * @see https://learn.microsoft.com/graph/api/calendar-list-calendarview
   */
  async getCalendarView(
    userId: string,
    startDateTime: string,
    endDateTime: string
  ): Promise<Event[]> {
    try {
      const response = await this.client
        .api(`/users/${encodeURIComponent(userId)}/calendar/calendarView`)
        .query({
          startDateTime,
          endDateTime
        })
        .get();
      return response.value || [];
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 获取忙/闲时间表
   * 
   * 权限要求：Calendars.Read, Calendars.Read.Shared
   * 
   * @param schedules - 要查询的用户或资源邮箱列表
   * @param startTime - 开始时间
   * @param endTime - 结束时间
   * @param availabilityViewInterval - 时间间隔（分钟，默认30）
   * @returns 忙/闲信息
   * @see https://learn.microsoft.com/graph/api/calendar-getschedule
   */
  async getSchedule(
    schedules: string[],
    startTime: { dateTime: string; timeZone: string },
    endTime: { dateTime: string; timeZone: string },
    availabilityViewInterval: number = 30
  ): Promise<unknown> {
    try {
      return await this.client
        .api('/me/calendar/getSchedule')
        .post({
          schedules,
          startTime,
          endTime,
          availabilityViewInterval
        });
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  // ==================== 邮件增强 ====================

  /**
   * 用户邮件文件夹管理
   * 
   * 权限要求：Mail.Read, Mail.ReadWrite
   * 国家云支持：✓ 全球版 ✓ 中国版 ✓ 美国政府版
   * 
   * @param userId - 用户 ID 或 UPN
   * @see https://learn.microsoft.com/graph/api/resources/mailfolder
   */
  mailFolders(userId: string) {
    return new GraphEntityManager<MailFolder>(
      this.client,
      `/users/${encodeURIComponent(userId)}/mailFolders`
    );
  }

  /**
   * 获取指定邮件文件夹的邮件
   * 
   * @param userId - 用户 ID 或 UPN
   * @param folderId - 文件夹 ID 或知名文件夹名称（如 'inbox', 'drafts', 'sentitems'）
   */
  messagesInFolder(userId: string, folderId: string) {
    return new GraphEntityManager<Message>(
      this.client,
      `/users/${encodeURIComponent(userId)}/mailFolders/${encodeURIComponent(folderId)}/messages`
    );
  }

  /**
   * 用户邮件规则管理
   * 
   * 权限要求：MailboxSettings.ReadWrite
   * 国家云支持：✓ 全球版 ✓ 中国版 ✓ 美国政府版
   * 
   * @param userId - 用户 ID 或 UPN
   * @see https://learn.microsoft.com/graph/api/resources/messagerule
   */
  messageRules(userId: string) {
    return new GraphEntityManager<MessageRule>(
      this.client,
      `/users/${encodeURIComponent(userId)}/mailFolders/inbox/messageRules`
    );
  }

  /**
   * 获取用户邮箱设置
   * 
   * 权限要求：MailboxSettings.Read, MailboxSettings.ReadWrite
   * 国家云支持：✓ 全球版 ✓ 中国版 ✓ 美国政府版
   * 
   * @param userId - 用户 ID 或 UPN
   * @returns 邮箱设置
   * @see https://learn.microsoft.com/graph/api/user-get-mailboxsettings
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
   * 
   * @param userId - 用户 ID 或 UPN
   * @param settings - 邮箱设置
   * @returns 更新后的邮箱设置
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
   * 获取重点收件箱设置
   * 
   * 权限要求：Mail.Read, Mail.ReadWrite
   * 国家云支持：✓ 全球版 ✓ 中国版 ✓ 美国政府版
   * 
   * @param userId - 用户 ID 或 UPN
   * @returns 重点收件箱设置
   * @see https://learn.microsoft.com/graph/api/inferenceclassification-get
   */
  async getInferenceClassification(userId: string): Promise<InferenceClassification> {
    try {
      return await this.client
        .api(`/users/${encodeURIComponent(userId)}/inferenceClassification`)
        .get();
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 获取邮件附件
   * 
   * 权限要求：Mail.Read, Mail.ReadWrite
   * 
   * @param userId - 用户 ID 或 UPN
   * @param messageId - 邮件 ID
   * @returns 附件列表
   * @see https://learn.microsoft.com/graph/api/message-list-attachments
   */
  async getMessageAttachments(userId: string, messageId: string): Promise<Attachment[]> {
    try {
      const response = await this.client
        .api(`/users/${encodeURIComponent(userId)}/messages/${encodeURIComponent(messageId)}/attachments`)
        .get();
      return response.value || [];
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 添加邮件附件
   * 
   * 权限要求：Mail.ReadWrite
   * 
   * @param userId - 用户 ID 或 UPN
   * @param messageId - 邮件 ID
   * @param attachment - 附件信息
   * @returns 创建的附件
   * @see https://learn.microsoft.com/graph/api/message-post-attachments
   */
  async addMessageAttachment(
    userId: string,
    messageId: string,
    attachment: Omit<FileAttachment, 'id'>
  ): Promise<Attachment> {
    try {
      return await this.client
        .api(`/users/${encodeURIComponent(userId)}/messages/${encodeURIComponent(messageId)}/attachments`)
        .post(attachment);
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 发送邮件
   * 
   * 权限要求：Mail.Send
   * 
   * @param userId - 用户 ID 或 UPN
   * @param message - 邮件内容
   * @param saveToSentItems - 是否保存到已发送项目
   * @see https://learn.microsoft.com/graph/api/user-sendmail
   */
  async sendMail(userId: string, message: Partial<Message>, saveToSentItems: boolean = true): Promise<void> {
    try {
      await this.client
        .api(`/users/${encodeURIComponent(userId)}/sendMail`)
        .post({
          message,
          saveToSentItems
        });
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  // ==================== 联系人增强 ====================

  /**
   * 用户联系人文件夹管理
   * 
   * 权限要求：Contacts.Read, Contacts.ReadWrite
   * 国家云支持：✓ 全球版 ✓ 中国版 ✓ 美国政府版
   * 
   * @param userId - 用户 ID 或 UPN
   * @see https://learn.microsoft.com/graph/api/resources/contactfolder
   */
  contactFolders(userId: string) {
    return new GraphEntityManager<ContactFolder>(
      this.client,
      `/users/${encodeURIComponent(userId)}/contactFolders`
    );
  }

  /**
   * 组织联系人管理
   * 
   * 权限要求：OrgContact.Read.All
   * 国家云支持：✓ 全球版 ✓ 中国版 ✓ 美国政府版
   * 
   * @see https://learn.microsoft.com/graph/api/resources/orgcontact
   */
  get orgContacts() {
    return new GraphEntityManager<OrgContact>(this.client, '/contacts');
  }

  // ==================== 文件存储增强 ====================

  /**
   * 驱动器管理
   * 
   * 权限要求：Files.Read, Files.ReadWrite, Files.Read.All, Files.ReadWrite.All
   * 国家云支持：✓ 全球版 ✓ 中国版 ✓ 美国政府版
   * 
   * @see https://learn.microsoft.com/graph/api/resources/drive
   */
  get drives() {
    return new GraphEntityManager<Drive>(this.client, '/drives');
  }

  /**
   * 获取用户驱动器
   * 
   * @param userId - 用户 ID 或 UPN
   * @returns 用户驱动器
   */
  async getUserDrive(userId: string): Promise<Drive> {
    try {
      return await this.client
        .api(`/users/${encodeURIComponent(userId)}/drive`)
        .get();
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 获取组驱动器
   * 
   * @param groupId - 组 ID
   * @returns 组驱动器
   */
  async getGroupDrive(groupId: string): Promise<Drive> {
    try {
      return await this.client
        .api(`/groups/${encodeURIComponent(groupId)}/drive`)
        .get();
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 获取与我共享的项目
   * 
   * 权限要求：Files.Read.All, Files.ReadWrite.All
   * 
   * @returns 共享项目列表
   * @see https://learn.microsoft.com/graph/api/drive-sharedwithme
   */
  async getSharedWithMe(): Promise<DriveItem[]> {
    try {
      const response = await this.client
        .api('/me/drive/sharedWithMe')
        .get();
      return response.value || [];
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 获取最近使用的文件
   * 
   * 权限要求：Files.Read, Files.ReadWrite
   * 
   * @returns 最近使用的文件列表
   * @see https://learn.microsoft.com/graph/api/drive-recent
   */
  async getRecentFiles(): Promise<DriveItem[]> {
    try {
      const response = await this.client
        .api('/me/drive/recent')
        .get();
      return response.value || [];
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  // ==================== OneNote ====================

  /**
   * 用户 OneNote 笔记本管理
   * 
   * 权限要求：Notes.Read, Notes.ReadWrite, Notes.Read.All, Notes.ReadWrite.All
   * 国家云支持：✓ 全球版 ⚠️ 中国版（不可用）✓ 美国政府版
   * 
   * @param userId - 用户 ID 或 UPN
   * @see https://learn.microsoft.com/graph/api/resources/notebook
   */
  notebooks(userId: string) {
    return new GraphEntityManager<Notebook>(
      this.client,
      `/users/${encodeURIComponent(userId)}/onenote/notebooks`
    );
  }

  /**
   * 组 OneNote 笔记本
   * 
   * @param groupId - 组 ID
   */
  groupNotebooks(groupId: string) {
    return new GraphEntityManager<Notebook>(
      this.client,
      `/groups/${encodeURIComponent(groupId)}/onenote/notebooks`
    );
  }

  /**
   * OneNote 分区管理
   * 
   * @param userId - 用户 ID 或 UPN
   * @see https://learn.microsoft.com/graph/api/resources/section
   */
  sections(userId: string) {
    return new GraphEntityManager<Section>(
      this.client,
      `/users/${encodeURIComponent(userId)}/onenote/sections`
    );
  }

  /**
   * 笔记本的分区
   * 
   * @param userId - 用户 ID 或 UPN
   * @param notebookId - 笔记本 ID
   */
  notebookSections(userId: string, notebookId: string) {
    return new GraphEntityManager<Section>(
      this.client,
      `/users/${encodeURIComponent(userId)}/onenote/notebooks/${encodeURIComponent(notebookId)}/sections`
    );
  }

  /**
   * OneNote 页面管理
   * 
   * @param userId - 用户 ID 或 UPN
   * @see https://learn.microsoft.com/graph/api/resources/page
   */
  oneNotePages(userId: string) {
    return new GraphEntityManager<OneNotePage>(
      this.client,
      `/users/${encodeURIComponent(userId)}/onenote/pages`
    );
  }

  /**
   * 分区的页面
   * 
   * @param userId - 用户 ID 或 UPN
   * @param sectionId - 分区 ID
   */
  sectionPages(userId: string, sectionId: string) {
    return new GraphEntityManager<OneNotePage>(
      this.client,
      `/users/${encodeURIComponent(userId)}/onenote/sections/${encodeURIComponent(sectionId)}/pages`
    );
  }

  /**
   * 创建 OneNote 页面
   * 
   * 权限要求：Notes.ReadWrite, Notes.ReadWrite.All
   * 
   * @param userId - 用户 ID 或 UPN
   * @param sectionId - 分区 ID
   * @param htmlContent - 页面 HTML 内容
   * @returns 创建的页面
   * @see https://learn.microsoft.com/graph/api/section-post-pages
   */
  async createOneNotePage(userId: string, sectionId: string, htmlContent: string): Promise<OneNotePage> {
    try {
      return await this.client
        .api(`/users/${encodeURIComponent(userId)}/onenote/sections/${encodeURIComponent(sectionId)}/pages`)
        .header('Content-Type', 'text/html')
        .post(htmlContent);
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  // ==================== Planner 增强 ====================

  /**
   * 计划的存储桶管理
   * 
   * 权限要求：Tasks.Read, Tasks.ReadWrite, Group.Read.All
   * 国家云支持：✓ 全球版 ⚠️ 中国版（不可用）✓ 美国政府版（GCC）
   * 
   * @param planId - 计划 ID
   * @see https://learn.microsoft.com/graph/api/resources/plannerbucket
   */
  planBuckets(planId: string) {
    return new GraphEntityManager<Bucket>(
      this.client,
      `/planner/plans/${encodeURIComponent(planId)}/buckets`
    );
  }

  /**
   * 存储桶的任务
   * 
   * @param bucketId - 存储桶 ID
   */
  bucketTasks(bucketId: string) {
    return new GraphEntityManager<Task>(
      this.client,
      `/planner/buckets/${encodeURIComponent(bucketId)}/tasks`
    );
  }

  // ==================== To Do ====================

  /**
   * 用户 To Do 任务列表管理
   * 
   * 权限要求：Tasks.Read, Tasks.ReadWrite
   * 国家云支持：✓ 全球版 ⚠️ 中国版（不可用）✓ 美国政府版
   * 
   * @param userId - 用户 ID 或 UPN
   * @see https://learn.microsoft.com/graph/api/resources/todotasklist
   */
  todoLists(userId: string) {
    return new GraphEntityManager<TodoTaskList>(
      this.client,
      `/users/${encodeURIComponent(userId)}/todo/lists`
    );
  }

  /**
   * To Do 列表的任务
   * 
   * @param userId - 用户 ID 或 UPN
   * @param listId - 列表 ID
   * @see https://learn.microsoft.com/graph/api/resources/todotask
   */
  todoTasks(userId: string, listId: string) {
    return new GraphEntityManager<TodoTask>(
      this.client,
      `/users/${encodeURIComponent(userId)}/todo/lists/${encodeURIComponent(listId)}/tasks`
    );
  }

  // ==================== SharePoint Lists ====================

  /**
   * 站点列表管理
   * 
   * 权限要求：Sites.Read.All, Sites.ReadWrite.All
   * 国家云支持：✓ 全球版 ✓ 中国版 ✓ 美国政府版
   * 
   * @param siteId - 站点 ID
   * @see https://learn.microsoft.com/graph/api/resources/list
   */
  siteLists(siteId: string) {
    return new GraphEntityManager<List>(
      this.client,
      `/sites/${encodeURIComponent(siteId)}/lists`
    );
  }

  /**
   * 列表项管理
   * 
   * @param siteId - 站点 ID
   * @param listId - 列表 ID
   * @see https://learn.microsoft.com/graph/api/resources/listitem
   */
  listItems(siteId: string, listId: string) {
    return new GraphEntityManager<ListItem>(
      this.client,
      `/sites/${encodeURIComponent(siteId)}/lists/${encodeURIComponent(listId)}/items`
    );
  }

  /**
   * 列表列定义管理
   * 
   * @param siteId - 站点 ID
   * @param listId - 列表 ID
   * @see https://learn.microsoft.com/graph/api/resources/columndefinition
   */
  listColumns(siteId: string, listId: string) {
    return new GraphEntityManager<ColumnDefinition>(
      this.client,
      `/sites/${encodeURIComponent(siteId)}/lists/${encodeURIComponent(listId)}/columns`
    );
  }

  /**
   * 通过路径获取站点
   * 
   * @param hostname - 主机名（如 'contoso.sharepoint.com'）
   * @param sitePath - 站点路径（如 '/sites/teamsite'）
   * @returns 站点信息
   * @see https://learn.microsoft.com/graph/api/site-getbypath
   */
  async getSiteByPath(hostname: string, sitePath: string): Promise<Site> {
    try {
      return await this.client
        .api(`/sites/${hostname}:${sitePath}`)
        .get();
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  // ==================== Teams 消息 ====================

  /**
   * 频道消息管理
   * 
   * 权限要求：ChannelMessage.Read.All, Group.Read.All
   * 国家云支持：✓ 全球版 ⚠️ 中国版（功能受限）✓ 美国政府版（GCC）
   * 
   * @param teamId - 团队 ID
   * @param channelId - 频道 ID
   * @see https://learn.microsoft.com/graph/api/resources/chatmessage
   */
  channelMessages(teamId: string, channelId: string) {
    return new GraphEntityManager<ChatMessage>(
      this.client,
      `/teams/${encodeURIComponent(teamId)}/channels/${encodeURIComponent(channelId)}/messages`
    );
  }

  /**
   * 频道消息回复
   * 
   * @param teamId - 团队 ID
   * @param channelId - 频道 ID
   * @param messageId - 消息 ID
   */
  channelMessageReplies(teamId: string, channelId: string, messageId: string) {
    return new GraphEntityManager<ChatMessage>(
      this.client,
      `/teams/${encodeURIComponent(teamId)}/channels/${encodeURIComponent(channelId)}/messages/${encodeURIComponent(messageId)}/replies`
    );
  }

  /**
   * 聊天消息管理
   * 
   * 权限要求：ChatMessage.Read, Chat.ReadWrite
   * 
   * @param chatId - 聊天 ID
   */
  chatMessages(chatId: string) {
    return new GraphEntityManager<ChatMessage>(
      this.client,
      `/chats/${encodeURIComponent(chatId)}/messages`
    );
  }

  /**
   * 发送频道消息
   * 
   * 权限要求：ChannelMessage.Send
   * 
   * @param teamId - 团队 ID
   * @param channelId - 频道 ID
   * @param message - 消息内容
   * @returns 发送的消息
   * @see https://learn.microsoft.com/graph/api/channel-post-messages
   */
  async sendChannelMessage(teamId: string, channelId: string, message: Omit<ChatMessage, 'id'>): Promise<ChatMessage> {
    try {
      return await this.client
        .api(`/teams/${encodeURIComponent(teamId)}/channels/${encodeURIComponent(channelId)}/messages`)
        .post(message);
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 发送聊天消息
   * 
   * 权限要求：ChatMessage.Send
   * 
   * @param chatId - 聊天 ID
   * @param message - 消息内容
   * @returns 发送的消息
   * @see https://learn.microsoft.com/graph/api/chat-post-messages
   */
  async sendChatMessage(chatId: string, message: Omit<ChatMessage, 'id'>): Promise<ChatMessage> {
    try {
      return await this.client
        .api(`/chats/${encodeURIComponent(chatId)}/messages`)
        .post(message);
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 团队成员管理
   * 
   * 权限要求：TeamMember.Read.All, TeamMember.ReadWrite.All
   * 
   * @param teamId - 团队 ID
   * @see https://learn.microsoft.com/graph/api/resources/conversationmember
   */
  teamMembers(teamId: string) {
    return new GraphEntityManager<ConversationMember>(
      this.client,
      `/teams/${encodeURIComponent(teamId)}/members`
    );
  }

  /**
   * 频道成员管理
   * 
   * @param teamId - 团队 ID
   * @param channelId - 频道 ID
   */
  channelMembers(teamId: string, channelId: string) {
    return new GraphEntityManager<ConversationMember>(
      this.client,
      `/teams/${encodeURIComponent(teamId)}/channels/${encodeURIComponent(channelId)}/members`
    );
  }

  /**
   * 团队已安装应用管理
   * 
   * 权限要求：TeamsApp.Read.All, TeamsAppInstallation.ReadForTeam
   * 
   * @param teamId - 团队 ID
   * @see https://learn.microsoft.com/graph/api/resources/teamsappinstallation
   */
  teamApps(teamId: string) {
    return new GraphEntityManager<TeamsAppInstallation>(
      this.client,
      `/teams/${encodeURIComponent(teamId)}/installedApps`
    );
  }

  /**
   * 团队标签管理
   * 
   * 权限要求：TeamworkTag.Read, TeamworkTag.ReadWrite
   * 
   * @param teamId - 团队 ID
   * @see https://learn.microsoft.com/graph/api/resources/teamworktag
   */
  teamTags(teamId: string) {
    return new GraphEntityManager<TeamworkTag>(
      this.client,
      `/teams/${encodeURIComponent(teamId)}/tags`
    );
  }

  // ==================== 报告和分析 ====================

  /**
   * 审计日志 - 目录审计
   * 
   * 权限要求：AuditLog.Read.All, Directory.Read.All
   * 国家云支持：✓ 全球版 ✓ 中国版 ✓ 美国政府版
   * 
   * @see https://learn.microsoft.com/graph/api/resources/directoryaudit
   */
  get directoryAudits() {
    return new GraphEntityManager<DirectoryAudit>(this.client, '/auditLogs/directoryAudits');
  }

  /**
   * 审计日志 - 登录日志
   * 
   * 权限要求：AuditLog.Read.All, Directory.Read.All
   * 国家云支持：✓ 全球版 ✓ 中国版 ✓ 美国政府版
   * 
   * @see https://learn.microsoft.com/graph/api/resources/signin
   */
  get signInLogs() {
    return new GraphEntityManager<SignIn>(this.client, '/auditLogs/signIns');
  }

  /**
   * 获取使用报告（Office 365 活动用户详情）
   * 
   * 权限要求：Reports.Read.All
   * 国家云支持：✓ 全球版 ⚠️ 中国版（功能受限）✓ 美国政府版（GCC）
   * 
   * @param period - 时间段（D7, D30, D90, D180）
   * @returns CSV 格式的报告数据
   * @see https://learn.microsoft.com/graph/api/reportroot-getoffice365activeuserdetail
   */
  async getOffice365ActiveUserDetail(period: 'D7' | 'D30' | 'D90' | 'D180'): Promise<string> {
    try {
      return await this.client
        .api(`/reports/getOffice365ActiveUserDetail(period='${period}')`)
        .get();
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 获取邮件活动用户详情
   * 
   * @param period - 时间段
   * @returns CSV 格式的报告数据
   */
  async getEmailActivityUserDetail(period: 'D7' | 'D30' | 'D90' | 'D180'): Promise<string> {
    try {
      return await this.client
        .api(`/reports/getEmailActivityUserDetail(period='${period}')`)
        .get();
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 获取 OneDrive 使用情况账户详情
   * 
   * @param period - 时间段
   * @returns CSV 格式的报告数据
   */
  async getOneDriveUsageAccountDetail(period: 'D7' | 'D30' | 'D90' | 'D180'): Promise<string> {
    try {
      return await this.client
        .api(`/reports/getOneDriveUsageAccountDetail(period='${period}')`)
        .get();
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 获取 SharePoint 活动用户详情
   * 
   * @param period - 时间段
   * @returns CSV 格式的报告数据
   */
  async getSharePointActivityUserDetail(period: 'D7' | 'D30' | 'D90' | 'D180'): Promise<string> {
    try {
      return await this.client
        .api(`/reports/getSharePointActivityUserDetail(period='${period}')`)
        .get();
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 获取 Teams 用户活动用户详情
   * 
   * @param period - 时间段
   * @returns CSV 格式的报告数据
   */
  async getTeamsUserActivityUserDetail(period: 'D7' | 'D30' | 'D90' | 'D180'): Promise<string> {
    try {
      return await this.client
        .api(`/reports/getTeamsUserActivityUserDetail(period='${period}')`)
        .get();
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 用户洞察 - 趋势资源
   * 
   * 权限要求：Sites.Read.All, Sites.ReadWrite.All
   * 国家云支持：✓ 全球版 ⚠️ 中国版（不可用）✓ 美国政府版（部分）
   * 
   * @param userId - 用户 ID
   * @see https://learn.microsoft.com/graph/api/resources/insights-trending
   */
  trendingInsights(userId: string) {
    return new GraphEntityManager<Trending>(
      this.client,
      `/users/${encodeURIComponent(userId)}/insights/trending`
    );
  }

  /**
   * 用户洞察 - 使用过的资源
   * 
   * @param userId - 用户 ID
   * @see https://learn.microsoft.com/graph/api/resources/insights-used
   */
  usedInsights(userId: string) {
    return new GraphEntityManager<UsedInsight>(
      this.client,
      `/users/${encodeURIComponent(userId)}/insights/used`
    );
  }

  /**
   * 用户洞察 - 共享资源
   * 
   * @param userId - 用户 ID
   * @see https://learn.microsoft.com/graph/api/resources/insights-shared
   */
  sharedInsights(userId: string) {
    return new GraphEntityManager<SharedInsight>(
      this.client,
      `/users/${encodeURIComponent(userId)}/insights/shared`
    );
  }

  /**
   * 人员分析
   * 
   * 权限要求：People.Read, People.Read.All
   * 国家云支持：✓ 全球版 ⚠️ 中国版（功能受限）✓ 美国政府版
   * 
   * @param userId - 用户 ID
   * @see https://learn.microsoft.com/graph/api/resources/person
   */
  people(userId: string) {
    return new GraphEntityManager<Person>(
      this.client,
      `/users/${encodeURIComponent(userId)}/people`
    );
  }

  // ==================== 安全与合规 ====================

  /**
   * 安全警报管理
   * 
   * 权限要求：SecurityEvents.Read.All, SecurityEvents.ReadWrite.All
   * 国家云支持：✓ 全球版 ⚠️ 中国版（不可用）✓ 美国政府版
   * 
   * @see https://learn.microsoft.com/graph/api/resources/alert
   */
  get securityAlerts() {
    return new GraphEntityManager<Alert>(this.client, '/security/alerts');
  }

  /**
   * 更新安全警报
   * 
   * @param alertId - 警报 ID
   * @param updates - 更新内容
   * @returns 更新后的警报
   */
  async updateSecurityAlert(alertId: string, updates: Partial<Alert>): Promise<Alert> {
    try {
      return await this.client
        .api(`/security/alerts/${encodeURIComponent(alertId)}`)
        .patch(updates);
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 安全评分管理
   * 
   * 权限要求：SecurityEvents.Read.All
   * 国家云支持：✓ 全球版 ⚠️ 中国版（不可用）✓ 美国政府版
   * 
   * @see https://learn.microsoft.com/graph/api/resources/securescore
   */
  get secureScores() {
    return new GraphEntityManager<SecureScore>(this.client, '/security/secureScores');
  }

  /**
   * 风险检测管理
   * 
   * 权限要求：IdentityRiskEvent.Read.All
   * 国家云支持：✓ 全球版 ⚠️ 中国版（不可用）✓ 美国政府版（需要 P2 许可）
   * 
   * @see https://learn.microsoft.com/graph/api/resources/riskdetection
   */
  get riskDetections() {
    return new GraphEntityManager<RiskDetection>(this.client, '/identityProtection/riskDetections');
  }

  /**
   * 风险用户管理
   * 
   * 权限要求：IdentityRiskyUser.Read.All, IdentityRiskyUser.ReadWrite.All
   * 国家云支持：✓ 全球版 ⚠️ 中国版（不可用）✓ 美国政府版（需要 P2 许可）
   * 
   * @see https://learn.microsoft.com/graph/api/resources/riskyuser
   */
  get riskyUsers() {
    return new GraphEntityManager<RiskyUser>(this.client, '/identityProtection/riskyUsers');
  }

  /**
   * 确认风险用户已被入侵
   * 
   * 权限要求：IdentityRiskyUser.ReadWrite.All
   * 
   * @param userIds - 用户 ID 列表
   * @see https://learn.microsoft.com/graph/api/riskyusers-confirmcompromised
   */
  async confirmUsersCompromised(userIds: string[]): Promise<void> {
    try {
      await this.client
        .api('/identityProtection/riskyUsers/confirmCompromised')
        .post({ userIds });
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 解除风险用户
   * 
   * 权限要求：IdentityRiskyUser.ReadWrite.All
   * 
   * @param userIds - 用户 ID 列表
   * @see https://learn.microsoft.com/graph/api/riskyusers-dismiss
   */
  async dismissRiskyUsers(userIds: string[]): Promise<void> {
    try {
      await this.client
        .api('/identityProtection/riskyUsers/dismiss')
        .post({ userIds });
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  // ==================== 扩展 ====================

  /**
   * 架构扩展管理
   * 
   * 权限要求：Application.ReadWrite.All
   * 国家云支持：✓ 全球版 ✓ 中国版 ✓ 美国政府版
   * 
   * @see https://learn.microsoft.com/graph/api/resources/schemaextension
   */
  get schemaExtensions() {
    return new GraphEntityManager<SchemaExtension>(this.client, '/schemaExtensions');
  }

  /**
   * 创建架构扩展
   * 
   * @param extension - 扩展定义
   * @returns 创建的扩展
   * @see https://learn.microsoft.com/graph/api/schemaextension-post-schemaextensions
   */
  async createSchemaExtension(extension: Omit<SchemaExtension, 'id'>): Promise<SchemaExtension> {
    try {
      return await this.client
        .api('/schemaExtensions')
        .post(extension);
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 获取实体的开放扩展
   * 
   * @param resourceType - 资源类型（如 'users', 'groups'）
   * @param resourceId - 资源 ID
   * @returns 扩展列表
   * @see https://learn.microsoft.com/graph/api/opentypeextension-get
   */
  async getOpenExtensions(resourceType: string, resourceId: string): Promise<Extension[]> {
    try {
      const response = await this.client
        .api(`/${resourceType}/${encodeURIComponent(resourceId)}/extensions`)
        .get();
      return response.value || [];
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 创建开放扩展
   * 
   * @param resourceType - 资源类型
   * @param resourceId - 资源 ID
   * @param extension - 扩展数据
   * @returns 创建的扩展
   */
  async createOpenExtension(
    resourceType: string,
    resourceId: string,
    extension: Omit<Extension, 'id'>
  ): Promise<Extension> {
    try {
      return await this.client
        .api(`/${resourceType}/${encodeURIComponent(resourceId)}/extensions`)
        .post(extension);
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  // ==================== 目录对象 ====================

  /**
   * 通过 ID 获取目录对象
   * 
   * 权限要求：Directory.Read.All
   * 国家云支持：✓ 全球版 ✓ 中国版 ✓ 美国政府版
   * 
   * @param id - 对象 ID
   * @returns 目录对象
   * @see https://learn.microsoft.com/graph/api/directoryobject-get
   */
  async getDirectoryObject(id: string): Promise<DirectoryObject> {
    try {
      return await this.client
        .api(`/directoryObjects/${encodeURIComponent(id)}`)
        .get();
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 通过 ID 列表获取目录对象
   * 
   * @param ids - 对象 ID 列表（最多 1000 个）
   * @param types - 要返回的对象类型（可选）
   * @returns 目录对象列表
   * @see https://learn.microsoft.com/graph/api/directoryobject-getbyids
   */
  async getDirectoryObjectsByIds(ids: string[], types?: string[]): Promise<DirectoryObject[]> {
    try {
      const body: { ids: string[]; types?: string[] } = { ids };
      if (types && types.length > 0) {
        body.types = types;
      }
      const response = await this.client
        .api('/directoryObjects/getByIds')
        .post(body);
      return response.value || [];
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 检查成员组
   * 
   * @param memberId - 成员 ID
   * @param groupIds - 组 ID 列表（最多 20 个）
   * @returns 成员所属的组 ID 列表
   * @see https://learn.microsoft.com/graph/api/directoryobject-checkmembergroups
   */
  async checkMemberGroups(memberId: string, groupIds: string[]): Promise<string[]> {
    try {
      const response = await this.client
        .api(`/directoryObjects/${encodeURIComponent(memberId)}/checkMemberGroups`)
        .post({ groupIds });
      return response.value || [];
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 获取成员对象
   * 
   * @param directoryObjectId - 目录对象 ID
   * @param securityEnabledOnly - 是否只返回启用安全的组
   * @returns 成员对象列表
   * @see https://learn.microsoft.com/graph/api/directoryobject-getmemberobjects
   */
  async getMemberObjects(directoryObjectId: string, securityEnabledOnly: boolean = false): Promise<string[]> {
    try {
      const response = await this.client
        .api(`/directoryObjects/${encodeURIComponent(directoryObjectId)}/getMemberObjects`)
        .post({ securityEnabledOnly });
      return response.value || [];
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  // ==================== 辅助方法 ====================

  /**
   * 获取当前认证用户信息
   * 
   * 注意：此方法需要委托权限，不适用于纯应用权限场景
   * 
   * @returns 当前用户
   */
  async me(): Promise<User> {
    try {
      return await this.client.api('/me').get();
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 获取组织信息
   * 
   * 权限要求：Organization.Read.All
   */
  async organization(): Promise<unknown> {
    try {
      const response = await this.client.api('/organization').get();
      return response.value[0];
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }
}

// 便捷的 ORM 工厂函数
export function createGraphORM(client: Client): GraphORM {
  return new GraphORM(client);
}
