import { Client } from '@microsoft/microsoft-graph-client';
import { DriveItem, Permission, RequestDebugInfo, GraphCollection } from '../types';
import { GraphQueryBuilder } from '../query-builder';

/**
 * 文件服务
 * 
 * 提供文件相关的业务逻辑功能
 * 基本的文件 CRUD 请使用 DriveItemRepository
 * 
 * 权限要求：Files.Read, Files.ReadWrite, Files.Read.All, Files.ReadWrite.All
 * 国家云支持：✓ 全球版 ✓ 中国版 ✓ 美国政府版
 * 
 * @see https://learn.microsoft.com/graph/api/resources/driveitem
 */
export class FileService {
  private lastDebugInfo?: RequestDebugInfo;

  constructor(private client: Client) {}

  /**
   * 创建 query-builder 实例（私有辅助方法）
   */
  private query<T>(endpoint: string) {
    return new GraphQueryBuilder<T>(this.client, endpoint);
  }

  /**
   * 获取最后一次操作的调试信息
   */
  getDebug(): RequestDebugInfo | undefined {
    return this.lastDebugInfo;
  }

  /**
   * 上传小文件（< 4MB）
   * 
   * 对于大文件，请使用 uploadLargeFile
   * 使用 query-builder 支持调试和自定义 header
   */
  async uploadSmallFile(
    userId: string,
    itemPath: string,
    content: ArrayBuffer | Blob | string
  ): Promise<DriveItem> {
    return this.query('').put<DriveItem>(
      `users/${encodeURIComponent(userId)}/drive/root:${itemPath}:/content`,
      content
    );
  }

  /**
   * 上传大文件（创建上传会话）
   * 
   * 返回上传会话 URL，需要客户端自行实现分块上传
   * 使用 query-builder 支持调试和自定义 header
   */
  async createUploadSession(
    userId: string,
    itemPath: string,
    conflictBehavior: 'rename' | 'replace' | 'fail' = 'rename'
  ): Promise<{ uploadUrl: string }> {
    return this.query('').post<{ uploadUrl: string }>(
      'createUploadSession',
      {
        item: {
          '@microsoft.graph.conflictBehavior': conflictBehavior
        }
      },
      `users/${encodeURIComponent(userId)}/drive/root:${itemPath}`
    );
  }

  /**
   * 下载文件内容
   * 
   * 返回文件的下载 URL
   * 使用 query-builder 支持调试和自定义 header
   */
  async downloadFile(userId: string, itemId: string): Promise<string> {
    const item = await this.query<{ '@microsoft.graph.downloadUrl': string }>(`users/${encodeURIComponent(userId)}/drive/items/${encodeURIComponent(itemId)}`)
      .select(['@microsoft.graph.downloadUrl'])
      .first();
    return item ? item['@microsoft.graph.downloadUrl'] : '';
  }

  /**
   * 创建文件夹
   * 
   * 使用 query-builder 支持调试和自定义 header
   */
  async createFolder(
    userId: string,
    parentItemId: string,
    folderName: string
  ): Promise<DriveItem> {
    const parentPath = parentItemId === 'root'
      ? `users/${encodeURIComponent(userId)}/drive/root/children`
      : `users/${encodeURIComponent(userId)}/drive/items/${encodeURIComponent(parentItemId)}/children`;
    
    return this.query('').post<DriveItem>(
      'children',
      {
        name: folderName,
        folder: {},
        '@microsoft.graph.conflictBehavior': 'rename'
      },
      parentPath
    );
  }

  /**
   * 列出文件夹内容
   * 
   * 使用 query-builder 支持调试和自定义 header
   */
  async listChildren(userId: string, itemId: string = 'root'): Promise<GraphCollection<DriveItem>> {
    const path = itemId === 'root'
      ? `users/${encodeURIComponent(userId)}/drive/root/children`
      : `users/${encodeURIComponent(userId)}/drive/items/${encodeURIComponent(itemId)}/children`;

    return await this.query<DriveItem>(path).get();
  }

  /**
   * 搜索文件
   * 
   * 使用 query-builder 支持调试和自定义 header
   */
  async search(userId: string, query: string): Promise<GraphCollection<DriveItem>> {
    return await this.query<DriveItem>(
      `users/${encodeURIComponent(userId)}/drive/root/search(q='${encodeURIComponent(query)}')`
    ).get();
  }

  /**
   * 移动文件/文件夹
   * 
   * 使用 query-builder 支持调试和自定义 header
   */
  async move(
    userId: string,
    itemId: string,
    targetParentId: string,
    newName?: string
  ): Promise<DriveItem> {
    const body: { parentReference: { id: string }; name?: string } = {
      parentReference: { id: targetParentId }
    };
    if (newName) {
      body.name = newName;
    }
    return this.query<DriveItem>(
      `users/${encodeURIComponent(userId)}/drive/items}`
    ).update(itemId, body);
  }

  /**
   * 复制文件/文件夹
   * 
   * 使用 query-builder 支持调试和自定义 header
   */
  async copy(
    userId: string,
    itemId: string,
    targetParentId: string,
    newName?: string
  ): Promise<string> {
    const body: { parentReference: { id: string }; name?: string } = {
      parentReference: { id: targetParentId }
    };
    if (newName) {
      body.name = newName;
    }
    return this.query('').post<string>(
      'copy',
      body,
      `users/${encodeURIComponent(userId)}/drive/items/${encodeURIComponent(itemId)}`
    );
  }

  /**
   * 创建共享链接
   * 
   * 使用 query-builder 支持调试和自定义 header
   */
  async createShareLink(
    userId: string,
    itemId: string,
    type: 'view' | 'edit' | 'embed',
    scope: 'anonymous' | 'organization'
  ): Promise<Permission> {
    return this.query('').post<Permission>(
      'createLink',
      { type, scope },
      `users/${encodeURIComponent(userId)}/drive/items/${encodeURIComponent(itemId)}`
    );
  }

  /**
   * 邀请用户访问文件
   * 
   * 使用 query-builder 支持调试和自定义 header
   */
  async invite(
    userId: string,
    itemId: string,
    recipients: string[],
    roles: ('read' | 'write')[],
    requireSignIn: boolean = true,
    sendInvitation: boolean = true,
    message?: string
  ): Promise<Permission[]> {
    const response = await this.query('').post<{ value: Permission[] }>(
      'invite',
      {
        recipients: recipients.map(email => ({ email })),
        roles,
        requireSignIn,
        sendInvitation,
        message
      },
      `users/${encodeURIComponent(userId)}/drive/items/${encodeURIComponent(itemId)}`
    );
    return response.value || [];
  }

  /**
   * 列出文件权限
   * 
   * 使用 query-builder 支持调试和自定义 header
   */
  async listPermissions(userId: string, itemId: string): Promise<GraphCollection<Permission>> {
    return await this.query<Permission>(
      `users/${encodeURIComponent(userId)}/drive/items/${encodeURIComponent(itemId)}/permissions`
    ).get();
  }

  /**
   * 删除权限
   * 
   * 使用 query-builder 支持调试和自定义 header
   */
  async deletePermission(userId: string, itemId: string, permissionId: string): Promise<void> {
    return await this.query<Permission>(
      `users/${encodeURIComponent(userId)}/drive/items/${encodeURIComponent(itemId)}/permissions`
    ).delete(permissionId);
  }

  /**
   * 获取文件缩略图
   * 
   * 使用 query-builder 支持调试和自定义 header
   */
  async getThumbnails(userId: string, itemId: string): Promise<GraphCollection<unknown>> {
    return await this.query<unknown>(
      `users/${encodeURIComponent(userId)}/drive/items/${encodeURIComponent(itemId)}/thumbnails`
    ).get();
  }

  /**
   * 驱动器项增量查询
   * 
   * 跟踪文件和文件夹的变化（新增、修改、删除）
   * 使用 query-builder 支持调试和自定义 header
   */
  async delta(userId: string, deltaLink?: string): Promise<GraphCollection<DriveItem>> {
    const endpoint = deltaLink || `users/${encodeURIComponent(userId)}/drive/root/delta`;
    return await this.query<DriveItem>(endpoint).get();
  }
}
