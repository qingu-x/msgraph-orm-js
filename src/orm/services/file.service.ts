import { Client } from '@microsoft/microsoft-graph-client';
import { DriveItem, Permission, DeltaCollection } from '../types';
import { GraphOrmError } from '../errors';

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
  constructor(private client: Client) {}

  /**
   * 上传小文件（< 4MB）
   * 
   * 对于大文件，请使用 uploadLargeFile
   */
  async uploadSmallFile(
    userId: string,
    itemPath: string,
    content: ArrayBuffer | Blob | string
  ): Promise<DriveItem> {
    try {
      return await this.client
        .api(`/users/${encodeURIComponent(userId)}/drive/root:${itemPath}:/content`)
        .put(content);
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 上传大文件（创建上传会话）
   * 
   * 返回上传会话 URL，需要客户端自行实现分块上传
   */
  async createUploadSession(
    userId: string,
    itemPath: string,
    conflictBehavior: 'rename' | 'replace' | 'fail' = 'rename'
  ): Promise<{ uploadUrl: string }> {
    try {
      return await this.client
        .api(`/users/${encodeURIComponent(userId)}/drive/root:${itemPath}:/createUploadSession`)
        .post({
          item: {
            '@microsoft.graph.conflictBehavior': conflictBehavior
          }
        });
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 下载文件内容
   * 
   * 返回文件的下载 URL
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
   * 创建文件夹
   */
  async createFolder(
    userId: string,
    parentItemId: string,
    folderName: string
  ): Promise<DriveItem> {
    try {
      const parentPath = parentItemId === 'root'
        ? `/users/${encodeURIComponent(userId)}/drive/root/children`
        : `/users/${encodeURIComponent(userId)}/drive/items/${encodeURIComponent(parentItemId)}/children`;

      return await this.client
        .api(parentPath)
        .post({
          name: folderName,
          folder: {},
          '@microsoft.graph.conflictBehavior': 'rename'
        });
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 列出文件夹内容
   */
  async listChildren(userId: string, itemId: string = 'root'): Promise<DriveItem[]> {
    try {
      const path = itemId === 'root'
        ? `/users/${encodeURIComponent(userId)}/drive/root/children`
        : `/users/${encodeURIComponent(userId)}/drive/items/${encodeURIComponent(itemId)}/children`;

      const response = await this.client.api(path).get();
      return response.value || [];
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 搜索文件
   */
  async search(userId: string, query: string): Promise<DriveItem[]> {
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
   * 移动文件/文件夹
   */
  async move(
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
   * 复制文件/文件夹
   */
  async copy(
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
      return response;
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 创建共享链接
   */
  async createShareLink(
    userId: string,
    itemId: string,
    type: 'view' | 'edit' | 'embed',
    scope: 'anonymous' | 'organization'
  ): Promise<Permission> {
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
   * 邀请用户访问文件
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
    try {
      const response = await this.client
        .api(`/users/${encodeURIComponent(userId)}/drive/items/${encodeURIComponent(itemId)}/invite`)
        .post({
          recipients: recipients.map(email => ({ email })),
          roles,
          requireSignIn,
          sendInvitation,
          message
        });
      return response.value || [];
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 列出文件权限
   */
  async listPermissions(userId: string, itemId: string): Promise<Permission[]> {
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
   * 删除权限
   */
  async deletePermission(userId: string, itemId: string, permissionId: string): Promise<void> {
    try {
      await this.client
        .api(`/users/${encodeURIComponent(userId)}/drive/items/${encodeURIComponent(itemId)}/permissions/${encodeURIComponent(permissionId)}`)
        .delete();
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 获取文件缩略图
   */
  async getThumbnails(userId: string, itemId: string): Promise<unknown[]> {
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
   * 驱动器项增量查询
   * 
   * 跟踪文件和文件夹的变化（新增、修改、删除）
   */
  async delta(userId: string, deltaLink?: string): Promise<DeltaCollection<DriveItem>> {
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
}
