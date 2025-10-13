import { Client } from '@microsoft/microsoft-graph-client';
import { DriveItem, DeltaCollection } from '../types';
import { GraphOrmError } from '../errors';

/**
 * 文件服务
 * 
 * 提供 OneDrive 和 SharePoint 文件操作
 * 
 * 权限要求：Files.Read, Files.ReadWrite, Files.Read.All, Files.ReadWrite.All
 * 国家云支持：✓ 全球版 ✓ 中国版 ✓ 美国政府版
 * 
 * @see https://learn.microsoft.com/graph/api/resources/driveitem
 */
export class FileService {
  constructor(private client: Client) {}

  /**
   * 获取用户驱动器根目录的项
   */
  getUserRootItems(userId: string) {
    return this.listChildren(userId, 'root');
  }

  /**
   * 获取驱动器项（通过 ID）
   */
  async getItem(userId: string, itemId: string): Promise<DriveItem> {
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
   */
  async getItemByPath(userId: string, itemPath: string): Promise<DriveItem> {
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
   */
  async listChildren(userId: string, itemId: string): Promise<DriveItem[]> {
    try {
      const endpoint = itemId === 'root'
        ? `/users/${encodeURIComponent(userId)}/drive/root/children`
        : `/users/${encodeURIComponent(userId)}/drive/items/${encodeURIComponent(itemId)}/children`;
      
      const response = await this.client.api(endpoint).get();
      return response.value || [];
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 创建文件夹
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
   * 下载文件内容（获取下载 URL）
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
   */
  async downloadAsFormat(userId: string, itemId: string, format: string): Promise<string> {
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
   */
  async updateMetadata(userId: string, itemId: string, updates: Partial<DriveItem>): Promise<DriveItem> {
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
   */
  async delete(userId: string, itemId: string): Promise<void> {
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
   */
  async permanentDelete(userId: string, itemId: string): Promise<void> {
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
   * 复制驱动器项
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
   * 搜索驱动器项
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
   * 驱动器项增量查询（Delta Query）
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

  /**
   * 创建共享链接
   */
  async createShareLink(
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
   */
  async invite(
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
   */
  async listPermissions(userId: string, itemId: string): Promise<unknown[]> {
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
   * 获取驱动器项缩略图
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
   * 获取与我共享的项目
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
}

