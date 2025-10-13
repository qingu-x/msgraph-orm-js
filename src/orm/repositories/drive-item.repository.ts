import { Client } from '@microsoft/microsoft-graph-client';
import { GraphRepository } from '../repository';
import { DriveItem } from '../types';
import { GraphOrmError } from '../errors';

/**
 * 驱动器项仓储
 * 
 * 权限要求：Files.Read, Files.ReadWrite, Files.Read.All, Files.ReadWrite.All
 * 国家云支持：✓ 全球版 ✓ 中国版 ✓ 美国政府版
 * 
 * @see https://learn.microsoft.com/graph/api/resources/driveitem
 */
export class DriveItemRepository extends GraphRepository<DriveItem> {
  constructor(client: Client, endpoint: string) {
    super(client, endpoint);
  }

  /**
   * 通过路径获取项
   */
  async getByPath(itemPath: string): Promise<DriveItem> {
    try {
      // 移除 endpoint 中的 /children 部分（如果存在）
      const baseEndpoint = this.endpoint.replace(/\/children$/, '');
      return await this.client
        .api(`${baseEndpoint}:${itemPath}`)
        .get();
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 列出子项
   */
  async listChildren(itemId: string): Promise<DriveItem[]> {
    try {
      const endpoint = itemId === 'root'
        ? `${this.endpoint}/root/children`
        : `${this.endpoint}/${encodeURIComponent(itemId)}/children`;
      
      const response = await this.client.api(endpoint).get();
      return response.value || [];
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 移动项
   */
  async move(
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
        .api(`${this.endpoint}/${encodeURIComponent(itemId)}`)
        .patch(body);
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 复制项
   */
  async copy(
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
        .api(`${this.endpoint}/${encodeURIComponent(itemId)}/copy`)
        .post(body);
      return response;
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 获取下载 URL
   */
  async getDownloadUrl(itemId: string): Promise<string> {
    try {
      const item = await this.client
        .api(`${this.endpoint}/${encodeURIComponent(itemId)}`)
        .select('@microsoft.graph.downloadUrl')
        .get();
      return item['@microsoft.graph.downloadUrl'];
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 搜索项
   */
  async search(query: string): Promise<DriveItem[]> {
    try {
      // 从 endpoint 提取 drive 路径
      const baseEndpoint = this.endpoint.replace(/\/items.*$/, '');
      const response = await this.client
        .api(`${baseEndpoint}/root/search(q='${encodeURIComponent(query)}')`)
        .get();
      return response.value || [];
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 获取缩略图
   */
  async getThumbnails(itemId: string): Promise<unknown[]> {
    try {
      const response = await this.client
        .api(`${this.endpoint}/${encodeURIComponent(itemId)}/thumbnails`)
        .get();
      return response.value || [];
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 列出权限
   */
  async listPermissions(itemId: string): Promise<unknown[]> {
    try {
      const response = await this.client
        .api(`${this.endpoint}/${encodeURIComponent(itemId)}/permissions`)
        .get();
      return response.value || [];
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 删除权限
   */
  async deletePermission(itemId: string, permissionId: string): Promise<void> {
    try {
      await this.client
        .api(`${this.endpoint}/${encodeURIComponent(itemId)}/permissions/${encodeURIComponent(permissionId)}`)
        .delete();
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }
}

