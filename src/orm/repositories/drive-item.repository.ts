import { Client } from '@microsoft/microsoft-graph-client';
import { GraphRepository } from '../repository';
import { DriveItem, GraphCollection } from '../types';

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
   * 
   * 使用 query-builder 支持调试和自定义 header
   */
  async getByPath(itemPath: string): Promise<DriveItem | null> {
    // 移除 endpoint 中的 /children 部分（如果存在）
    const baseEndpoint = this.endpoint.replace(/\/children$/, '');
    const pathSegment = baseEndpoint.split('/').pop(); // 获取最后一段
    return this.query(`../${pathSegment}:${itemPath}`).first();
  }

  /**
   * 列出子项
   * 
   * 使用 query-builder 支持调试和自定义 header
   */
  async listChildren(itemId: string): Promise<GraphCollection<DriveItem>> {
    const path = itemId === 'root'
      ? 'root/children'
      : `${encodeURIComponent(itemId)}/children`;
    
    return await this.query(path).get() as GraphCollection<DriveItem>;
  }

  /**
   * 移动项
   * 
   * 使用 update 方法（基于 query-builder）支持调试和自定义 header
   */
  async move(
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
    return this.update(itemId, body as Partial<DriveItem>);
  }

  /**
   * 复制项
   * 
   * 使用 query-builder 支持调试和自定义 header
   */
  async copy(
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
    return this.query().post(`${encodeURIComponent(itemId)}/copy`, body);
  }

  /**
   * 获取下载 URL
   * 
   * 使用 query-builder 支持调试和自定义 header
   */
  async getDownloadUrl(itemId: string): Promise<string> {
    const item = await this.query()
      .select(['@microsoft.graph.downloadUrl'])
      .findById(itemId);
    return (item as unknown as { '@microsoft.graph.downloadUrl': string })['@microsoft.graph.downloadUrl'];
  }

  /**
   * 搜索项
   * 
   * 使用 query-builder 支持调试和自定义 header
   */
  async search(query: string): Promise<GraphCollection<DriveItem>> {
    const searchPath = `../root/search(q='${encodeURIComponent(query)}')`;
    return await this.query(searchPath).get() as GraphCollection<DriveItem>;
  }

  /**
   * 获取缩略图
   * 
   * 使用 query-builder 支持调试和自定义 header
   */
  async getThumbnails(itemId: string): Promise<GraphCollection<unknown>> {
    return await this.query(`${encodeURIComponent(itemId)}/thumbnails`).get() as GraphCollection<unknown>;
  }

  /**
   * 列出权限
   * 
   * 使用 query-builder 支持调试和自定义 header
   */
  async listPermissions(itemId: string): Promise<GraphCollection<unknown>> {
    return await this.query(`${encodeURIComponent(itemId)}/permissions`).get() as GraphCollection<unknown>;
  }

  /**
   * 删除权限
   * 
   * 使用 query-builder 支持调试和自定义 header
   */
  async deletePermission(itemId: string, permissionId: string): Promise<void> {
    return await this.query(`${encodeURIComponent(itemId)}/permissions`).delete(permissionId);
  }
}

