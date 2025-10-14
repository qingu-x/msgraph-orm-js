import { Client } from '@microsoft/microsoft-graph-client';
import { Message, MailFolder, RequestDebugInfo, GraphCollection } from '../types';
import { GraphQueryBuilder } from '../query-builder';

/**
 * 邮件服务
 * 
 * 提供邮件相关的业务逻辑功能
 * 基本的邮件 CRUD 请使用 MessageRepository
 * 
 * 权限要求：Mail.Read, Mail.ReadWrite, Mail.Send
 * 国家云支持：✓ 全球版 ✓ 中国版 ✓ 美国政府版
 * 
 * @see https://learn.microsoft.com/graph/api/resources/message
 */
export class MailService {
  private lastDebugInfo?: RequestDebugInfo;

  constructor(private client: Client) {}

  /**
   * 创建 query-builder 实例（私有辅助方法）
   */
  private query<T>(endpoint: string) {
    return new GraphQueryBuilder<T>(this.client, endpoint);
  }

  /**
   * 发送邮件
   * 
   * 权限要求：Mail.Send
   * 使用 query-builder 支持调试和自定义 header
   */
  async send(userId: string, message: Omit<Message, 'id'>): Promise<void> {
    return this.query('').post<void>(
      'sendMail',
      { message, saveToSentItems: true },
      `users/${encodeURIComponent(userId)}`
    );
  }

  /**
   * 发送草稿邮件
   * 
   * 使用 query-builder 支持调试和自定义 header
   */
  async sendDraft(userId: string, messageId: string): Promise<void> {
    return this.query('').post<void>(
      'send',
      {},
      `users/${encodeURIComponent(userId)}/messages/${encodeURIComponent(messageId)}`
    );
  }

  /**
   * 创建草稿
   * 
   * 使用 query-builder 支持调试和自定义 header
   */
  async createDraft(userId: string, message: Omit<Message, 'id'>): Promise<Message> {
    return this.query('').post<Message>(
      'messages',
      message,
      `users/${encodeURIComponent(userId)}`
    );
  }

  /**
   * 获取收件箱邮件
   * 
   * 快捷方法，等同于获取 inbox 文件夹的邮件
   * 使用 query-builder 支持调试和自定义 header
   */
  async getInbox(userId: string, top?: number): Promise<GraphCollection<Message>> {
    const request = this.query<Message>(`users/${encodeURIComponent(userId)}/mailFolders/inbox/messages`)
      .orderBy('receivedDateTime', 'desc');

    if (top) {
      request.top(top);
    }
    return await request.get();
  }

  /**
   * 获取已发送邮件
   * 
   * 使用 query-builder 支持调试和自定义 header
   */
  async getSentItems(userId: string, top?: number): Promise<GraphCollection<Message>> {
    const request = this.query<Message>(`users/${encodeURIComponent(userId)}/mailFolders/sentitems/messages`)
      .orderBy('sentDateTime', 'desc');

    if (top) {
      request.top(top);
    }
    return await request.get();
  }

  /**
   * 获取草稿箱邮件
   * 
   * 使用 query-builder 支持调试和自定义 header
   */
  async getDrafts(userId: string, top?: number): Promise<GraphCollection<Message>> {
    const request = this.query<Message>(`users/${encodeURIComponent(userId)}/mailFolders/drafts/messages`)
      .orderBy('lastModifiedDateTime', 'desc');

    if (top) {
      request.top(top);
    }
    return await request.get();
  }

  /**
   * 获取邮件文件夹
   * 
   * 使用 query-builder 支持调试和自定义 header
   */
  async getMailFolders(userId: string): Promise<GraphCollection<MailFolder>> {
    const request = this.query<MailFolder>(`users/${encodeURIComponent(userId)}/mailFolders`)
      .orderBy('displayName', 'asc');
    return await request.get();
  }

  /**
   * 创建邮件文件夹
   * 
   * 使用 query-builder 支持调试和自定义 header
   */
  async createMailFolder(
    userId: string,
    displayName: string,
    isHidden: boolean = false
  ): Promise<MailFolder> {
    return this.query('').post<MailFolder>(
      'mailFolders',
      { displayName, isHidden },
      `users/${encodeURIComponent(userId)}`
    );
  }

  /**
   * 获取子文件夹
   * 
   * 使用 query-builder 支持调试和自定义 header
   */
  async getChildFolders(userId: string, parentFolderId: string): Promise<GraphCollection<MailFolder>> {
    const request = this.query<MailFolder>(`users/${encodeURIComponent(userId)}/mailFolders/${encodeURIComponent(parentFolderId)}/childFolders`)
      .orderBy('displayName', 'asc');
    return await request.get();
  }

  /**
   * 删除邮件文件夹
   * 
   * 使用 query-builder 支持调试和自定义 header
   */
  async deleteMailFolder(userId: string, folderId: string): Promise<void> {
    return this.query(`users/${encodeURIComponent(userId)}/mailFolders`).delete(folderId);
  }

  /**
   * 邮件增量查询
   * 
   * 跟踪邮件的变化（新增、修改、删除）
   * 使用 query-builder 支持调试和自定义 header
   */
  async delta(userId: string, deltaLink?: string): Promise<GraphCollection<Message>> {
    const endpoint = deltaLink || `users/${encodeURIComponent(userId)}/messages/delta`;
    return await this.query<Message>(endpoint).get();
  }

  /**
   * 邮件增量查询（旧版）
   * 
   * 跟踪邮件的变化（新增、修改、删除）
   * 使用 query-builder 支持调试和自定义 header
   */
  async deltaLegacy(userId: string, deltaLink?: string): Promise<GraphCollection<Message>> {
    const endpoint = deltaLink || `users/${encodeURIComponent(userId)}/messages/delta`;
    return await this.query<Message>(endpoint).get();
  }
}
