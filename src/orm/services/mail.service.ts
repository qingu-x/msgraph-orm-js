import { Client } from '@microsoft/microsoft-graph-client';
import { Message, MailFolder, DeltaCollection } from '../types';
import { GraphOrmError } from '../errors';

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
  constructor(private client: Client) {}

  /**
   * 发送邮件
   * 
   * 权限要求：Mail.Send
   */
  async send(userId: string, message: Omit<Message, 'id'>): Promise<void> {
    try {
      await this.client
        .api(`/users/${encodeURIComponent(userId)}/sendMail`)
        .post({
          message,
          saveToSentItems: true
        });
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 发送草稿邮件
   */
  async sendDraft(userId: string, messageId: string): Promise<void> {
    try {
      await this.client
        .api(`/users/${encodeURIComponent(userId)}/messages/${encodeURIComponent(messageId)}/send`)
        .post({});
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 创建草稿
   */
  async createDraft(userId: string, message: Omit<Message, 'id'>): Promise<Message> {
    try {
      return await this.client
        .api(`/users/${encodeURIComponent(userId)}/messages`)
        .post(message);
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 获取收件箱邮件
   * 
   * 快捷方法，等同于获取 inbox 文件夹的邮件
   */
  async getInbox(userId: string, top?: number): Promise<Message[]> {
    try {
      let request = this.client
        .api(`/users/${encodeURIComponent(userId)}/mailFolders/inbox/messages`)
        .orderby('receivedDateTime DESC');
      
      if (top) {
        request = request.top(top);
      }
      
      const response = await request.get();
      return response.value || [];
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 获取已发送邮件
   */
  async getSentItems(userId: string, top?: number): Promise<Message[]> {
    try {
      let request = this.client
        .api(`/users/${encodeURIComponent(userId)}/mailFolders/sentitems/messages`)
        .orderby('sentDateTime DESC');
      
      if (top) {
        request = request.top(top);
      }
      
      const response = await request.get();
      return response.value || [];
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 获取草稿箱邮件
   */
  async getDrafts(userId: string, top?: number): Promise<Message[]> {
    try {
      let request = this.client
        .api(`/users/${encodeURIComponent(userId)}/mailFolders/drafts/messages`)
        .orderby('lastModifiedDateTime DESC');
      
      if (top) {
        request = request.top(top);
      }
      
      const response = await request.get();
      return response.value || [];
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 获取邮件文件夹
   */
  async getMailFolders(userId: string): Promise<MailFolder[]> {
    try {
      const response = await this.client
        .api(`/users/${encodeURIComponent(userId)}/mailFolders`)
        .get();
      return response.value || [];
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 创建邮件文件夹
   */
  async createMailFolder(
    userId: string,
    displayName: string,
    isHidden: boolean = false
  ): Promise<MailFolder> {
    try {
      return await this.client
        .api(`/users/${encodeURIComponent(userId)}/mailFolders`)
        .post({
          displayName,
          isHidden
        });
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 获取子文件夹
   */
  async getChildFolders(userId: string, parentFolderId: string): Promise<MailFolder[]> {
    try {
      const response = await this.client
        .api(`/users/${encodeURIComponent(userId)}/mailFolders/${encodeURIComponent(parentFolderId)}/childFolders`)
        .get();
      return response.value || [];
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 删除邮件文件夹
   */
  async deleteMailFolder(userId: string, folderId: string): Promise<void> {
    try {
      await this.client
        .api(`/users/${encodeURIComponent(userId)}/mailFolders/${encodeURIComponent(folderId)}`)
        .delete();
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 邮件增量查询
   * 
   * 跟踪邮件的变化（新增、修改、删除）
   */
  async delta(userId: string, deltaLink?: string): Promise<DeltaCollection<Message>> {
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
}
