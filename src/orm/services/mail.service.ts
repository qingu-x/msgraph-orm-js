import { Client } from '@microsoft/microsoft-graph-client';
import { Message, Attachment, FileAttachment, MailFolder, DeltaCollection } from '../types';
import { GraphOrmError } from '../errors';

/**
 * 邮件服务
 * 
 * 提供邮件发送、接收和管理功能
 * 
 * 权限要求：Mail.Read, Mail.ReadWrite, Mail.Send
 * 国家云支持：✓ 全球版 ✓ 中国版 ✓ 美国政府版
 * 
 * @see https://learn.microsoft.com/graph/api/resources/message
 */
export class MailService {
  constructor(private client: Client) {}

  /**
   * 获取用户所有邮件
   */
  async getMessages(userId: string, top?: number): Promise<Message[]> {
    try {
      let request = this.client
        .api(`/users/${encodeURIComponent(userId)}/messages`)
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
   * 获取收件箱邮件
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
        .orderby('createdDateTime DESC');
      
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
   * 获取指定文件夹的邮件
   */
  async getMessagesInFolder(userId: string, folderId: string, top?: number): Promise<Message[]> {
    try {
      let request = this.client
        .api(`/users/${encodeURIComponent(userId)}/mailFolders/${encodeURIComponent(folderId)}/messages`)
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
   * 获取单个邮件
   */
  async getMessage(userId: string, messageId: string): Promise<Message> {
    try {
      return await this.client
        .api(`/users/${encodeURIComponent(userId)}/messages/${encodeURIComponent(messageId)}`)
        .get();
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 发送邮件
   */
  async send(userId: string, message: Partial<Message>, saveToSentItems: boolean = true): Promise<void> {
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

  /**
   * 创建草稿
   */
  async createDraft(userId: string, message: Partial<Message>): Promise<Message> {
    try {
      return await this.client
        .api(`/users/${encodeURIComponent(userId)}/messages`)
        .post(message);
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 更新邮件（例如标记为已读）
   */
  async updateMessage(userId: string, messageId: string, updates: Partial<Message>): Promise<Message> {
    try {
      return await this.client
        .api(`/users/${encodeURIComponent(userId)}/messages/${encodeURIComponent(messageId)}`)
        .patch(updates);
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 删除邮件
   */
  async deleteMessage(userId: string, messageId: string): Promise<void> {
    try {
      await this.client
        .api(`/users/${encodeURIComponent(userId)}/messages/${encodeURIComponent(messageId)}`)
        .delete();
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 移动邮件到文件夹
   */
  async moveMessage(userId: string, messageId: string, destinationFolderId: string): Promise<Message> {
    try {
      return await this.client
        .api(`/users/${encodeURIComponent(userId)}/messages/${encodeURIComponent(messageId)}/move`)
        .post({
          destinationId: destinationFolderId
        });
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 复制邮件到文件夹
   */
  async copyMessage(userId: string, messageId: string, destinationFolderId: string): Promise<Message> {
    try {
      return await this.client
        .api(`/users/${encodeURIComponent(userId)}/messages/${encodeURIComponent(messageId)}/copy`)
        .post({
          destinationId: destinationFolderId
        });
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 回复邮件
   */
  async reply(userId: string, messageId: string, comment?: string): Promise<void> {
    try {
      await this.client
        .api(`/users/${encodeURIComponent(userId)}/messages/${encodeURIComponent(messageId)}/reply`)
        .post({
          comment: comment || ''
        });
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 全部回复
   */
  async replyAll(userId: string, messageId: string, comment?: string): Promise<void> {
    try {
      await this.client
        .api(`/users/${encodeURIComponent(userId)}/messages/${encodeURIComponent(messageId)}/replyAll`)
        .post({
          comment: comment || ''
        });
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 转发邮件
   */
  async forward(
    userId: string,
    messageId: string,
    toRecipients: Array<{ emailAddress: { address: string; name?: string } }>,
    comment?: string
  ): Promise<void> {
    try {
      await this.client
        .api(`/users/${encodeURIComponent(userId)}/messages/${encodeURIComponent(messageId)}/forward`)
        .post({
          toRecipients,
          comment: comment || ''
        });
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 获取邮件附件
   */
  async getAttachments(userId: string, messageId: string): Promise<Attachment[]> {
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
   * 添加附件
   */
  async addAttachment(
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
   * 删除附件
   */
  async deleteAttachment(userId: string, messageId: string, attachmentId: string): Promise<void> {
    try {
      await this.client
        .api(`/users/${encodeURIComponent(userId)}/messages/${encodeURIComponent(messageId)}/attachments/${encodeURIComponent(attachmentId)}`)
        .delete();
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
  async createMailFolder(userId: string, displayName: string, parentFolderId?: string): Promise<MailFolder> {
    try {
      const endpoint = parentFolderId
        ? `/users/${encodeURIComponent(userId)}/mailFolders/${encodeURIComponent(parentFolderId)}/childFolders`
        : `/users/${encodeURIComponent(userId)}/mailFolders`;

      return await this.client
        .api(endpoint)
        .post({ displayName });
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
   * 搜索邮件
   */
  async search(userId: string, searchQuery: string, top?: number): Promise<Message[]> {
    try {
      let request = this.client
        .api(`/users/${encodeURIComponent(userId)}/messages`)
        .filter(`contains(subject,'${searchQuery}') or contains(from/emailAddress/name,'${searchQuery}')`)
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
   * 邮件增量查询
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

  /**
   * 标记为已读
   */
  async markAsRead(userId: string, messageId: string): Promise<Message> {
    return this.updateMessage(userId, messageId, { isRead: true });
  }

  /**
   * 标记为未读
   */
  async markAsUnread(userId: string, messageId: string): Promise<Message> {
    return this.updateMessage(userId, messageId, { isRead: false });
  }

  /**
   * 标记为重要
   */
  async markAsImportant(userId: string, messageId: string): Promise<Message> {
    return this.updateMessage(userId, messageId, { importance: 'high' });
  }
}

