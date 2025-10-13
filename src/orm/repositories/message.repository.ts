import { Client } from '@microsoft/microsoft-graph-client';
import { GraphRepository } from '../repository';
import { Message, Attachment, FileAttachment } from '../types';
import { GraphOrmError } from '../errors';

/**
 * 邮件仓储
 * 
 * 权限要求：Mail.Read, Mail.ReadWrite
 * 国家云支持：✓ 全球版 ✓ 中国版 ✓ 美国政府版
 * 
 * @see https://learn.microsoft.com/graph/api/resources/message
 */
export class MessageRepository extends GraphRepository<Message> {
  constructor(client: Client, endpoint: string) {
    super(client, endpoint);
  }

  /**
   * 回复邮件
   */
  async reply(messageId: string, comment?: string): Promise<void> {
    try {
      await this.client
        .api(`${this.endpoint}/${encodeURIComponent(messageId)}/reply`)
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
  async replyAll(messageId: string, comment?: string): Promise<void> {
    try {
      await this.client
        .api(`${this.endpoint}/${encodeURIComponent(messageId)}/replyAll`)
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
    messageId: string,
    toRecipients: Array<{ emailAddress: { address: string; name?: string } }>,
    comment?: string
  ): Promise<void> {
    try {
      await this.client
        .api(`${this.endpoint}/${encodeURIComponent(messageId)}/forward`)
        .post({
          toRecipients,
          comment: comment || ''
        });
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 移动邮件到文件夹
   */
  async move(messageId: string, destinationFolderId: string): Promise<Message> {
    try {
      return await this.client
        .api(`${this.endpoint}/${encodeURIComponent(messageId)}/move`)
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
  async copy(messageId: string, destinationFolderId: string): Promise<Message> {
    try {
      return await this.client
        .api(`${this.endpoint}/${encodeURIComponent(messageId)}/copy`)
        .post({
          destinationId: destinationFolderId
        });
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 标记为已读
   */
  async markAsRead(messageId: string): Promise<Message> {
    return this.update(messageId, { isRead: true });
  }

  /**
   * 标记为未读
   */
  async markAsUnread(messageId: string): Promise<Message> {
    return this.update(messageId, { isRead: false });
  }

  /**
   * 标记为重要
   */
  async markAsImportant(messageId: string): Promise<Message> {
    return this.update(messageId, { importance: 'high' });
  }

  /**
   * 获取邮件附件
   */
  async getAttachments(messageId: string): Promise<Attachment[]> {
    try {
      const response = await this.client
        .api(`${this.endpoint}/${encodeURIComponent(messageId)}/attachments`)
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
    messageId: string,
    attachment: Omit<FileAttachment, 'id'>
  ): Promise<Attachment> {
    try {
      return await this.client
        .api(`${this.endpoint}/${encodeURIComponent(messageId)}/attachments`)
        .post(attachment);
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 删除附件
   */
  async deleteAttachment(messageId: string, attachmentId: string): Promise<void> {
    try {
      await this.client
        .api(`${this.endpoint}/${encodeURIComponent(messageId)}/attachments/${encodeURIComponent(attachmentId)}`)
        .delete();
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }
}

