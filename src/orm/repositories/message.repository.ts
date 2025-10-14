import { Client } from '@microsoft/microsoft-graph-client';
import { GraphRepository } from '../repository';
import { Message, Attachment, FileAttachment, GraphCollection } from '../types';

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
   * 
   * 使用 query-builder 支持调试和自定义 header
   */
  async reply(messageId: string, comment?: string): Promise<void> {
    return this.query().post('reply', {
      comment: comment || ''
    }, messageId);
  }

  /**
   * 全部回复
   * 
   * 使用 query-builder 支持调试和自定义 header
   */
  async replyAll(messageId: string, comment?: string): Promise<void> {
    return this.query().post('replyAll', {
      comment: comment || ''
    }, messageId);
  }

  /**
   * 转发邮件
   * 
   * 使用 query-builder 支持调试和自定义 header
   */
  async forward(
    messageId: string,
    toRecipients: Array<{ emailAddress: { address: string; name?: string } }>,
    comment?: string
  ): Promise<void> {
    return this.query().post('forward', {
      toRecipients,
      comment: comment || ''
    }, messageId);
  }

  /**
   * 移动邮件到文件夹
   * 
   * 使用 query-builder 支持调试和自定义 header
   */
  async move(messageId: string, destinationFolderId: string): Promise<Message> {
    return this.query().post<Message>('move', {
      destinationId: destinationFolderId
    }, messageId);
  }

  /**
   * 复制邮件到文件夹
   * 
   * 使用 query-builder 支持调试和自定义 header
   */
  async copy(messageId: string, destinationFolderId: string): Promise<Message> {
    return this.query().post<Message>('copy', {
      destinationId: destinationFolderId
    }, messageId);
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
   * 
   * 使用 query-builder 支持调试和自定义 header
   */
  async getAttachments(messageId: string): Promise<GraphCollection<Attachment>> {
    return await this.query<Attachment>(`${encodeURIComponent(messageId)}/attachments`).get();
  }

  /**
   * 添加附件
   * 
   * 使用 query-builder 支持调试和自定义 header
   */
  async addAttachment(
    messageId: string,
    attachment: Omit<FileAttachment, 'id'>
  ): Promise<Attachment> {
    return this.query().post<Attachment>('attachments', attachment, messageId);
  }

  /**
   * 删除附件
   * 
   * 使用 query-builder 支持调试和自定义 header
   */
  async deleteAttachment(messageId: string, attachmentId: string): Promise<void> {
    return this.query().delete(
      `${encodeURIComponent(messageId)}/attachments/${encodeURIComponent(attachmentId)}`
    );
  }
}

