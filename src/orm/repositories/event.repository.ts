import { Client } from '@microsoft/microsoft-graph-client';
import { GraphRepository } from '../repository';
import { Event } from '../types';
import { GraphOrmError } from '../errors';

/**
 * 事件仓储
 * 
 * 权限要求：Calendars.Read, Calendars.ReadWrite
 * 国家云支持：✓ 全球版 ✓ 中国版 ✓ 美国政府版
 * 
 * @see https://learn.microsoft.com/graph/api/resources/event
 */
export class EventRepository extends GraphRepository<Event> {
  constructor(client: Client, endpoint: string) {
    super(client, endpoint);
  }

  /**
   * 接受事件邀请
   */
  async accept(eventId: string, comment?: string): Promise<void> {
    try {
      await this.client
        .api(`${this.endpoint}/${encodeURIComponent(eventId)}/accept`)
        .post({
          comment: comment || '',
          sendResponse: true
        });
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 暂时接受事件邀请
   */
  async tentativelyAccept(eventId: string, comment?: string): Promise<void> {
    try {
      await this.client
        .api(`${this.endpoint}/${encodeURIComponent(eventId)}/tentativelyAccept`)
        .post({
          comment: comment || '',
          sendResponse: true
        });
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 拒绝事件邀请
   */
  async decline(eventId: string, comment?: string): Promise<void> {
    try {
      await this.client
        .api(`${this.endpoint}/${encodeURIComponent(eventId)}/decline`)
        .post({
          comment: comment || '',
          sendResponse: true
        });
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 取消事件
   */
  async cancel(eventId: string, comment?: string): Promise<void> {
    try {
      await this.client
        .api(`${this.endpoint}/${encodeURIComponent(eventId)}/cancel`)
        .post({
          comment: comment || ''
        });
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 获取事件实例（针对循环事件）
   */
  async getInstances(
    eventId: string,
    startDateTime: string,
    endDateTime: string
  ): Promise<Event[]> {
    try {
      const response = await this.client
        .api(`${this.endpoint}/${encodeURIComponent(eventId)}/instances`)
        .query({
          startDateTime,
          endDateTime
        })
        .get();
      return response.value || [];
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 获取事件附件
   */
  async getAttachments(eventId: string): Promise<unknown[]> {
    try {
      const response = await this.client
        .api(`${this.endpoint}/${encodeURIComponent(eventId)}/attachments`)
        .get();
      return response.value || [];
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 添加事件附件
   */
  async addAttachment(eventId: string, attachment: unknown): Promise<unknown> {
    try {
      return await this.client
        .api(`${this.endpoint}/${encodeURIComponent(eventId)}/attachments`)
        .post(attachment);
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }
}

