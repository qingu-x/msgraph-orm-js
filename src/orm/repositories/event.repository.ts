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
  private defaultTimeZone: string = 'UTC';

  constructor(client: Client, endpoint: string, defaultTimeZone?: string) {
    super(client, endpoint);
    if (defaultTimeZone) {
      this.defaultTimeZone = defaultTimeZone;
    }
  }

  /**
   * 设置默认时区
   * 
   * @param timeZone IANA 时区名称，如 'Asia/Shanghai', 'America/New_York', 'UTC'
   */
  setDefaultTimeZone(timeZone: string): void {
    this.defaultTimeZone = timeZone;
  }

  /**
   * 获取当前默认时区
   */
  getDefaultTimeZone(): string {
    return this.defaultTimeZone;
  }

  /**
   * 自动为事件的时间字段添加时区
   */
  private applyTimeZone(event: Partial<Event>): Partial<Event> {
    const result = { ...event };
    
    if (result.start && !result.start.timeZone) {
      result.start = { ...result.start, timeZone: this.defaultTimeZone };
    }
    
    if (result.end && !result.end.timeZone) {
      result.end = { ...result.end, timeZone: this.defaultTimeZone };
    }
    
    return result;
  }

  /**
   * 创建查询构建器（自动应用默认时区）
   * 
   * 查询时会自动在 HTTP Header 中添加 Prefer: outlook.timezone，
   * 使返回的事件时间使用指定的时区
   */
  query() {
    const builder = super.query();
    // 自动应用默认时区到查询
    if (this.defaultTimeZone && this.defaultTimeZone !== 'UTC') {
      builder.timezone(this.defaultTimeZone);
    }
    return builder;
  }

  /**
   * 创建事件（自动应用默认时区）
   */
  async create(entity: Omit<Event, 'id'>): Promise<Event> {
    const eventWithTimeZone = this.applyTimeZone(entity as Partial<Event>);
    return super.create(eventWithTimeZone as Omit<Event, 'id'>);
  }

  /**
   * 更新事件（自动应用默认时区）
   */
  async update(id: string, entity: Partial<Event>): Promise<Event> {
    const eventWithTimeZone = this.applyTimeZone(entity);
    return super.update(id, eventWithTimeZone);
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

