import { Client } from '@microsoft/microsoft-graph-client';
import { GraphRepository } from '../repository';
import { Event, GraphCollection } from '../types';

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
  query<R = Event>(path?: string) {
    const builder = super.query<R>(path);
    // 自动应用默认时区到查询
    if (this.defaultTimeZone && this.defaultTimeZone !== 'UTC') {
      builder.header('Prefer', `outlook.timezone="${this.defaultTimeZone}"`);
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
   * 
   * 使用 query-builder 支持调试和自定义 header
   */
  async accept(eventId: string, comment?: string): Promise<void> {
    return this.query().post('accept', {
      comment: comment || '',
      sendResponse: true
    }, eventId);
  }

  /**
   * 暂时接受事件邀请
   * 
   * 使用 query-builder 支持调试和自定义 header
   */
  async tentativelyAccept(eventId: string, comment?: string): Promise<void> {
    return this.query().post('tentativelyAccept', {
      comment: comment || '',
      sendResponse: true
    }, eventId);
  }

  /**
   * 拒绝事件邀请
   * 
   * 使用 query-builder 支持调试和自定义 header
   */
  async decline(eventId: string, comment?: string): Promise<void> {
    return this.query().post('decline', {
      comment: comment || '',
      sendResponse: true
    }, eventId);
  }

  /**
   * 取消事件
   * 
   * 使用 query-builder 支持调试和自定义 header
   */
  async cancel(eventId: string, comment?: string): Promise<void> {
    return this.query().post('cancel', {
      comment: comment || ''
    }, eventId);
  }

  /**
   * 获取事件实例（针对循环事件）
   * 
   * 使用 query-builder 支持调试和自定义 header
   */
  async getInstances(
    eventId: string,
    startDateTime: string,
    endDateTime: string
  ): Promise<GraphCollection<Event>> {
    return await this.query(`${encodeURIComponent(eventId)}/instances`)
      .where('start/dateTime', 'ge', '\'' + startDateTime + '\'')
      .and('end/dateTime', 'le', '\'' + endDateTime + '\'').get();
  }

  /**
   * 获取事件附件
   * 
   * 使用 query-builder 支持调试和自定义 header
   */
  async getAttachments(eventId: string): Promise<GraphCollection<unknown>> {
    return await this.query(`${encodeURIComponent(eventId)}/attachments`).get();
  }

  /**
   * 添加事件附件
   * 
   * 使用 query-builder 支持调试和自定义 header
   */
  async addAttachment(eventId: string, attachment: unknown): Promise<unknown> {
    return this.query().post('attachments', attachment, eventId);
  }
}

