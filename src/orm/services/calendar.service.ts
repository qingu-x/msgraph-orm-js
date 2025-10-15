import { Client } from '@microsoft/microsoft-graph-client';
import { Event, Calendar, CalendarGroup, Room, GraphCollection, RequestDebugInfo } from '../types';
import { GraphOrmError, GraphErrorCode } from '../errors';
import { GraphQueryBuilder } from '../query-builder';

/**
 * 日历服务
 * 
 * 提供日历相关的业务逻辑功能
 * 基本的事件 CRUD 请使用 EventRepository
 * 
 * 权限要求：Calendars.Read, Calendars.ReadWrite, Place.Read.All
 * 国家云支持：✓ 全球版 ✓ 中国版 ✓ 美国政府版
 * 
 * @see https://learn.microsoft.com/graph/api/resources/calendar
 * @see https://learn.microsoft.com/graph/api/resources/event
 */
export class CalendarService {
  protected lastQueryBuilder?: GraphQueryBuilder<unknown>;
  constructor(private client: Client) {
    this.lastQueryBuilder = undefined;
  }

  /**
   * 创建 query-builder 实例（私有辅助方法）
   */
  private query<T>(endpoint: string): GraphQueryBuilder<T> {
    this.lastQueryBuilder = new GraphQueryBuilder<T>(this.client, endpoint) as GraphQueryBuilder<unknown>;
    return this.lastQueryBuilder as GraphQueryBuilder<T>;
  }

  getDebug(): RequestDebugInfo {
    const debugInfo = this.lastQueryBuilder?.getDebug();
    if (!debugInfo) {
      throw new GraphOrmError({
        code: GraphErrorCode.RESOURCE_NOT_FOUND,
        message: '调试信息未找到',
        timestamp: new Date()
      });
    }
    return debugInfo;
  }

  /**
   * 获取用户主日历
   * 
   * 使用 query-builder 支持调试和自定义 header
   */
  async getUserCalendar(userId: string): Promise<Calendar | null> {
    return this.query<Calendar>(`users/${encodeURIComponent(userId)}/calendar`).first();
  }

  /**
   * 获取用户所有日历
   * 
   * 使用 query-builder 支持调试和自定义 header
   */
  async getCalendars(userId: string): Promise<GraphCollection<Calendar>> {
    return await this.query<Calendar>(`users/${encodeURIComponent(userId)}/calendars`).get() as GraphCollection<Calendar>;
  }

  /**
   * 创建日历
   * 
   * 使用 query-builder 支持调试和自定义 header
   */
  async createCalendar(userId: string, calendar: Omit<Calendar, 'id'>): Promise<Calendar> {
    return this.query('').post<Calendar>('calendars', calendar, `users/${encodeURIComponent(userId)}`);
  }

  /**
   * 获取日历组
   * 
   * 使用 query-builder 支持调试和自定义 header
   */
  async getCalendarGroups(userId: string): Promise<GraphCollection<CalendarGroup>> {
    return await this.query<CalendarGroup>(`users/${encodeURIComponent(userId)}/calendarGroups`).get() as GraphCollection<CalendarGroup>;
  }

  /**
   * 获取日历视图（指定时间范围的事件）
   * 
   * 这是一个特殊的 API，不同于普通的事件查询
   * 使用 query-builder 支持调试和自定义 header
   */
  async getCalendarView(
    userId: string,
    startDateTime: string,
    endDateTime: string
  ): Promise<GraphCollection<Event>> {
    return await this.query<Event>(`users/${encodeURIComponent(userId)}/calendar/calendarView`,)
    .where('start/dateTime', 'ge', '\'' + startDateTime + '\'')
    .where('end/dateTime', 'le', '\'' + endDateTime + '\'') 
    .get();
  }

  /**
   * 获取忙/闲时间表
   * 
   * 查询指定用户或资源的忙闲状态
   * 使用 query-builder 支持调试和自定义 header
   */
  async getSchedule(
    schedules: string[],
    startTime: { dateTime: string; timeZone: string },
    endTime: { dateTime: string; timeZone: string },
    availabilityViewInterval: number = 30
  ): Promise<unknown> {
    return this.query('').post('calendar/getSchedule', {
      schedules,
      startTime,
      endTime,
      availabilityViewInterval
    }, 'me');
  }

  /**
   * 查找会议时间
   * 
   * 根据与会者的日历自动寻找可用的会议时间
   * 使用 query-builder 支持调试和自定义 header
   */
  async findMeetingTimes(
    attendees: Array<{ emailAddress: { address: string; name?: string }; type: string }>,
    timeConstraint: {
      activityDomain?: string;
      timeslots?: Array<{
        start: { dateTime: string; timeZone: string };
        end: { dateTime: string; timeZone: string };
      }>;
    },
    meetingDuration?: string,
    maxCandidates?: number
  ): Promise<unknown> {
    return this.query('').post('findMeetingTimes', {
      attendees,
      timeConstraint,
      meetingDuration,
      maxCandidates
    }, 'me');
  }

  /**
   * 获取会议室列表（使用 places API）
   * 
   * 权限要求：Place.Read.All
   * 国家云支持：✓ 全球版 ✓ 美国政府版（GCC）⚠️ 中国版（功能受限）
   * 使用 query-builder 支持调试和自定义 header
   */
  async getRooms(): Promise<GraphCollection<Room>> {
    return await this.query<Room>('places/microsoft.graph.room').get();
  }

  /**
   * 获取会议室列表（使用 places API）
   * 
   * 使用 query-builder 支持调试和自定义 header
   */
  async getRoomLists(): Promise<GraphCollection<Room>> {
    return await this.query<Room>('places/microsoft.graph.roomList').get();
  }

  /**
   * 查找用户可访问的会议室（使用 findRooms 函数）
   * 
   * 权限要求：Calendars.Read
   * 国家云支持：✓ 全球版 ⚠️ 美国政府版（部分支持）❌ 中国版（不可用）
   * 使用 query-builder 支持调试和自定义 header
   */
  async findRooms(userId: string): Promise<Room[]> {
    try {
      const response = await this.query('').post<{ value: Room[] } | Room[]>(
        'findRooms', null, `users/${encodeURIComponent(userId)}`
      );
      return Array.isArray(response) ? response : (response.value || []);
    } catch (error) {
      const ormError = GraphOrmError.fromGraphError(error);
      if (ormError.code === GraphErrorCode.RESOURCE_NOT_FOUND || 
          ormError.code === GraphErrorCode.BAD_REQUEST) {
        return [];
      }
      throw ormError;
    }
  }

  /**
   * 查找会议室列表
   * 
   * 使用 query-builder 支持调试和自定义 header
   */
  async findRoomLists(userId: string): Promise<Room[]> {
    try {
      const response = await this.query('').post<{ value: Room[] } | Room[]>(
        'findRoomLists', null, `users/${encodeURIComponent(userId)}`
      );
      return Array.isArray(response) ? response : (response.value || []);
    } catch (error) {
      const ormError = GraphOrmError.fromGraphError(error);
      if (ormError.code === GraphErrorCode.RESOURCE_NOT_FOUND || 
          ormError.code === GraphErrorCode.BAD_REQUEST) {
        return [];
      }
      throw ormError;
    }
  }

  /**
   * 查找指定会议室列表下的会议室
   * 
   * 使用 query-builder 支持调试和自定义 header
   */
  async findRoomsInList(userId: string, roomListEmail: string): Promise<Room[]> {
    try {
      const response = await this.query('').post<{ value: Room[] } | Room[]>(
        `findRooms(RoomList='${roomListEmail}')`, null, `users/${encodeURIComponent(userId)}`
      );
      return Array.isArray(response) ? response : (response.value || []);
    } catch (error) {
      const ormError = GraphOrmError.fromGraphError(error);
      if (ormError.code === GraphErrorCode.RESOURCE_NOT_FOUND || 
          ormError.code === GraphErrorCode.BAD_REQUEST) {
        return [];
      }
      throw ormError;
    }
  }

  /**
   * 事件增量查询
   * 
   * 跟踪事件的变化（新增、修改、删除）
   * 使用 query-builder 支持调试和自定义 header
   */
  async deltaEvents(userId: string, deltaLink?: string): Promise<GraphCollection<Event>> {
    const endpoint = deltaLink || `users/${encodeURIComponent(userId)}/events/delta`;
    return await this.query<Event>(endpoint).get();
  }
}
