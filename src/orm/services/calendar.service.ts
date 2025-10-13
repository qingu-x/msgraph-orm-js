import { Client } from '@microsoft/microsoft-graph-client';
import { Event, Calendar, CalendarGroup, Room, DeltaCollection } from '../types';
import { GraphOrmError, GraphErrorCode } from '../errors';

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
  constructor(private client: Client) {}

  /**
   * 获取用户主日历
   */
  async getUserCalendar(userId: string): Promise<Calendar> {
    try {
      return await this.client
        .api(`/users/${encodeURIComponent(userId)}/calendar`)
        .get();
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 获取用户所有日历
   */
  async getCalendars(userId: string): Promise<Calendar[]> {
    try {
      const response = await this.client
        .api(`/users/${encodeURIComponent(userId)}/calendars`)
        .get();
      return response.value || [];
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 创建日历
   */
  async createCalendar(userId: string, calendar: Omit<Calendar, 'id'>): Promise<Calendar> {
    try {
      return await this.client
        .api(`/users/${encodeURIComponent(userId)}/calendars`)
        .post(calendar);
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 获取日历组
   */
  async getCalendarGroups(userId: string): Promise<CalendarGroup[]> {
    try {
      const response = await this.client
        .api(`/users/${encodeURIComponent(userId)}/calendarGroups`)
        .get();
      return response.value || [];
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 获取日历视图（指定时间范围的事件）
   * 
   * 这是一个特殊的 API，不同于普通的事件查询
   */
  async getCalendarView(
    userId: string,
    startDateTime: string,
    endDateTime: string
  ): Promise<Event[]> {
    try {
      const response = await this.client
        .api(`/users/${encodeURIComponent(userId)}/calendar/calendarView`)
        .query({
          startDateTime,
          endDateTime
        })
        .orderby('start/dateTime')
        .get();
      return response.value || [];
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 获取忙/闲时间表
   * 
   * 查询指定用户或资源的忙闲状态
   */
  async getSchedule(
    schedules: string[],
    startTime: { dateTime: string; timeZone: string },
    endTime: { dateTime: string; timeZone: string },
    availabilityViewInterval: number = 30
  ): Promise<unknown> {
    try {
      return await this.client
        .api('/me/calendar/getSchedule')
        .post({
          schedules,
          startTime,
          endTime,
          availabilityViewInterval
        });
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 查找会议时间
   * 
   * 根据与会者的日历自动寻找可用的会议时间
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
    try {
      return await this.client
        .api('/me/findMeetingTimes')
        .post({
          attendees,
          timeConstraint,
          meetingDuration,
          maxCandidates
        });
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 获取会议室列表（使用 places API）
   * 
   * 权限要求：Place.Read.All
   * 国家云支持：✓ 全球版 ✓ 美国政府版（GCC）⚠️ 中国版（功能受限）
   */
  async getRooms(): Promise<Room[]> {
    try {
      const response = await this.client
        .api('/places/microsoft.graph.room')
        .get();
      return response.value || [];
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 获取会议室列表（使用 places API）
   */
  async getRoomLists(): Promise<Room[]> {
    try {
      const response = await this.client
        .api('/places/microsoft.graph.roomList')
        .get();
      return response.value || [];
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 查找用户可访问的会议室（使用 findRooms 函数）
   * 
   * 权限要求：Calendars.Read
   * 国家云支持：✓ 全球版 ⚠️ 美国政府版（部分支持）❌ 中国版（不可用）
   */
  async findRooms(userId: string): Promise<Room[]> {
    try {
      const response = await this.client
        .api(`/users/${encodeURIComponent(userId)}/findRooms`)
        .post(null);
      return response.value || response || [];
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
   */
  async findRoomLists(userId: string): Promise<Room[]> {
    try {
      const response = await this.client
        .api(`/users/${encodeURIComponent(userId)}/findRoomLists`)
        .post(null);
      return response.value || response || [];
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
   */
  async findRoomsInList(userId: string, roomListEmail: string): Promise<Room[]> {
    try {
      const response = await this.client
        .api(`/users/${encodeURIComponent(userId)}/findRooms(RoomList='${roomListEmail}')`)
        .post(null);
      return response.value || response || [];
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
   */
  async deltaEvents(userId: string, deltaLink?: string): Promise<DeltaCollection<Event>> {
    try {
      const endpoint = deltaLink || `/users/${encodeURIComponent(userId)}/events/delta`;
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
