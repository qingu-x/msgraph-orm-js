import { Client } from '@microsoft/microsoft-graph-client';
import { Event, Calendar, CalendarGroup, Room, DeltaCollection } from '../types';
import { GraphOrmError, GraphErrorCode } from '../errors';

/**
 * 日历服务
 * 
 * 提供日历和事件管理功能
 * 
 * 权限要求：Calendars.Read, Calendars.ReadWrite
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
   * 获取事件列表
   */
  async getEvents(userId: string, top?: number): Promise<Event[]> {
    try {
      let request = this.client
        .api(`/users/${encodeURIComponent(userId)}/events`)
        .orderby('start/dateTime');
      
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
   * 获取日历视图（指定时间范围的事件）
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
   * 获取单个事件
   */
  async getEvent(userId: string, eventId: string): Promise<Event> {
    try {
      return await this.client
        .api(`/users/${encodeURIComponent(userId)}/events/${encodeURIComponent(eventId)}`)
        .get();
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 创建事件
   */
  async createEvent(userId: string, event: Omit<Event, 'id'>): Promise<Event> {
    try {
      return await this.client
        .api(`/users/${encodeURIComponent(userId)}/events`)
        .post(event);
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 更新事件
   */
  async updateEvent(userId: string, eventId: string, updates: Partial<Event>): Promise<Event> {
    try {
      return await this.client
        .api(`/users/${encodeURIComponent(userId)}/events/${encodeURIComponent(eventId)}`)
        .patch(updates);
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 删除事件
   */
  async deleteEvent(userId: string, eventId: string): Promise<void> {
    try {
      await this.client
        .api(`/users/${encodeURIComponent(userId)}/events/${encodeURIComponent(eventId)}`)
        .delete();
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 接受事件邀请
   */
  async acceptEvent(userId: string, eventId: string, comment?: string): Promise<void> {
    try {
      await this.client
        .api(`/users/${encodeURIComponent(userId)}/events/${encodeURIComponent(eventId)}/accept`)
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
  async tentativelyAcceptEvent(userId: string, eventId: string, comment?: string): Promise<void> {
    try {
      await this.client
        .api(`/users/${encodeURIComponent(userId)}/events/${encodeURIComponent(eventId)}/tentativelyAccept`)
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
  async declineEvent(userId: string, eventId: string, comment?: string): Promise<void> {
    try {
      await this.client
        .api(`/users/${encodeURIComponent(userId)}/events/${encodeURIComponent(eventId)}/decline`)
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
  async cancelEvent(userId: string, eventId: string, comment?: string): Promise<void> {
    try {
      await this.client
        .api(`/users/${encodeURIComponent(userId)}/events/${encodeURIComponent(eventId)}/cancel`)
        .post({
          comment: comment || ''
        });
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 获取忙/闲时间表
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

  /**
   * 获取事件实例（针对循环事件）
   */
  async getEventInstances(
    userId: string,
    eventId: string,
    startDateTime: string,
    endDateTime: string
  ): Promise<Event[]> {
    try {
      const response = await this.client
        .api(`/users/${encodeURIComponent(userId)}/events/${encodeURIComponent(eventId)}/instances`)
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
  async getEventAttachments(userId: string, eventId: string): Promise<unknown[]> {
    try {
      const response = await this.client
        .api(`/users/${encodeURIComponent(userId)}/events/${encodeURIComponent(eventId)}/attachments`)
        .get();
      return response.value || [];
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }

  /**
   * 添加事件附件
   */
  async addEventAttachment(userId: string, eventId: string, attachment: unknown): Promise<unknown> {
    try {
      return await this.client
        .api(`/users/${encodeURIComponent(userId)}/events/${encodeURIComponent(eventId)}/attachments`)
        .post(attachment);
    } catch (error) {
      throw GraphOrmError.fromGraphError(error);
    }
  }
}

