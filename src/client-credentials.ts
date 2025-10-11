import { ClientSecretCredential, ClientSecretCredentialOptions } from '@azure/identity';
import { TokenCredentialAuthenticationProvider } from '@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials';
import { HttpsProxyAgent } from 'https-proxy-agent';
import { Client } from '@microsoft/microsoft-graph-client';
import { Endpoints, MeetingParams } from './types';

export interface ClientParams {
  tenantId: string;
  clientId: string;
  clientSecret: string;
  options?: ClientSecretCredentialOptions;
}

export class ApplicationClient<T> {
  private credential: ClientSecretCredential;
  private graphClient!: Client;
  private authProvider: TokenCredentialAuthenticationProvider;
  private proxyAgent: HttpsProxyAgent<string> | undefined;
  private meetings: T[] = [];
  private endpoints: Endpoints;

  constructor(params: ClientParams, endpoints: Endpoints) {
    this.endpoints = endpoints;
    this.credential = new ClientSecretCredential(
      params.tenantId,
      params.clientId,
      params.clientSecret,
      {
        authorityHost: endpoints.azureAD,
        ...params.options,
      }
    );

    // 客户端凭据流使用 /.default scope
    this.authProvider = new TokenCredentialAuthenticationProvider(this.credential, {
      scopes: [this.endpoints.graph + '/.default'],
    });

    // 创建 HTTPS 代理代理
    this.proxyAgent = params.options?.proxyOptions 
      ? new HttpsProxyAgent("http://" + params.options.proxyOptions.host + ":" + params.options.proxyOptions.port) 
      : undefined;

    this.generateGraphClient();
  }

  private generateGraphClient() {
    this.graphClient = Client.initWithMiddleware({
      authProvider: this.authProvider,
      baseUrl: this.endpoints.graph,
      fetchOptions: {
        agent: this.proxyAgent,
      },
    });
  }

  // 规范化查询对象：移除 undefined，将 boolean 转为字符串
  private buildQuery(input: { [key: string]: string | number | boolean | undefined }) {
    const out: { [key: string]: string | number } = {};
    for (const key in input) {
      const val = input[key];
      if (val !== undefined) {
        out[key] = typeof val === 'boolean' ? String(val) : val;
      }
    }
    return out;
  }

  public updateClientParams(params: ClientParams) {
    this.credential = new ClientSecretCredential(
      params.tenantId,
      params.clientId,
      params.clientSecret,
      {
        authorityHost: this.endpoints.azureAD,
        ...params.options,
      }
    );
    this.proxyAgent = params.options?.proxyOptions 
      ? new HttpsProxyAgent("http://" + params.options.proxyOptions.host + ":" + params.options.proxyOptions.port) 
      : undefined;
    this.generateGraphClient();
  }

  public getGraphClient() {
    return this.graphClient;
  }

  // === 用户管理 ===
  // 搜索用户（支持 $search 或 $filter），带分页
  public async searchUsers(query: { search?: string; filter?: string; top?: number; skiptoken?: string; count?: boolean } = {}) {
    let request = this.graphClient.api("/users");
    if (query.search) {
      request = request
        .header("ConsistencyLevel", "eventual")
        .query(this.buildQuery({
          $search: `"${query.search}"`,
          $count: query.count === false ? undefined : true,
          $top: query.top,
          $skiptoken: query.skiptoken,
        }));
    } else {
      request = request.query(
        this.buildQuery({
          $filter: query.filter,
          $top: query.top,
          $skiptoken: query.skiptoken,
        })
      );
    }
    return await request.get();
  }

  // 获取用户详情（支持 userId 或 userPrincipalName）
  public async getUser(userIdOrUpn: string) {
    return await this.graphClient.api(`/users/${encodeURIComponent(userIdOrUpn)}`).get();
  }

  // === 会议管理（应用权限需指定用户） ===
  // 搜索/分页某用户的会议（日历事件）
  public async searchUserEvents(userIdOrUpn: string, query: { top?: number; skiptoken?: string; filter?: string; orderby?: string } = {}) {
    return await this.graphClient
      .api(`/users/${encodeURIComponent(userIdOrUpn)}/events`)
      .query(
        this.buildQuery({
          $top: query.top,
          $skiptoken: query.skiptoken,
          $filter: query.filter,
          $orderby: query.orderby,
        })
      )
      .get();
  }

  // 查看会议详情
  public async getUserEvent(userIdOrUpn: string, eventId: string) {
    return await this.graphClient.api(`/users/${encodeURIComponent(userIdOrUpn)}/events/${encodeURIComponent(eventId)}`).get();
  }

  // 创建会议（在用户日历中创建事件）
  public async createUserEvent(userIdOrUpn: string, eventPayload: unknown) {
    return await this.graphClient.api(`/users/${encodeURIComponent(userIdOrUpn)}/events`).post(eventPayload);
  }

  // === 会议室 ===
  // 列出所有会议室
  public async listRooms(params: MeetingParams = {}) {
    let request = this.graphClient.api("/places/microsoft.graph.room");
    if (Object.keys(params).length > 0) {
      request = request.query(params);
    }
    return await request.get();
  }

  // 列出会议室列表（Room Lists）
  public async listRoomLists(params: MeetingParams = {}) {
    let request = this.graphClient.api("/places/microsoft.graph.roomList");
    if (Object.keys(params).length > 0) {
      request = request.query(params);
    }
    return await request.get();
  }

  // 查看指定会议室列表下的会议室
  public async listRoomsInRoomList(roomListAddressOrId: string, params: MeetingParams = {}) {
    let request = this.graphClient.api(`/places/${encodeURIComponent(roomListAddressOrId)}/microsoft.graph.roomList/rooms`);
    if (Object.keys(params).length > 0) {
      request = request.query(params);
    }
    return await request.get();
  }
}

export default ApplicationClient;
