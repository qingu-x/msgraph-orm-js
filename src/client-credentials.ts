import { ClientSecretCredential, ClientSecretCredentialOptions } from '@azure/identity';
import { TokenCredentialAuthenticationProvider } from '@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials/index.js';
import { HttpsProxyAgent } from 'https-proxy-agent';
import { Client } from '@microsoft/microsoft-graph-client';
import { Endpoints } from './types';

export interface ClientParams {
  tenantId: string;
  clientId: string;
  clientSecret: string;
  options?: ClientSecretCredentialOptions;
}

export class GraphClient<T> {
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
}

export default GraphClient;
