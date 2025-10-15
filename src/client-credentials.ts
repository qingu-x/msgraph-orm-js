import { ClientSecretCredential, ClientSecretCredentialOptions } from '@azure/identity';
import { TokenCredentialAuthenticationProvider } from '@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials/index.js';
import { AuthProvider, Client } from '@microsoft/microsoft-graph-client';
import { Endpoints } from './types';

// 动态导入 https-proxy-agent，仅在 Node.js 环境中使用
// eslint-disable-next-line @typescript-eslint/no-explicit-any
let HttpsProxyAgent: any;
if (typeof window === 'undefined') {
  // Node.js 环境
  try {
    // eslint-disable-next-line @typescript-eslint/no-require-imports
    HttpsProxyAgent = require('https-proxy-agent').HttpsProxyAgent;
  } catch {
    // 如果导入失败，忽略
  }
}

export interface ClientParams {
  tenantId: string;
  clientId: string;
  clientSecret: string;
  options?: ClientSecretCredentialOptions;
}

export class GraphClient<T> {
  private credential: ClientSecretCredential;
  private graphClient!: Client;
  private authProvider: TokenCredentialAuthenticationProvider | AuthProvider;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  private proxyAgent: any | undefined;
  private meetings: T[] = [];
  private endpoints: Endpoints;

  constructor(params: ClientParams, endpoints: Endpoints, token?: string) {
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
    this.authProvider = token ? async (done) => {
      try {
        done(null, token);
      } catch (error) {
        console.error('Auth provider error:', error);
        done(error, null);
      }
    } : new TokenCredentialAuthenticationProvider(this.credential, {
      scopes: [this.endpoints.graph + '/.default'],
    });

    // 创建 HTTPS 代理（仅在 Node.js 环境中）
    this.proxyAgent = params.options?.proxyOptions && HttpsProxyAgent
      ? new HttpsProxyAgent("http://" + params.options.proxyOptions.host + ":" + params.options.proxyOptions.port) 
      : undefined;

    this.generateGraphClient();
  }

  private generateGraphClient() {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const config: any = {
      authProvider: this.authProvider,
      baseUrl: this.endpoints.graph,
    };
    
    // 仅在 Node.js 环境且有代理时添加 fetchOptions
    if (this.proxyAgent) {
      config.fetchOptions = {
        agent: this.proxyAgent,
      };
    }
    this.graphClient = this.authProvider instanceof TokenCredentialAuthenticationProvider ? Client.initWithMiddleware(config) : Client.init(config);
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
    // 创建 HTTPS 代理（仅在 Node.js 环境中）
    this.proxyAgent = params.options?.proxyOptions && HttpsProxyAgent
      ? new HttpsProxyAgent("http://" + params.options.proxyOptions.host + ":" + params.options.proxyOptions.port) 
      : undefined;
    this.generateGraphClient();
  }

  public updateToken(token: string) {
    this.authProvider = async (done) => {
      try {
        done(null, token);
      } catch (error) {
        console.error('Auth provider error:', error);
        done(error, null);
      }
    };
    this.generateGraphClient();
  }

  public getGraphClient() {
    return this.graphClient;
  }
}

export default GraphClient;
