import { ClientSecretCredential, ClientSecretCredentialOptions } from '@azure/identity';
import { TokenCredentialAuthenticationProvider } from '@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials/index.js';
import { AuthProvider, Client, ClientOptions, Options } from '@microsoft/microsoft-graph-client';
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

export interface GraphClientOptions {
  endpoints: Endpoints;
  params: ClientParams | string;
  options?: ClientSecretCredentialOptions;
}

export class GraphClient<T> {
  private credential?: ClientSecretCredential;
  private graphClient: Client;
  private clientOptions: ClientOptions | Options = {};
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  private proxyAgent: any | undefined;
  private meetings: T[] = [];
  private endpoints: Endpoints;

  constructor({
    endpoints,
    params,
    options
  }: GraphClientOptions) {
    this.endpoints = endpoints;
    // 创建 HTTPS 代理（仅在 Node.js 环境中）
    this.proxyAgent = options?.proxyOptions && HttpsProxyAgent
      ? new HttpsProxyAgent("http://" + options.proxyOptions.host + ":" + options.proxyOptions.port) 
      : undefined;
    if(typeof params === 'string') {
      const tokenAuthProvider: AuthProvider = async (done: (error: Error | null, token: string | null) => void) => {
        try {
          done(null, params);
        } catch (error) {
          console.error('Token auth provider error:', error);
          done(null, null);
        }
      };
      this.graphClient = Client.init({
        authProvider: tokenAuthProvider,
        baseUrl: this.endpoints.graph,
        fetchOptions: this.proxyAgent ? {
          agent: this.proxyAgent,
        } : undefined,
      });
    }else {
      const credential: ClientSecretCredential = new ClientSecretCredential(
        params.tenantId,
        params.clientId,
        params.clientSecret,
        {
          authorityHost: endpoints.azureAD,
          ...options,
        }
      );
      const authProvider = new TokenCredentialAuthenticationProvider(credential, {
        scopes: [this.endpoints.graph + '/.default'],
      });
      this.clientOptions = {
        authProvider,
        baseUrl: this.endpoints.graph,
        fetchOptions: this.proxyAgent ? {
          agent: this.proxyAgent,
        } : undefined,
      }
      this.graphClient = Client.initWithMiddleware(this.clientOptions);
    }
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

  public updateToken(token: string) {
    if(this.clientOptions && this.clientOptions.authProvider && !(this.clientOptions.authProvider instanceof TokenCredentialAuthenticationProvider)) {
      this.clientOptions.authProvider = async (done: (error: Error | null, token: string | null) => void) => {
        try {
          done(null, token);
        } catch (error) {
          console.error('Auth provider error:', error);
          done(null, null);
        }
      };
      this.graphClient = Client.init(this.clientOptions as Options);
    }
  }

  public getGraphClient() {
    return this.graphClient;
  }
}

export default GraphClient;
