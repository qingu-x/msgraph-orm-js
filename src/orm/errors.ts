// Graph ORM 统一错误处理

export enum GraphErrorCode {
  // 通用错误
  UNKNOWN_ERROR = 'UNKNOWN_ERROR',
  NETWORK_ERROR = 'NETWORK_ERROR',
  
  // 权限错误
  UNAUTHORIZED = 'UNAUTHORIZED',
  FORBIDDEN = 'FORBIDDEN',
  
  // 资源错误
  NOT_FOUND = 'NOT_FOUND',
  RESOURCE_NOT_FOUND = 'RESOURCE_NOT_FOUND',
  NO_ENTITY_FOUND = 'NO_ENTITY_FOUND',
  
  // 请求错误
  BAD_REQUEST = 'BAD_REQUEST',
  INVALID_PARAMETER = 'INVALID_PARAMETER',
  
  // API 错误
  API_NOT_AVAILABLE = 'API_NOT_AVAILABLE',
  QUOTA_EXCEEDED = 'QUOTA_EXCEEDED',
}

export interface GraphErrorDetails {
  code: GraphErrorCode;
  message: string;
  statusCode?: number;
  originalError?: unknown;
  requestId?: string;
  timestamp: Date;
}

export class GraphOrmError extends Error {
  public readonly code: GraphErrorCode;
  public readonly statusCode?: number;
  public readonly originalError?: unknown;
  public readonly requestId?: string;
  public readonly timestamp: Date;

  constructor(details: GraphErrorDetails) {
    super(details.message);
    this.name = 'GraphOrmError';
    this.code = details.code;
    this.statusCode = details.statusCode;
    this.originalError = details.originalError;
    this.requestId = details.requestId;
    this.timestamp = details.timestamp;

    // 保持原型链
    Object.setPrototypeOf(this, GraphOrmError.prototype);
  }

  /**
   * 从 Microsoft Graph 错误创建 GraphOrmError
   */
  static fromGraphError(error: unknown): GraphOrmError {
    // 如果已经是 GraphOrmError，直接返回，避免重复封装
    if (error instanceof GraphOrmError) {
      return error;
    }

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const graphError = error as any;
    
    const statusCode = graphError?.statusCode;
    const errorCode = graphError?.code;
    const errorMessage = graphError?.message || '未知错误';
    const requestId = graphError?.requestId;

    // 根据状态码映射错误类型
    let code: GraphErrorCode;
    let message: string;

    switch (statusCode) {
      case 400:
        code = GraphErrorCode.BAD_REQUEST;
        message = `请求错误: ${errorMessage}`;
        break;
      case 401:
        code = GraphErrorCode.UNAUTHORIZED;
        message = `未授权: ${errorMessage}`;
        break;
      case 403:
        code = GraphErrorCode.FORBIDDEN;
        message = `权限不足: ${errorMessage}`;
        break;
      case 404:
        code = GraphErrorCode.RESOURCE_NOT_FOUND;
        message = `资源不存在: ${errorMessage}`;
        break;
      case 429:
        code = GraphErrorCode.QUOTA_EXCEEDED;
        message = `请求频率超限: ${errorMessage}`;
        break;
      default:
        code = GraphErrorCode.UNKNOWN_ERROR;
        message = `未知错误 (${statusCode}): ${errorMessage}`;
    }

    return new GraphOrmError({
      code,
      message,
      statusCode,
      originalError: error,
      requestId,
      timestamp: new Date(),
    });
  }

  /**
   * 创建自定义错误
   */
  static create(code: GraphErrorCode, message: string, statusCode?: number): GraphOrmError {
    return new GraphOrmError({
      code,
      message,
      statusCode,
      timestamp: new Date(),
    });
  }

  /**
   * 转换为 JSON 格式
   */
  toJSON() {
    return {
      name: this.name,
      code: this.code,
      message: this.message,
      statusCode: this.statusCode,
      requestId: this.requestId,
      timestamp: this.timestamp.toISOString(),
    };
  }
}

