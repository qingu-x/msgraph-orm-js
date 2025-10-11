import ky, { type KyInstance } from "ky";

interface BaseOptions {
  baseURL: string;
}

interface DemoParams {
  [key: string]: unknown;
}

interface ApiResponse<T = unknown> {
  data: T;
  status: number;
  message?: string;
}

class Demo<T> {
  private ky: KyInstance;
  private demos: T[];

  constructor(options: BaseOptions) {
    this.ky = ky.create({
      prefixUrl: options.baseURL,
      headers: {
        "content-type": "application/json",
      },
    });
    this.demos = [];
  }

  async getDemoList(params: DemoParams): Promise<ApiResponse<T[]>> {
    const response = await this.ky.get("demos", { 
      searchParams: new URLSearchParams(params as Record<string, string>) 
    }).json<ApiResponse<T[]>>();
    return response;
  }
}

export default Demo;
