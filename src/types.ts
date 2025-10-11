export interface Endpoints {
  azure: string;
  azureAD: string;
  graph: string;
}

export const defaultEndpoints: { [key: string]: Endpoints } = {
  global: {
    azure: 'https://portal.azure.com',
    azureAD: 'https://login.microsoftonline.com',
    graph: 'https://graph.microsoft.com',
  },
  usGCC: {
    azure: 'https://portal.azure.com',
    azureAD: 'https://login.microsoftonline.us',
    graph: 'https://graph.microsoft.com',
  },
  usGCCHigh: {
    azure: 'https://portal.azure.us',
    azureAD: 'https://login.microsoftonline.us',
    graph: 'https://graph.microsoft.us',
  },
  usDOD: {
    azure: 'https://portal.azure.us',
    azureAD: 'https://login.microsoftonline.us',
    graph: 'https://dod-graph.microsoft.us',
  },
  cn: {
    azure: 'https://portal.azure.cn',
    azureAD: 'https://login.chinacloudapi.cn',
    graph: 'https://microsoftgraph.chinacloudapi.cn',
  }
};

export interface MeetingParams {
  [key: string]: string | number;
}
