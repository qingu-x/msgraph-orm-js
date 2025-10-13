// Graph ORM 类型定义
export interface GraphEntity {
  id: string;
  '@odata.etag'?: string;
  '@odata.context'?: string;
}

export interface GraphCollection<T> {
  meta: {
    nextLink?: string;
    count?: number;
  }
  data: T[];
}

// 查询操作符
export type QueryOperator = 
  | 'eq' | 'ne' | 'gt' | 'ge' | 'lt' | 'le'
  | 'startswith' | 'endswith' | 'contains'
  | 'and' | 'or' | 'not';

// 排序方向
export type OrderDirection = 'asc' | 'desc';

// 查询条件
export interface QueryCondition {
  field: string;
  operator: QueryOperator;
  value: string | number | boolean | null;
  logicalOperator?: 'and' | 'or'; // 与前一个条件的逻辑关系
}

// 查询构建器接口
export interface QueryBuilder<T> {
  where(field: string, operator: QueryOperator | string | number | boolean | null, value?: string | number | boolean | null): QueryBuilder<T>;
  and(field: string, operator: QueryOperator | string | number | boolean | null, value?: string | number | boolean | null): QueryBuilder<T>;
  or(field: string, operator: QueryOperator | string | number | boolean | null, value?: string | number | boolean | null): QueryBuilder<T>;
  orderBy(field: string, direction?: OrderDirection): QueryBuilder<T>;
  select(fields: string[]): QueryBuilder<T>;
  top(count: number): QueryBuilder<T>;
  skip(count: number): QueryBuilder<T>;
  skipToken(token: string): QueryBuilder<T>;
  search(field: string, value: string): QueryBuilder<T>;
  expand(property: string | string[]): QueryBuilder<T>;
  format(format: 'json' | 'atom'): QueryBuilder<T>;
  count(): QueryBuilder<T>;
  getRaw(): {
    endpoint: string;
    params: Record<string, string | number>;
    headers: Record<string, string>;
    url: string;
  };
  get(): Promise<GraphCollection<T>>;
  first(): Promise<T>;
  pagination(): AsyncIterableIterator<T>;
}

// CRUD 操作的原始请求信息
export interface CrudRawRequest {
  method: 'GET' | 'POST' | 'PATCH' | 'DELETE';
  endpoint: string;
  url: string;
  headers?: Record<string, string>;
  body?: Record<string, unknown>;
  params?: Record<string, string | number>;
}

// 实体管理器接口
export interface EntityManager<T extends GraphEntity> {
  findById(id: string): Promise<T>;
  findOne(query: Partial<T>): Promise<T | null>;
  findMany(query?: Partial<T>): Promise<GraphCollection<T>>;
  create(entity: Omit<T, 'id'>): Promise<T>;
  update(id: string, entity: Partial<T>): Promise<T>;
  delete(id: string): Promise<void>;
  query(): QueryBuilder<T>;
  
  // 调试方法：返回原始请求信息而不执行
  getRawFindById(id: string): CrudRawRequest;
  getRawCreate(entity: Omit<T, 'id'>): CrudRawRequest;
  getRawUpdate(id: string, entity: Partial<T>): CrudRawRequest;
  getRawDelete(id: string): CrudRawRequest;
}

// 用户实体
export interface User extends GraphEntity {
  displayName: string;
  givenName?: string;
  surname?: string;
  mail?: string;
  userPrincipalName: string;
  jobTitle?: string;
  officeLocation?: string;
  businessPhones: string[];
  mobilePhone?: string;
  preferredLanguage?: string;
}

// 设备实体
export interface Device extends GraphEntity {
  displayName: string;
  deviceId?: string;
  operatingSystem?: string;
  operatingSystemVersion?: string;
  isCompliant?: boolean;
  isManaged?: boolean;
  trustType?: string;
  alternativeSecurityIds?: unknown[];
  physicalIds?: unknown[];
  extensionAttributes?: Record<string, unknown>;
}

// 事件实体
export interface Event extends GraphEntity {
  subject: string;
  body?: {
    contentType: string;
    content: string;
  };
  start: {
    dateTime: string;
    timeZone: string;
  };
  end: {
    dateTime: string;
    timeZone: string;
  };
  location?: {
    displayName: string;
    address?: Record<string, unknown>;
  };
  attendees?: Array<{
    type: string;
    status: {
      response: string;
      time: string;
    };
    emailAddress: {
      name: string;
      address: string;
    };
  }>;
  organizer?: {
    emailAddress: {
      name: string;
      address: string;
    };
  };
  isAllDay?: boolean;
  isCancelled?: boolean;
  isOnlineMeeting?: boolean;
  onlineMeetingUrl?: string;
  showAs?: string;
  importance?: string;
  sensitivity?: string;
}

// 会议室实体
export interface Room extends GraphEntity {
  displayName: string;
  emailAddress: string;
  address?: {
    street?: string;
    city?: string;
    state?: string;
    countryOrRegion?: string;
    postalCode?: string;
  };
  capacity?: number;
  floor?: string;
  label?: string;
  nickname?: string;
  tags?: string[];
  audioDeviceName?: string;
  videoDeviceName?: string;
  displayDeviceName?: string;
  isWheelChairAccessible?: boolean;
  bookingType?: string;
  building?: string;
  floorNumber?: number;
}

// 组实体
export interface Group extends GraphEntity {
  displayName: string;
  description?: string;
  mail?: string;
  mailNickname?: string;
  mailEnabled?: boolean;
  securityEnabled?: boolean;
  groupTypes?: string[];
  visibility?: string;
  createdDateTime?: string;
  renewedDateTime?: string;
  expirationDateTime?: string;
  isArchived?: boolean;
  membershipRule?: string;
  membershipRuleProcessingState?: string;
}

// 联系人实体
export interface Contact extends GraphEntity {
  displayName: string;
  givenName?: string;
  surname?: string;
  emailAddresses?: Array<{
    name?: string;
    address?: string;
  }>;
  businessPhones?: string[];
  mobilePhone?: string;
  homePhones?: string[];
  jobTitle?: string;
  companyName?: string;
  department?: string;
  officeLocation?: string;
  businessAddress?: {
    street?: string;
    city?: string;
    state?: string;
    countryOrRegion?: string;
    postalCode?: string;
  };
  homeAddress?: {
    street?: string;
    city?: string;
    state?: string;
    countryOrRegion?: string;
    postalCode?: string;
  };
  birthday?: string;
  personalNotes?: string;
}

// 邮件实体
export interface Message extends GraphEntity {
  subject: string;
  body?: {
    contentType: string;
    content: string;
  };
  from?: {
    emailAddress: {
      name?: string;
      address?: string;
    };
  };
  toRecipients?: Array<{
    emailAddress: {
      name?: string;
      address?: string;
    };
  }>;
  ccRecipients?: Array<{
    emailAddress: {
      name?: string;
      address?: string;
    };
  }>;
  bccRecipients?: Array<{
    emailAddress: {
      name?: string;
      address?: string;
    };
  }>;
  replyTo?: Array<{
    emailAddress: {
      name?: string;
      address?: string;
    };
  }>;
  sentDateTime?: string;
  receivedDateTime?: string;
  hasAttachments?: boolean;
  importance?: string;
  isRead?: boolean;
  isDraft?: boolean;
  conversationId?: string;
  internetMessageId?: string;
  webLink?: string;
}

// 驱动器项实体
export interface DriveItem extends GraphEntity {
  name: string;
  size?: number;
  webUrl?: string;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  parentReference?: {
    driveId?: string;
    driveType?: string;
    id?: string;
    path?: string;
  };
  file?: {
    mimeType?: string;
    hashes?: {
      sha1Hash?: string;
      quickXorHash?: string;
    };
  };
  folder?: {
    childCount?: number;
  };
  fileSystemInfo?: {
    createdDateTime?: string;
    lastModifiedDateTime?: string;
  };
}

// 权限实体
export interface Permission extends GraphEntity {
  roles?: string[];
  link?: {
    type?: string;
    scope?: string;
    webUrl?: string;
    webHtml?: string;
  };
  grantedTo?: {
    user?: {
      id?: string;
      displayName?: string;
    };
    application?: {
      id?: string;
      displayName?: string;
    };
  };
  grantedToIdentities?: Array<{
    user?: {
      id?: string;
      displayName?: string;
    };
  }>;
  invitation?: {
    email?: string;
    signInRequired?: boolean;
  };
  inheritedFrom?: {
    driveId?: string;
    id?: string;
    path?: string;
  };
}

// 站点实体
export interface Site extends GraphEntity {
  displayName: string;
  name?: string;
  webUrl?: string;
  description?: string;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  siteCollection?: {
    hostname?: string;
  };
}

// 应用程序实体
export interface Application extends GraphEntity {
  displayName: string;
  appId?: string;
  signInAudience?: string;
  createdDateTime?: string;
  description?: string;
  notes?: string;
  publisherDomain?: string;
  identifierUris?: string[];
  tags?: string[];
}

// 服务主体实体
export interface ServicePrincipal extends GraphEntity {
  displayName: string;
  appId?: string;
  servicePrincipalType?: string;
  accountEnabled?: boolean;
  appRoleAssignmentRequired?: boolean;
  publisherName?: string;
  replyUrls?: string[];
  tags?: string[];
}

// Teams 团队实体
export interface Team extends GraphEntity {
  displayName: string;
  description?: string;
  internalId?: string;
  classification?: string;
  visibility?: string;
  webUrl?: string;
  isArchived?: boolean;
  isMembershipLimitedToOwners?: boolean;
  memberSettings?: {
    allowCreateUpdateChannels?: boolean;
    allowDeleteChannels?: boolean;
    allowAddRemoveApps?: boolean;
    allowCreateUpdateRemoveTabs?: boolean;
    allowCreateUpdateRemoveConnectors?: boolean;
  };
  guestSettings?: {
    allowCreateUpdateChannels?: boolean;
    allowDeleteChannels?: boolean;
  };
  messagingSettings?: {
    allowUserEditMessages?: boolean;
    allowUserDeleteMessages?: boolean;
    allowOwnerDeleteMessages?: boolean;
    allowTeamMentions?: boolean;
    allowChannelMentions?: boolean;
  };
  funSettings?: {
    allowGiphy?: boolean;
    giphyContentRating?: string;
    allowStickersAndMemes?: boolean;
    allowCustomMemes?: boolean;
  };
  discoverySettings?: {
    showInTeamsSearchAndSuggestions?: boolean;
  };
}

// Teams 频道实体
export interface Channel extends GraphEntity {
  displayName: string;
  description?: string;
  email?: string;
  webUrl?: string;
  membershipType?: string;
  isFavoriteByDefault?: boolean;
  createdDateTime?: string;
}

// Teams 聊天实体
export interface Chat extends GraphEntity {
  topic?: string;
  chatType?: string;
  createdDateTime?: string;
  lastUpdatedDateTime?: string;
  webUrl?: string;
}

// Planner 计划实体
export interface Plan extends GraphEntity {
  title: string;
  owner?: string;
  createdDateTime?: string;
  createdBy?: {
    user?: {
      id?: string;
      displayName?: string;
    };
  };
}

// Planner 任务实体
export interface Task extends GraphEntity {
  planId: string;
  bucketId?: string;
  title: string;
  orderHint?: string;
  assigneePriority?: string;
  percentComplete?: number;
  startDateTime?: string;
  createdDateTime?: string;
  dueDateTime?: string;
  hasDescription?: boolean;
  previewType?: string;
  completedDateTime?: string;
  completedBy?: {
    user?: {
      id?: string;
      displayName?: string;
    };
  };
  referenceCount?: number;
  checklistItemCount?: number;
  activeChecklistItemCount?: number;
  conversationThreadId?: string;
}

// 订阅实体（Webhooks）
export interface Subscription extends GraphEntity {
  resource: string;
  changeType: string;
  clientState?: string;
  notificationUrl: string;
  expirationDateTime: string;
  creatorId?: string;
  latestSupportedTlsVersion?: string;
  encryptionCertificate?: string;
  encryptionCertificateId?: string;
  includeResourceData?: boolean;
  lifecycleNotificationUrl?: string;
}

// Delta 查询结果
export interface DeltaCollection<T> {
  meta: {
    deltaLink?: string;
    nextLink?: string;
    count?: number;
  };
  data: T[];
}

// 搜索请求
export interface SearchRequest {
  entityTypes: string[];
  query: {
    queryString: string;
  };
  from?: number;
  size?: number;
  fields?: string[];
  sortProperties?: Array<{
    name: string;
    isDescending?: boolean;
  }>;
  enableTopResults?: boolean;
  aggregations?: Array<{
    field: string;
    size?: number;
    bucketDefinition?: {
      sortBy?: string;
      isDescending?: boolean;
      minimumCount?: number;
    };
  }>;
}

// 搜索响应
export interface SearchResponse {
  searchTerms?: string[];
  hitsContainers: Array<{
    total?: number;
    moreResultsAvailable?: boolean;
    hits?: Array<{
      hitId?: string;
      rank?: number;
      summary?: string;
      resource?: unknown;
    }>;
    aggregations?: Array<{
      field?: string;
      buckets?: Array<{
        key?: string;
        count?: number;
        aggregationFilterToken?: string;
      }>;
    }>;
  }>;
}

// ==================== 目录和身份管理 ====================

// 目录角色实体
export interface DirectoryRole extends GraphEntity {
  displayName: string;
  description?: string;
  roleTemplateId?: string;
}

// 目录角色模板实体
export interface DirectoryRoleTemplate extends GraphEntity {
  displayName: string;
  description?: string;
}

// 管理单元实体
export interface AdministrativeUnit extends GraphEntity {
  displayName: string;
  description?: string;
  visibility?: string;
  membershipType?: string;
  membershipRule?: string;
  membershipRuleProcessingState?: string;
}

// 许可证详情实体
export interface LicenseDetails extends GraphEntity {
  servicePlans: Array<{
    servicePlanId?: string;
    servicePlanName?: string;
    provisioningStatus?: string;
    appliesTo?: string;
  }>;
  skuId?: string;
  skuPartNumber?: string;
}

// 订阅的 SKU 实体
export interface SubscribedSku extends GraphEntity {
  skuId?: string;
  skuPartNumber?: string;
  appliesTo?: string;
  capabilityStatus?: string;
  consumedUnits?: number;
  prepaidUnits?: {
    enabled?: number;
    suspended?: number;
    warning?: number;
  };
}

// 邀请实体
export interface Invitation extends GraphEntity {
  invitedUserDisplayName?: string;
  invitedUserEmailAddress: string;
  invitedUserType?: string;
  inviteRedirectUrl: string;
  inviteRedeemUrl?: string;
  sendInvitationMessage?: boolean;
  status?: string;
  invitedUser?: User;
}

// ==================== 日历 ====================

// 日历实体
export interface Calendar extends GraphEntity {
  name: string;
  color?: string;
  changeKey?: string;
  canShare?: boolean;
  canViewPrivateItems?: boolean;
  canEdit?: boolean;
  owner?: {
    name?: string;
    address?: string;
  };
  isDefaultCalendar?: boolean;
  isRemovable?: boolean;
  isTallyingResponses?: boolean;
}

// 日历组实体
export interface CalendarGroup extends GraphEntity {
  name: string;
  classId?: string;
  changeKey?: string;
}

// ==================== 邮件增强 ====================

// 邮件文件夹实体
export interface MailFolder extends GraphEntity {
  displayName: string;
  parentFolderId?: string;
  childFolderCount?: number;
  unreadItemCount?: number;
  totalItemCount?: number;
  isHidden?: boolean;
}

// 邮件规则实体
export interface MessageRule extends GraphEntity {
  displayName: string;
  sequence?: number;
  isEnabled?: boolean;
  conditions?: {
    senderContains?: string[];
    subjectContains?: string[];
    bodyContains?: string[];
    importance?: string;
    fromAddresses?: Array<{
      name?: string;
      address?: string;
    }>;
  };
  actions?: {
    moveToFolder?: string;
    copyToFolder?: string;
    delete?: boolean;
    markAsRead?: boolean;
    markImportance?: string;
    forwardTo?: Array<{
      emailAddress?: {
        name?: string;
        address?: string;
      };
    }>;
  };
}

// 邮箱设置实体
export interface MailboxSettings {
  automaticRepliesSetting?: {
    status?: string;
    externalAudience?: string;
    internalReplyMessage?: string;
    externalReplyMessage?: string;
    scheduledStartDateTime?: {
      dateTime?: string;
      timeZone?: string;
    };
    scheduledEndDateTime?: {
      dateTime?: string;
      timeZone?: string;
    };
  };
  archiveFolder?: string;
  timeZone?: string;
  delegateMeetingMessageDeliveryOptions?: string;
  dateFormat?: string;
  timeFormat?: string;
  language?: {
    locale?: string;
    displayName?: string;
  };
  workingHours?: {
    daysOfWeek?: string[];
    startTime?: string;
    endTime?: string;
    timeZone?: {
      name?: string;
    };
  };
}

// 重点收件箱分类实体
export interface InferenceClassification extends GraphEntity {
  overrides?: Array<{
    id?: string;
    classifyAs?: string;
    senderEmailAddress?: {
      name?: string;
      address?: string;
    };
  }>;
}

// 邮件附件基类
export interface Attachment extends GraphEntity {
  name: string;
  contentType?: string;
  size?: number;
  isInline?: boolean;
  lastModifiedDateTime?: string;
}

// 文件附件实体
export interface FileAttachment extends Attachment {
  contentBytes?: string;
  contentId?: string;
  contentLocation?: string;
}

// 项目附件实体
export interface ItemAttachment extends Attachment {
  item?: Message | Event;
}

// 引用附件实体
export interface ReferenceAttachment extends Attachment {
  sourceUrl?: string;
  providerType?: string;
  thumbnailUrl?: string;
  previewUrl?: string;
  permission?: string;
  isFolder?: boolean;
}

// ==================== 联系人 ====================

// 联系人文件夹实体
export interface ContactFolder extends GraphEntity {
  displayName: string;
  parentFolderId?: string;
}

// 组织联系人实体
export interface OrgContact extends GraphEntity {
  displayName: string;
  givenName?: string;
  surname?: string;
  mail?: string;
  mailNickname?: string;
  jobTitle?: string;
  companyName?: string;
  department?: string;
  officeLocation?: string;
  businessPhones?: string[];
  mobilePhone?: string;
  addresses?: Array<{
    city?: string;
    countryOrRegion?: string;
    postalCode?: string;
    state?: string;
    street?: string;
  }>;
}

// ==================== 文件存储 ====================

// 驱动器实体
export interface Drive extends GraphEntity {
  name?: string;
  driveType?: string;
  owner?: {
    user?: {
      id?: string;
      displayName?: string;
    };
    group?: {
      id?: string;
      displayName?: string;
    };
  };
  quota?: {
    total?: number;
    used?: number;
    remaining?: number;
    deleted?: number;
    state?: string;
  };
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  webUrl?: string;
}

// ==================== OneNote ====================

// OneNote 笔记本实体
export interface Notebook extends GraphEntity {
  displayName: string;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  isDefault?: boolean;
  isShared?: boolean;
  sectionGroupsUrl?: string;
  sectionsUrl?: string;
  links?: {
    oneNoteClientUrl?: {
      href?: string;
    };
    oneNoteWebUrl?: {
      href?: string;
    };
  };
  createdBy?: {
    user?: {
      id?: string;
      displayName?: string;
    };
  };
}

// OneNote 分区实体
export interface Section extends GraphEntity {
  displayName: string;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  isDefault?: boolean;
  pagesUrl?: string;
  links?: {
    oneNoteClientUrl?: {
      href?: string;
    };
    oneNoteWebUrl?: {
      href?: string;
    };
  };
  parentNotebook?: {
    id?: string;
    displayName?: string;
  };
}

// OneNote 页面实体
export interface OneNotePage extends GraphEntity {
  title?: string;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  level?: number;
  order?: number;
  contentUrl?: string;
  content?: string;
  links?: {
    oneNoteClientUrl?: {
      href?: string;
    };
    oneNoteWebUrl?: {
      href?: string;
    };
  };
  parentSection?: {
    id?: string;
    displayName?: string;
  };
}

// ==================== Planner ====================

// Planner 存储桶实体
export interface Bucket extends GraphEntity {
  name: string;
  planId: string;
  orderHint?: string;
}

// ==================== To Do ====================

// To Do 任务列表实体
export interface TodoTaskList extends GraphEntity {
  displayName: string;
  isOwner?: boolean;
  isShared?: boolean;
  wellknownListName?: string;
}

// To Do 任务实体
export interface TodoTask extends GraphEntity {
  title: string;
  body?: {
    content?: string;
    contentType?: string;
  };
  bodyLastModifiedDateTime?: string;
  completedDateTime?: {
    dateTime?: string;
    timeZone?: string;
  };
  dueDateTime?: {
    dateTime?: string;
    timeZone?: string;
  };
  importance?: string;
  isReminderOn?: boolean;
  reminderDateTime?: {
    dateTime?: string;
    timeZone?: string;
  };
  status?: string;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
}

// ==================== SharePoint ====================

// SharePoint 列表实体
export interface List extends GraphEntity {
  displayName: string;
  name?: string;
  description?: string;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  list?: {
    hidden?: boolean;
    template?: string;
  };
  webUrl?: string;
}

// SharePoint 列表项实体
export interface ListItem extends GraphEntity {
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  contentType?: {
    id?: string;
    name?: string;
  };
  fields?: Record<string, unknown>;
  webUrl?: string;
}

// SharePoint 列实体
export interface ColumnDefinition extends GraphEntity {
  name: string;
  displayName?: string;
  description?: string;
  columnGroup?: string;
  enforceUniqueValues?: boolean;
  hidden?: boolean;
  indexed?: boolean;
  readOnly?: boolean;
  required?: boolean;
  text?: {
    allowMultipleLines?: boolean;
    maxLength?: number;
  };
  boolean?: Record<string, unknown>;
  calculated?: {
    format?: string;
    formula?: string;
  };
  choice?: {
    allowTextEntry?: boolean;
    choices?: string[];
    displayAs?: string;
  };
  currency?: {
    locale?: string;
  };
  dateTime?: {
    displayAs?: string;
    format?: string;
  };
  lookup?: {
    allowMultipleValues?: boolean;
    columnName?: string;
    listId?: string;
    primaryLookupColumnId?: string;
  };
  number?: {
    decimalPlaces?: string;
    displayAs?: string;
    maximum?: number;
    minimum?: number;
  };
  personOrGroup?: {
    allowMultipleSelection?: boolean;
    displayAs?: string;
    chooseFromType?: string;
  };
}

// ==================== Teams 增强 ====================

// 频道消息实体
export interface ChatMessage extends GraphEntity {
  messageType?: string;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  deletedDateTime?: string;
  subject?: string;
  body?: {
    content?: string;
    contentType?: string;
  };
  summary?: string;
  from?: {
    user?: {
      id?: string;
      displayName?: string;
    };
    application?: {
      id?: string;
      displayName?: string;
    };
  };
  attachments?: Array<{
    id?: string;
    contentType?: string;
    contentUrl?: string;
    content?: string;
    name?: string;
    thumbnailUrl?: string;
  }>;
  mentions?: Array<{
    id?: string;
    mentionText?: string;
    mentioned?: {
      user?: {
        id?: string;
        displayName?: string;
      };
    };
  }>;
  importance?: string;
  replyToId?: string;
  webUrl?: string;
}

// 团队成员实体
export interface ConversationMember extends GraphEntity {
  displayName?: string;
  roles?: string[];
  visibleHistoryStartDateTime?: string;
  userId?: string;
  email?: string;
}

// 团队应用实体
export interface TeamsAppInstallation extends GraphEntity {
  teamsApp?: {
    id?: string;
    displayName?: string;
    distributionMethod?: string;
  };
  teamsAppDefinition?: {
    id?: string;
    displayName?: string;
    version?: string;
  };
}

// 团队标签实体
export interface TeamworkTag extends GraphEntity {
  displayName: string;
  description?: string;
  memberCount?: number;
  tagType?: string;
  teamId?: string;
}

// ==================== 报告和分析 ====================

// 审计日志实体
export interface DirectoryAudit extends GraphEntity {
  activityDateTime: string;
  activityDisplayName: string;
  category: string;
  correlationId?: string;
  result?: string;
  resultReason?: string;
  loggedByService?: string;
  operationType?: string;
  initiatedBy?: {
    user?: {
      id?: string;
      displayName?: string;
      userPrincipalName?: string;
    };
    app?: {
      appId?: string;
      displayName?: string;
      servicePrincipalId?: string;
    };
  };
  targetResources?: Array<{
    id?: string;
    displayName?: string;
    type?: string;
    userPrincipalName?: string;
    modifiedProperties?: Array<{
      displayName?: string;
      oldValue?: string;
      newValue?: string;
    }>;
  }>;
}

// 登录日志实体
export interface SignIn extends GraphEntity {
  createdDateTime: string;
  userDisplayName?: string;
  userPrincipalName?: string;
  userId?: string;
  appId?: string;
  appDisplayName?: string;
  ipAddress?: string;
  clientAppUsed?: string;
  correlationId?: string;
  conditionalAccessStatus?: string;
  isInteractive?: boolean;
  riskDetail?: string;
  riskLevelAggregated?: string;
  riskLevelDuringSignIn?: string;
  riskState?: string;
  status?: {
    errorCode?: number;
    failureReason?: string;
    additionalDetails?: string;
  };
  deviceDetail?: {
    deviceId?: string;
    displayName?: string;
    operatingSystem?: string;
    browser?: string;
    isCompliant?: boolean;
    isManaged?: boolean;
    trustType?: string;
  };
  location?: {
    city?: string;
    state?: string;
    countryOrRegion?: string;
    geoCoordinates?: {
      latitude?: number;
      longitude?: number;
    };
  };
}

// 洞察分析实体
export interface Trending extends GraphEntity {
  weight?: number;
  resourceVisualization?: {
    title?: string;
    type?: string;
    mediaType?: string;
    previewImageUrl?: string;
    previewText?: string;
    containerWebUrl?: string;
    containerDisplayName?: string;
    containerType?: string;
  };
  resourceReference?: {
    id?: string;
    webUrl?: string;
    type?: string;
  };
  lastModifiedDateTime?: string;
}

// 使用过的资源实体
export interface UsedInsight extends GraphEntity {
  lastUsed?: {
    lastAccessedDateTime?: string;
    lastModifiedDateTime?: string;
  };
  resourceVisualization?: {
    title?: string;
    type?: string;
    mediaType?: string;
    previewImageUrl?: string;
    previewText?: string;
    containerWebUrl?: string;
    containerDisplayName?: string;
    containerType?: string;
  };
  resourceReference?: {
    id?: string;
    webUrl?: string;
    type?: string;
  };
}

// 共享资源实体
export interface SharedInsight extends GraphEntity {
  lastShared?: {
    sharedDateTime?: string;
    sharingSubject?: string;
    sharingType?: string;
    sharedBy?: {
      displayName?: string;
      id?: string;
      address?: string;
    };
  };
  resourceVisualization?: {
    title?: string;
    type?: string;
    mediaType?: string;
    previewImageUrl?: string;
    previewText?: string;
    containerWebUrl?: string;
    containerDisplayName?: string;
    containerType?: string;
  };
  resourceReference?: {
    id?: string;
    webUrl?: string;
    type?: string;
  };
}

// 人员实体
export interface Person extends GraphEntity {
  displayName: string;
  givenName?: string;
  surname?: string;
  birthday?: string;
  personNotes?: string;
  isFavorite?: boolean;
  jobTitle?: string;
  companyName?: string;
  yomiCompany?: string;
  department?: string;
  officeLocation?: string;
  profession?: string;
  userPrincipalName?: string;
  imAddress?: string;
  scoredEmailAddresses?: Array<{
    address?: string;
    relevanceScore?: number;
    selectionLikelihood?: string;
  }>;
  phones?: Array<{
    type?: string;
    number?: string;
  }>;
  personType?: {
    class?: string;
    subclass?: string;
  };
}

// ==================== 安全与合规 ====================

// 安全警报实体
export interface Alert extends GraphEntity {
  title: string;
  description?: string;
  severity?: string;
  status?: string;
  category?: string;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  eventDateTime?: string;
  assignedTo?: string;
  azureSubscriptionId?: string;
  azureTenantId?: string;
  confidence?: number;
  detectionIds?: string[];
  feedback?: string;
  fileStates?: Array<{
    name?: string;
    path?: string;
    riskScore?: string;
    fileHash?: {
      hashType?: string;
      hashValue?: string;
    };
  }>;
  hostStates?: Array<{
    fqdn?: string;
    isAzureAdJoined?: boolean;
    isAzureAdRegistered?: boolean;
    isHybridAzureDomainJoined?: boolean;
    netBiosName?: string;
    os?: string;
    privateIpAddress?: string;
    publicIpAddress?: string;
    riskScore?: string;
  }>;
  networkConnections?: Array<{
    applicationName?: string;
    destinationAddress?: string;
    destinationPort?: string;
    direction?: string;
    domainRegisteredDateTime?: string;
    localDnsName?: string;
    natDestinationAddress?: string;
    natDestinationPort?: string;
    natSourceAddress?: string;
    natSourcePort?: string;
    protocol?: string;
    riskScore?: string;
    sourceAddress?: string;
    sourcePort?: string;
    uri?: string;
  }>;
  sourceMaterials?: string[];
  tags?: string[];
  vendorInformation?: {
    provider?: string;
    providerVersion?: string;
    subProvider?: string;
    vendor?: string;
  };
}

// 安全评分实体
export interface SecureScore extends GraphEntity {
  activeUserCount?: number;
  averageComparativeScores?: Array<{
    basis?: string;
    averageScore?: number;
  }>;
  azureTenantId?: string;
  controlScores?: Array<{
    controlName?: string;
    controlCategory?: string;
    score?: number;
    description?: string;
  }>;
  createdDateTime?: string;
  currentScore?: number;
  enabledServices?: string[];
  licensedUserCount?: number;
  maxScore?: number;
  vendorInformation?: {
    provider?: string;
    providerVersion?: string;
    subProvider?: string;
    vendor?: string;
  };
}

// 风险检测实体
export interface RiskDetection extends GraphEntity {
  requestId?: string;
  correlationId?: string;
  riskEventType?: string;
  riskState?: string;
  riskLevel?: string;
  riskDetail?: string;
  source?: string;
  detectionTimingType?: string;
  activity?: string;
  tokenIssuerType?: string;
  ipAddress?: string;
  location?: {
    city?: string;
    state?: string;
    countryOrRegion?: string;
    geoCoordinates?: {
      latitude?: number;
      longitude?: number;
    };
  };
  activityDateTime?: string;
  detectedDateTime?: string;
  lastUpdatedDateTime?: string;
  userId?: string;
  userDisplayName?: string;
  userPrincipalName?: string;
  additionalInfo?: string;
}

// 风险用户实体
export interface RiskyUser extends GraphEntity {
  userDisplayName?: string;
  userPrincipalName?: string;
  riskLevel?: string;
  riskState?: string;
  riskDetail?: string;
  riskLastUpdatedDateTime?: string;
  isDeleted?: boolean;
  isProcessing?: boolean;
}

// ==================== 扩展 ====================

// 架构扩展实体
export interface SchemaExtension extends GraphEntity {
  description?: string;
  targetTypes: string[];
  properties: Array<{
    name: string;
    type: string;
  }>;
  status?: string;
  owner?: string;
}

// 开放扩展实体
export interface Extension extends GraphEntity {
  extensionName: string;
  [key: string]: unknown;
}

// ==================== 其他 ====================

// 目录对象实体
export interface DirectoryObject extends GraphEntity {
  deletedDateTime?: string;
}
