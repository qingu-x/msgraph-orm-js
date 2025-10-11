# 文档索引

本文档提供项目所有文档的快速导航和说明。

## 📚 文档结构

```bash
msgraph-orm-js/
├── README.md                  # 📖 项目主文档（从这里开始）
├── CHANGELOG.md               # 📝 版本更新日志
└── docs/
    ├── DOCS_INDEX.md             # 📑 本文档（文档索引）
    ├── GRAPH_ORM_GUIDE.md        # 🎓 ORM 使用教程
    ├── EXAMPLES.md               # 💡 完整代码示例
    ├── API_COVERAGE.md           # ✅ API 覆盖率报告
    └── GRAPH_API_REFERENCE.md    # 📘 API 参考手册
```

## 🚀 快速导航

### 新手入门

1. **[README.md](../README.md)** - 从这里开始！
   - 项目简介和特性
   - 快速安装和配置
   - 基础用法示例
   - 常见问题

2. **[GRAPH_ORM_GUIDE.md](./GRAPH_ORM_GUIDE.md)** - ORM 使用教程
   - ORM 概念和设计思想
   - 查询构建器详解
   - 实体管理器使用
   - 高级功能说明

3. **[EXAMPLES.md](./EXAMPLES.md)** - 完整代码示例
   - 各个功能模块的实际代码
   - 常见场景的完整示例
   - 最佳实践示范
   - 错误处理示例

### 参考文档

1. **[API_COVERAGE.md](./API_COVERAGE.md)** - API 覆盖率报告
   - 所有已实现的 API 端点列表
   - 每个 API 的权限要求
   - 国家云支持情况
   - 未实现的 API 说明

2. **[GRAPH_API_REFERENCE.md](./GRAPH_API_REFERENCE.md)** - API 参考手册
   - API 端点分类
   - 功能模块说明
   - 实现状态和优先级
   - 国家云限制说明

3. **[CHANGELOG.md](./CHANGELOG.md)** - 版本更新日志
   - 版本历史
   - 新增功能列表
   - Bug 修复记录
   - 破坏性变更说明

## 📖 按场景查找文档

### 我想快速开始使用

→ [README.md](../README.md) > 快速开始部分

### 我想了解如何查询数据

→ [GRAPH_ORM_GUIDE.md](./GRAPH_ORM_GUIDE.md) > 查询构建器章节
→ [EXAMPLES.md](./EXAMPLES.md) > 高级查询部分

### 我想查看完整的代码示例

→ [EXAMPLES.md](./EXAMPLES.md)

### 我想知道某个功能是否支持

→ [API_COVERAGE.md](./API_COVERAGE.md) > 按功能模块查找

### 我想了解权限要求

→ [API_COVERAGE.md](./API_COVERAGE.md) > 每个 API 都标注了权限
→ [README.md](../README.md) > 权限说明章节

### 我想在中国版/美国政府版部署

→ [GRAPH_API_REFERENCE.md](./GRAPH_API_REFERENCE.md) > 国家云支持说明
→ [API_COVERAGE.md](./API_COVERAGE.md) > 国家云支持列

### 我想查看版本更新内容

→ [CHANGELOG.md](./CHANGELOG.md)

## 🎯 按功能模块查找

### 用户和身份管理

- **教程**: [GRAPH_ORM_GUIDE.md](./GRAPH_ORM_GUIDE.md) > 实体管理章节
- **示例**: [EXAMPLES.md](./EXAMPLES.md) > 用户管理
- **API 列表**: [API_COVERAGE.md](./API_COVERAGE.md) > 用户和身份管理

### 日历管理

- **示例**: [EXAMPLES.md](./EXAMPLES.md) > 日历管理
- **API 列表**: [API_COVERAGE.md](./API_COVERAGE.md) > 日历

### 邮件管理

- **示例**: [EXAMPLES.md](./EXAMPLES.md) > 邮件管理
- **API 列表**: [API_COVERAGE.md](./API_COVERAGE.md) > 邮件

### 文件存储 (OneDrive)

- **示例**: [EXAMPLES.md](./EXAMPLES.md) > 文件管理
- **API 列表**: [API_COVERAGE.md](./API_COVERAGE.md) > 文件存储

### Teams 协作

- **示例**: [EXAMPLES.md](./EXAMPLES.md) > Teams 协作
- **API 列表**: [API_COVERAGE.md](./API_COVERAGE.md) > Teams

### SharePoint

- **示例**: [EXAMPLES.md](./EXAMPLES.md) > SharePoint
- **API 列表**: [API_COVERAGE.md](./API_COVERAGE.md) > SharePoint

### 报告和分析

- **示例**: [EXAMPLES.md](./EXAMPLES.md) > 报告和分析
- **API 列表**: [API_COVERAGE.md](./API_COVERAGE.md) > 报告和分析

### 安全与合规

- **示例**: [EXAMPLES.md](./EXAMPLES.md) > 高级功能（安全部分）
- **API 列表**: [API_COVERAGE.md](./API_COVERAGE.md) > 安全与合规

## 🔍 按问题类型查找

### 如何认证？

→ [README.md](../README.md) > 快速开始
→ [EXAMPLES.md](./EXAMPLES.md) > 初始化

### 如何查询数据？

→ [GRAPH_ORM_GUIDE.md](./GRAPH_ORM_GUIDE.md) > 查询构建器
→ [EXAMPLES.md](./EXAMPLES.md) > 高级查询

### 如何创建/更新/删除？

→ [GRAPH_ORM_GUIDE.md](./GRAPH_ORM_GUIDE.md) > 实体管理器
→ [EXAMPLES.md](./EXAMPLES.md) > 用户管理（CRUD 操作）

### 如何处理错误？

→ [EXAMPLES.md](./EXAMPLES.md) > 错误处理
→ [README.md](../README.md) > 错误处理

### 如何分页？

→ [GRAPH_ORM_GUIDE.md](./GRAPH_ORM_GUIDE.md) > 查询构建器 > 分页
→ [EXAMPLES.md](./EXAMPLES.md) > 高级查询 > 分页查询

### 如何批处理请求？

→ [EXAMPLES.md](./EXAMPLES.md) > 高级查询 > 批处理请求

### 如何跟踪变化（Delta Query）？

→ [EXAMPLES.md](./EXAMPLES.md) > 高级查询 > Delta Query

### 如何订阅 Webhook？

→ [EXAMPLES.md](./EXAMPLES.md) > 高级查询 > Webhooks 订阅

## 📊 文档详细说明

### 1. README.md

**适合**: 所有用户  
**内容**:

- 项目简介和核心特性
- 安装和快速开始
- 基础功能演示
- 国家云配置
- API 覆盖范围概览
- 最佳实践
- 错误处理基础

**何时阅读**: 首次接触项目时必读

### 2. GRAPH_ORM_GUIDE.md

**适合**: 需要深入了解 ORM 机制的开发者  
**内容**:

- ORM 设计理念
- 查询构建器完整 API
- 实体管理器详解
- 高级功能使用方法
- 类型系统说明

**何时阅读**: 需要了解 ORM 工作原理或使用高级功能时

### 3. EXAMPLES.md

**适合**: 所有开发者  
**内容**:

- 完整的、可运行的代码示例
- 覆盖所有主要功能模块
- 常见场景的解决方案
- 最佳实践示范

**何时阅读**:

- 需要参考代码示例时
- 不确定如何实现某个功能时
- 想学习最佳实践时

### 4. API_COVERAGE.md

**适合**: 需要确认 API 支持情况的开发者  
**内容**:

- 所有已实现的 API 完整列表
- 每个 API 的 ORM 方法名
- 权限要求
- 国家云支持情况
- 未实现的 API 列表和原因

**何时阅读**:

- 需要确认某个 API 是否支持时
- 规划项目功能时
- 需要了解权限要求时

### 5. GRAPH_API_REFERENCE.md

**适合**: 架构师和高级开发者  
**内容**:

- API 端点分类和组织结构
- 功能模块详细说明
- 实现优先级
- 国家云限制详解

**何时阅读**:

- 进行架构设计时
- 需要全面了解 API 结构时
- 规划国家云部署时

### 6. CHANGELOG.md

**适合**: 所有用户  
**内容**:

- 版本历史
- 每个版本的新增功能
- Bug 修复
- 破坏性变更

**何时阅读**:

- 升级版本前
- 想了解新功能时
- 排查版本相关问题时

## 💡 使用建议

### 第一次使用？

```bash
1. README.md（了解项目）
   ↓
2. README.md > 快速开始（配置和运行）
   ↓
3. EXAMPLES.md（参考示例代码）
```

### 开发新功能？

```bash
1. API_COVERAGE.md（确认 API 支持）
   ↓
2. EXAMPLES.md（查找相关示例）
   ↓
3. GRAPH_ORM_GUIDE.md（深入了解机制）
```

### 排查问题？

```bash
1. EXAMPLES.md > 错误处理（基础排查）
   ↓
2. API_COVERAGE.md（确认 API 和权限）
   ↓
3. GitHub Issues（查找已知问题）
```

### 项目设计？

```bash
1. GRAPH_API_REFERENCE.md（了解 API 结构）
   ↓
2. API_COVERAGE.md（确认功能覆盖）
   ↓
3. README.md > 国家云支持（考虑部署环境）
```

## 🔗 外部资源

- [Microsoft Graph 官方文档](https://learn.microsoft.com/graph/)
- [Microsoft Graph Explorer](https://developer.microsoft.com/graph/graph-explorer)
- [Microsoft Graph SDK](https://github.com/microsoftgraph/msgraph-sdk-javascript)
- [国家云部署](https://learn.microsoft.com/graph/deployments)
- [权限参考](https://learn.microsoft.com/graph/permissions-reference)

## 📝 文档维护

### 文档更新频率

- **README.md**: 每个主要版本
- **CHANGELOG.md**: 每次发布
- **EXAMPLES.md**: 新增功能时
- **API_COVERAGE.md**: 新增 API 时
- **GRAPH_ORM_GUIDE.md**: API 变更时
- **GRAPH_API_REFERENCE.md**: 架构变更时

### 贡献文档

如果您发现文档问题或想要改进文档：

1. 提交 Issue 说明问题
2. 或直接提交 Pull Request
3. 遵循现有文档风格

## 🎓 学习路径

### 初级开发者

```bash
第1周: README.md + EXAMPLES.md（基础操作）
第2周: EXAMPLES.md（各模块深入）
第3周: GRAPH_ORM_GUIDE.md（理解机制）
```

### 中级开发者

```bash
第1天: README.md（快速了解）
第2天: EXAMPLES.md（核心功能）
第3天: API_COVERAGE.md（规划功能）
```

### 高级开发者

```bash
1小时: 浏览所有文档
重点: GRAPH_API_REFERENCE.md + API_COVERAGE.md
```

---

**提示**: 善用 Ctrl+F (Cmd+F) 在文档中搜索关键词！

**更新时间**: 2025-10-10
