# 🛠️ 贡献指南

感谢你对 `@qingu-x/msgraph-orm-js` 的关注！

## 🚀 快速开始

```bash
# 克隆仓库
git clone https://github.com/qingu-x/msgraph-orm-js.git
cd msgraph-orm-js

# 安装依赖
npm install

# 启动开发模式
npm run build:watch

# 运行测试
npm test
```

## 📋 前置要求

- Node.js >= 18.0.0
- npm >= 9.0.0

## 📁 项目结构

```text
msgraph-orm-js/
├── src/                    # 源代码
│   ├── index.ts            # 主入口
│   ├── client-credentials.ts # 认证
│   ├── types.ts            # 类型定义
│   └── orm/                # ORM 实现
├── lib/                    # 构建产物
│   ├── bundle.esm.js       # ES Module
│   ├── bundle.cjs.js       # CommonJS
│   └── bundle.browser.js   # UMD
├── types/                  # TS 类型声明
├── test/                   # 测试文件
├── docs/                   # 文档
└── package.json            # 包配置
```

## 🏗️ 构建和测试

### 开发

```bash
# 启动监听模式（自动重新构建）
npm run build:watch

# 构建所有格式
npm run build
```

### 测试

```bash
# 运行所有测试
npm test

# 监听模式
npm test -- --watch

# 覆盖率报告
npm test -- --coverage
```

### 代码检查

```bash
# 运行 ESLint
npm run lint

# 自动修复
npm run lint:fix
```

## 📦 模块导入使用

### ES Module (推荐)

```typescript
import { createGraphORM, ClientCredentialsAuthProvider } from '@qingu-x/msgraph-orm-js';
```

### CommonJS

```javascript
const { createGraphORM, ClientCredentialsAuthProvider } = require('@qingu-x/msgraph-orm-js');
```

### UMD (浏览器)

```html
<script src="node_modules/@qingu-x/msgraph-orm-js/lib/bundle.browser.js"></script>
<script>
  const { createGraphORM } = window.MicrosoftGraphORM;
</script>
```

## 📝 代码规范

### TypeScript

- 使用严格模式
- 明确的类型注解
- 避免 `any` 类型

### ESLint

遵循项目的 ESLint 配置：

```bash
# 检查代码
npm run lint

# 自动修复
npm run lint:fix
```

### 代码风格

- 使用 2 空格缩进
- 使用单引号
- 行尾添加分号
- 使用 trailing comma

## 🔖 提交代码

### Commit 规范

使用 [Conventional Commits](https://www.conventionalcommits.org/zh-hans/)：

```text
<类型>[可选作用域]: <描述>

[可选正文]

[可选脚注]
```

**类型：**

- `feat`: 新功能
- `fix`: Bug 修复
- `docs`: 文档更新
- `style`: 代码格式
- `refactor`: 重构
- `perf`: 性能优化
- `test`: 测试
- `build`: 构建系统
- `ci`: CI/CD
- `chore`: 其他修改

**示例：**

```text
feat: 添加 Delta Query 支持
fix: 修复分页查询 Bug
docs: 更新 README 文档
refactor: 重构查询构建器
```

### 使用 Commitizen

```bash
# 安装（可选）
npm install -g commitizen

# 交互式提交
npm run commit
```

### 提交流程

1. Fork 仓库
2. 创建功能分支：`git checkout -b feature/xxx`
3. 提交更改：`git commit -m 'feat: add xxx'`
4. 推送分支：`git push origin feature/xxx`
5. 提交 Pull Request

### PR 规范

- 清晰的标题和描述
- 关联相关 Issue
- 通过所有测试
- 代码审查通过

## 🐛 故障排除

### 构建失败

```bash
# 清理并重新安装
rm -rf node_modules package-lock.json
npm install

# 清理构建产物
rm -rf lib types
npm run build
```

### 类型错误

```bash
# 重新生成类型声明
npm run build
```

### 依赖问题

```bash
# 检查依赖
npm outdated

# 更新依赖
npm update
```

### 测试失败

```bash
# 清理 Jest 缓存
npm test -- --clearCache

# 运行单个测试
npm test -- path/to/test.ts
```

## 📚 相关资源

- [TypeScript 手册](https://www.typescriptlang.org/docs/)
- [ESLint 规则](https://eslint.org/docs/rules/)
- [Conventional Commits](https://www.conventionalcommits.org/zh-hans/)
- [Jest 文档](https://jestjs.io/docs/getting-started)
- [Vite 文档](https://cn.vitejs.dev/)

## 💬 获取帮助

如有问题，请：

1. 查看[文档](../README.md)
2. 查看[文档索引](./DOCS_INDEX.md)
3. 提交 [Issue](https://github.com/qingu-x/msgraph-orm-js/issues)
4. 加入讨论：[GitHub Discussions](https://github.com/qingu-x/msgraph-orm-js/discussions)

---

再次感谢你的贡献！🎉
