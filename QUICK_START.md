# 🚀 快速开始 - 构建和发布

## 📦 日常开发

```bash
# 安装依赖
npm install

# 开发时构建（监听模式）
npm run build:watch

# 运行测试
npm test

# 提交代码（使用规范化提交）
npm run commit
```

## 🏗️ 构建

```bash
# 构建所有格式（ESM + CJS + UMD）
npm run build

# 构建产物位置：
# - lib/bundle.esm.js      (ES Module - 推荐)
# - lib/bundle.cjs.js      (CommonJS - Node.js)
# - lib/bundle.browser.js  (UMD - 浏览器)
# - types/                 (TypeScript 声明文件)
```

## 📝 版本发布

```bash
# 1. 自动发布（推荐）
npm run release              # 自动判断版本号
npm run release:patch        # 0.0.1 → 0.0.2
npm run release:minor        # 0.0.1 → 0.1.0
npm run release:major        # 0.0.1 → 1.0.0

# 2. 推送到 Git
git push --follow-tags origin main

# 3. 发布到 NPM
npm publish --access public
```

## 📋 standard-version 配置

### 配置文件：`.versionrc.json`

- **CHANGELOG 输出**: `CHANGELOG.md`（根目录）+ `docs/CHANGELOG.md`（副本）
- **提交格式**: Angular 规范
- **自动化操作**: 生成 CHANGELOG → 更新版本号 → Git commit → Git tag

### 常用命令

```bash
# 预览发布（不实际执行）
npm run release -- --dry-run

# 首次发布
npm run release -- --first-release

# 预发布版本
npm run release -- --prerelease alpha

# 指定版本号
npm run release -- --release-as 1.2.3
```

## 🔍 验证发布

```bash
# 查看将要发布的文件
npm pack --dry-run

# 本地测试包
npm pack
cd /tmp
npm install /path/to/qingu-x-msgraph-orm-js-0.0.x.tgz
```

## ⚙️ 技术栈

- **构建**: Vite 7.x + TypeScript 5.x
- **格式**: ESM + CJS + UMD
- **类型**: 自动生成 .d.ts 文件
- **版本管理**: standard-version + commitizen

## 📚 更多文档

- [完整发布指南](./RELEASE.md)
- [项目文档](./README.md)
- [API 参考](./docs/DOCS_INDEX.md)
