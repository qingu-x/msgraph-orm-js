# 发布指南

本文档说明如何使用 `standard-version` 进行版本发布和 CHANGELOG 管理。

## 📦 版本发布流程

### 1. 确保代码已提交

```bash
git status  # 检查工作区是否干净
```

### 2. 运行发布命令

#### 自动判断版本号（推荐）

```bash
npm run release
```

`standard-version` 会根据提交信息自动判断版本类型：
- `feat:` 提交 → 增加 **minor** 版本（0.0.1 → 0.1.0）
- `fix:` 提交 → 增加 **patch** 版本（0.0.1 → 0.0.2）
- `BREAKING CHANGE:` → 增加 **major** 版本（0.0.1 → 1.0.0）

#### 手动指定版本类型

```bash
# 修订版本（bug 修复）: 0.0.1 → 0.0.2
npm run release:patch

# 次版本（新功能）: 0.0.1 → 0.1.0
npm run release:minor

# 主版本（破坏性变更）: 0.0.1 → 1.0.0
npm run release:major
```

#### 指定具体版本号

```bash
npm run release -- --release-as 1.2.3
```

#### 预发布版本

```bash
# 创建 alpha 版本: 0.0.1 → 0.0.2-alpha.0
npm run release -- --prerelease alpha

# 创建 beta 版本: 0.0.1 → 0.0.2-beta.0
npm run release -- --prerelease beta
```

### 3. standard-version 自动执行的操作

运行 `npm run release` 后，`standard-version` 会：

1. ✅ 根据 Git 提交历史生成 CHANGELOG.md
2. ✅ 复制 CHANGELOG.md 到 docs/CHANGELOG.md
3. ✅ 更新 package.json 中的版本号
4. ✅ 提交所有更改（commit message: `chore(release): v0.0.x`）
5. ✅ 创建 Git tag（例如 `v0.0.2`）

### 4. 推送到远程仓库

```bash
# 推送代码和标签
git push --follow-tags origin main

# 或者分开推送
git push origin main
git push origin --tags
```

### 5. 发布到 NPM

```bash
# 发布到 NPM（会自动触发 prepublishOnly 脚本进行构建）
npm publish --access public
```

## 📝 CHANGELOG 配置说明

### 配置文件位置

`.versionrc.json` - standard-version 的配置文件

### CHANGELOG 输出位置

- **主文件**: `CHANGELOG.md`（项目根目录）
- **副本**: `docs/CHANGELOG.md`（通过 postchangelog 脚本自动复制）

### 提交类型映射

| 提交类型 | CHANGELOG 标题 | 是否显示 |
|---------|---------------|---------|
| `feat:` | ✨ Features | ✅ 显示 |
| `fix:` | 🐛 Bug Fixes | ✅ 显示 |
| `perf:` | ⚡ Performance Improvements | ✅ 显示 |
| `revert:` | ⏪ Reverts | ✅ 显示 |
| `docs:` | 📝 Documentation | ✅ 显示 |
| `refactor:` | ♻️ Code Refactoring | ✅ 显示 |
| `build:` | 📦 Build System | ✅ 显示 |
| `style:` | 💄 Styles | ❌ 隐藏 |
| `test:` | ✅ Tests | ❌ 隐藏 |
| `chore:` | 🔧 Chore | ❌ 隐藏 |
| `ci:` | 👷 CI | ❌ 隐藏 |

### 如何更改 CHANGELOG 输出位置

如果要将 CHANGELOG 直接生成到 `docs/` 目录，修改 `.versionrc.json`：

```json
{
  "infile": "docs/CHANGELOG.md"
}
```

但**推荐保持当前配置**（根目录 + docs 副本），因为：
- NPM 包通常期望在根目录找到 CHANGELOG.md
- docs/ 目录适合文档站点使用

## ✍️ 提交规范

使用 Commitizen 进行规范提交：

```bash
# 使用交互式提交
npm run commit

# 或使用 git cz
git cz
```

### 提交信息格式

```
<type>(<scope>): <subject>

<body>

<footer>
```

**示例**：

```bash
feat(orm): 添加批量查询功能

- 支持一次查询多个实体
- 优化查询性能
- 添加单元测试

Closes #123
```

## 🔄 完整发布流程示例

```bash
# 1. 开发功能并提交
git add .
npm run commit
# 选择 feat，填写描述

# 2. 发布新版本
npm run release

# 3. 推送到 Git
git push --follow-tags origin main

# 4. 发布到 NPM
npm publish --access public

# 5. 创建 GitHub Release（可选）
# 在 GitHub 仓库页面手动创建，或使用 gh cli
gh release create v0.0.2 --notes "$(cat CHANGELOG.md)"
```

## 🛠️ 高级用法

### 首次发布

```bash
# 生成初始 CHANGELOG（所有历史记录）
npm run release -- --first-release
```

### 跳过某个步骤

```bash
# 跳过 changelog 生成
npm run release -- --skip.changelog

# 跳过 commit
npm run release -- --skip.commit

# 跳过 tag
npm run release -- --skip.tag
```

### Dry Run（预览）

```bash
# 预览发布操作但不实际执行
npm run release -- --dry-run
```

### 查看即将发布的版本

```bash
# 查看下一个版本号
npm run release -- --dry-run | grep "tagging release"
```

## 📊 版本号说明

遵循 [语义化版本 2.0.0](https://semver.org/lang/zh-CN/)：

**格式**: `MAJOR.MINOR.PATCH`

- **MAJOR (主版本号)**: 不兼容的 API 变更
- **MINOR (次版本号)**: 向下兼容的功能性新增
- **PATCH (修订号)**: 向下兼容的问题修正

**示例**：
- `1.0.0` → `2.0.0`: 重大破坏性变更
- `1.0.0` → `1.1.0`: 新增功能，保持兼容
- `1.0.0` → `1.0.1`: Bug 修复

## 🔍 故障排除

### 问题：standard-version 报错 "No commits since last release"

**解决**：确保有新的提交，或使用 `--first-release`

```bash
npm run release -- --first-release
```

### 问题：CHANGELOG 没有复制到 docs/

**解决**：手动执行复制脚本

```bash
cp CHANGELOG.md docs/CHANGELOG.md
git add docs/CHANGELOG.md
git commit -m "docs: update changelog"
```

### 问题：想要撤销发布

```bash
# 删除本地 tag
git tag -d v0.0.2

# 删除远程 tag
git push origin :refs/tags/v0.0.2

# 回退提交
git reset --hard HEAD~1

# 如果已经推送到 NPM，使用 deprecate
npm deprecate @qingu-x/msgraph-orm-js@0.0.2 "This version has been deprecated"
```

## 📚 相关链接

- [standard-version 文档](https://github.com/conventional-changelog/standard-version)
- [Conventional Commits 规范](https://www.conventionalcommits.org/zh-hans/)
- [语义化版本规范](https://semver.org/lang/zh-CN/)
- [Commitizen 文档](https://github.com/commitizen/cz-cli)

