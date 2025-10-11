# 📦 发布指南

如何发布新版本到 NPM。

## 🚀 快速发布（4步）

```bash
# 1. 确保工作区干净
git status

# 2. 运行发布脚本（自动判断版本）
npm run release

# 3. 推送到 Git
git push --follow-tags origin main

# 4. 发布到 NPM
npm publish --access public
```

## 📋 详细步骤

### 1. 准备发布

确保：

- ✅ 所有代码已提交
- ✅ 工作区干净
- ✅ 测试通过：`npm test`
- ✅ 构建成功：`npm run build`

### 2. 选择版本类型

#### 自动判断（推荐）

```bash
npm run release
```

根据提交信息自动判断：

- `feat:` → minor（0.0.1 → 0.1.0）
- `fix:` → patch（0.0.1 → 0.0.2）
- `BREAKING CHANGE:` → major（0.0.1 → 1.0.0）

#### 手动指定

```bash
npm run release:patch    # 修订版本（Bug 修复）
npm run release:minor    # 次版本（新功能）
npm run release:major    # 主版本（破坏性变更）
```

#### 预发布版本

```bash
npm run release -- --prerelease alpha    # Alpha 版本
npm run release -- --prerelease beta     # Beta 版本
npm run release -- --prerelease rc       # RC 版本
```

### 3. 自动操作

`standard-version` 会自动：

1. 分析 Git 提交历史
2. 生成/更新 CHANGELOG.md
3. 更新 package.json 版本号
4. Git commit（`chore(release): vX.X.X`）
5. 创建 Git tag（例如：`v0.0.5`）

### 4. 推送到 Git

```bash
# 推送代码和标签
git push --follow-tags origin main

# 或分开推送
git push origin main
git push origin --tags
```

### 5. 发布到 NPM

```bash
# 发布到 NPM
npm publish --access public

# 首次发布
npm publish --access public

# 预发布版本
npm publish --tag beta --access public
```

## 📊 版本号规范

遵循 [语义化版本](https://semver.org/lang/zh-CN/)：

```text
主版本号.次版本号.修订号

例如：1.2.3
```

- **主版本号（Major）**: 不兼容的 API 变更
- **次版本号（Minor）**: 向下兼容的新功能
- **修订号（Patch）**: 向下兼容的 Bug 修复

## ✅ 发布检查清单

发布前确认：

- [ ] 所有代码已提交并推送
- [ ] 测试全部通过
- [ ] 构建成功无错误
- [ ] CHANGELOG 已更新
- [ ] README 文档已更新
- [ ] 示例代码已验证
- [ ] 版本号符合规范
- [ ] 已登录 NPM（`npm whoami`）

发布后确认：

- [ ] NPM 上版本已更新
- [ ] GitHub 上 tag 已创建
- [ ] 安装测试：`npm install @qingu-x/msgraph-orm-js@latest`
- [ ] 功能测试正常

## 🐛 常见问题

### 发布失败

**问题**: `npm ERR! 403 Forbidden`

```bash
# 检查登录状态
npm whoami

# 重新登录
npm login
```

**问题**: `npm ERR! You need to authorize this machine`

```bash
# 启用 2FA
npm login
```

### 撤销发布

```bash
# 撤销特定版本（发布后 72 小时内）
npm unpublish @qingu-x/msgraph-orm-js@0.0.5

# 撤销整个包（慎用）
npm unpublish @qingu-x/msgraph-orm-js --force
```

### 版本回退

```bash
# 删除本地 tag
git tag -d v0.0.5

# 删除远程 tag
git push origin :refs/tags/v0.0.5

# 回退 commit
git reset --hard HEAD~1
```

### 修改已发布的 tag

```bash
# 删除本地 tag
git tag -d v0.0.5

# 创建新 tag
git tag v0.0.5

# 强制推送
git push origin v0.0.5 --force
```

## 🔧 配置说明

### .versionrc.json

```json
{
  "types": [
    { "type": "feat", "section": "✨ Features" },
    { "type": "fix", "section": "🐛 Bug Fixes" },
    { "type": "docs", "section": "📝 Documentation", "hidden": false },
    { "type": "style", "section": "💄 Styles", "hidden": true },
    { "type": "refactor", "section": "♻️ Code Refactoring", "hidden": false },
    { "type": "perf", "section": "⚡ Performance", "hidden": false },
    { "type": "test", "section": "✅ Tests", "hidden": true },
    { "type": "build", "section": "📦 Build System", "hidden": false },
    { "type": "ci", "section": "👷 CI/CD", "hidden": true },
    { "type": "chore", "section": "🔧 Chores", "hidden": true }
  ]
}
```

### package.json 脚本

```json
{
  "scripts": {
    "release": "standard-version",
    "release:patch": "standard-version --release-as patch",
    "release:minor": "standard-version --release-as minor",
    "release:major": "standard-version --release-as major"
  }
}
```

## 📚 相关资源

- [语义化版本规范](https://semver.org/lang/zh-CN/)
- [Conventional Commits](https://www.conventionalcommits.org/zh-hans/)
- [standard-version](https://github.com/conventional-changelog/standard-version)
- [NPM 发布文档](https://docs.npmjs.com/cli/publish)

---

**提示**: 首次发布前建议先发布到测试环境验证
