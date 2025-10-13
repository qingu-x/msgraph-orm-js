# Changelog

All notable changes to this project will be documented in this file. See [standard-version](https://github.com/conventional-changelog/standard-version) for commit guidelines.

## [0.1.1](https://github.com/qingu-x/msgraph-orm-js/compare/v0.1.0...v0.1.1) (2025-10-13)

### ✨ Features

* **repositories:** 添加 EventRepository, MessageRepository, DriveItemRepository
* **repositories:** 支持使用查询构建器查询事件、邮件和文件
* **services:** 简化 Service 类，专注于业务逻辑功能
* **services:** 添加 delta 增量查询方法到 Mail, Calendar, File 服务
* **types:** 添加 Permission 接口定义

### 📝 Documentation

* **readme:** 更新示例展示 Repository 和 Service 的使用场景
* **examples:** 添加架构说明部分，解释 Repository vs Service 的选择原则
* **readme:** 更新 GraphClient 使用方式

### 💡 改进

* **architecture:** 实现真正的 Repository + Service 混合模式
  - Repository 负责实体 CRUD，支持查询构建器
  - Service 负责业务逻辑和特殊 API
* **query:** 邮件、事件、文件现在支持复杂查询条件
* **user-repository:** 移除冗余的 calendarEvents 方法

## [0.1.0](https://github.com/qingu-x/msgraph-orm-js/compare/v0.0.9...v0.1.0) (2025-10-13)


### ⚠ BREAKING CHANGES

* **orm:** 创建了清晰的三层架构（ORM → Repositories + Services）

### ♻️ Code Refactoring

* **orm:** rebuild orm with repository & service ([37bc7d7](https://github.com/qingu-x/msgraph-orm-js/commit/37bc7d7a4d3d014f94baad241e681f5674355771))

### [0.0.9](https://github.com/qingu-x/msgraph-orm-js/compare/v0.0.8...v0.0.9) (2025-10-12)


### ✨ Features

* **orm:** add getRaw for debug ([1dce020](https://github.com/qingu-x/msgraph-orm-js/commit/1dce020e9ebf35204e71a48cd19de7cf13b27423))

### [0.0.8](https://github.com/qingu-x/msgraph-orm-js/compare/v0.0.7...v0.0.8) (2025-10-11)


### 🐛 Bug Fixes

* **types:** defaultEndpoints's key ([516b57f](https://github.com/qingu-x/msgraph-orm-js/commit/516b57f841462b8568f7f2df40fc0fa3a81e6045))

### [0.0.7](https://github.com/qingu-x/msgraph-orm-js/compare/v0.0.6...v0.0.7) (2025-10-11)


### 🐛 Bug Fixes

* **client-credentials:** import with fullpath ([4ebdf44](https://github.com/qingu-x/msgraph-orm-js/commit/4ebdf44acf67a63ec7424cc2e8336aa28d9fb133))

### [0.0.6](https://github.com/qingu-x/msgraph-orm-js/compare/v0.0.5...v0.0.6) (2025-10-11)


### 📝 Documentation

* **all:** rewrite all docs ([5832df7](https://github.com/qingu-x/msgraph-orm-js/commit/5832df736c9e0e87391ed8e1987218a0a81403af))

### [0.0.5](https://github.com/qingu-x/msgraph-orm-js/compare/v0.0.4...v0.0.5) (2025-10-11)


### 🐛 Bug Fixes

* **vite:** add builtinModules ([5d1d878](https://github.com/qingu-x/msgraph-orm-js/commit/5d1d8783f3d783e2c812a3906e62023042543071))

### [0.0.4](https://github.com/qingu-x/msgraph-orm-js/compare/v0.0.3...v0.0.4) (2025-10-11)


### 📝 Documentation

* **npm:** SOP about npm release & publish ([f51a139](https://github.com/qingu-x/msgraph-orm-js/commit/f51a13911a9bea373c14f1fdef84752793849fe0))

### [0.0.3](https://github.com/qingu-x/msgraph-orm-js/compare/v0.0.2...v0.0.3) (2025-10-11)


### ✨ Features | 新功能

* **build:** replace rollup to vite ([e6f1c25](https://github.com/qingu-x/msgraph-orm-js/commit/e6f1c25fb60cc9aeadbf63f7f80a6a9f309c0bb2))
