# CI/CD 练手项目

一个用于学习 CI/CD 流程的 Node.js REST API 项目，配合 GitHub Actions 实现自动化测试和部署。

## 项目结构

```
ci-cd-practice/
├── .github/
│   └── workflows/
│       └── ci.yml              # GitHub Actions 流水线配置
├── src/
│   ├── app.js                  # Express 应用（路由和业务逻辑）
│   └── server.js               # 服务器启动入口
├── tests/
│   └── app.test.js             # 集成测试（12 个测试用例）
├── package.json
└── README.md
```

## 快速开始

```bash
# 1. 安装依赖
npm install

# 2. 启动服务
npm start

# 3. 运行测试
npm test

# 4. 本地开发（文件变更自动重启）
npm run dev
```

## API 接口

| 方法 | 路径 | 说明 |
|------|------|------|
| GET | `/api/health` | 健康检查 |
| GET | `/api/tasks` | 获取所有任务 |
| GET | `/api/tasks/:id` | 获取单个任务 |
| POST | `/api/tasks` | 创建任务 |
| PATCH | `/api/tasks/:id` | 更新任务 |
| DELETE | `/api/tasks/:id` | 删除任务 |

## CI/CD 流水线说明

### 触发条件
- 推送到 `main` 分支
- 创建 Pull Request 到 `main` 分支

### 流水线阶段

```
代码提交 → 安装依赖 → 代码检查 → 运行测试 → 上传覆盖率 → 部署
```

1. **构建和测试**：在 Node.js 18 和 20 两个版本上运行测试
2. **部署**：测试通过后自动部署（当前为模拟部署）

## 练习建议

### 入门练习
1. 在 GitHub 上创建仓库，推送代码，观察 Actions 自动运行
2. 故意写一个会失败的测试，观察 CI 如何报错
3. 创建一个 PR，观察 PR 中的 CI 检查状态

### 进阶练习
1. 添加 ESLint 配置，让代码检查真正生效
2. 将模拟部署替换为真实部署（如 Vercel、Railway）
3. 添加 `staging` 分支，实现不同分支部署到不同环境
4. 配置测试覆盖率阈值，低于 80% 时流水线失败
5. 添加 Dockerfile，在 CI 中构建和推送 Docker 镜像

## 推荐部署平台（免费）

- **Vercel** — 适合前端和 Serverless
- **Railway** — 适合 Node.js 后端
- **Render** — 支持 Web 服务和定时任务
- **Fly.io** — 全球分布式部署
