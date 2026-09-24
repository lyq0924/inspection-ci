# 自动化巡检 CI/CD 部署指南

## 完整流程概览

```
准备工作 → 创建 GitHub 仓库 → 推送代码 → 配置 Secrets → 手动触发巡检 → 查看报告
```

---

## 一、准备工作

### 1.1 安装必要工具

| 工具 | 用途 | 下载地址 |
|------|------|----------|
| Git | 代码版本管理 | https://git-scm.com/download |
| Node.js | 运行巡检脚本 | https://nodejs.org（选 LTS 版本） |
| GitHub 账号 | 托管代码、运行 CI/CD | https://github.com/signup |

### 1.2 验证安装

打开 Git Bash，执行：

```bash
git --version        # 应显示 git version 2.x.x
node --version       # 应显示 v18.x.x 或 v20.x.x
npm --version        # 应显示 9.x.x 或 10.x.x
```

---

## 二、创建 GitHub 仓库

### 2.1 在 GitHub 上创建仓库

1. 打开浏览器，访问：https://github.com/new
2. 填写信息：
   - **Repository name**: `inspection-ci`
   - **Description**: `平台自动化巡检`（可选）
   - **选择 Public**（免费）
   - **勾选 Add a README file**
3. 点击 **Create repository**

### 2.2 记录仓库地址

创建成功后，复制仓库地址：

```
https://github.com/你的用户名/inspection-ci.git
```

---

## 三、推送代码到 GitHub

### 3.1 配置 Git 身份（首次使用需要）

```bash
git config --global user.email "你的GitHub邮箱"
git config --global user.name "你的GitHub用户名"
```

### 3.2 推送代码

打开 Git Bash，执行：

```bash
# 进入项目目录
cd /g/develop/AI/Qoder

# 初始化仓库并提交代码
git init
git add .
git commit -m "feat: 初始化巡检项目"
git branch -M main

# 关联远程仓库（替换成你的地址）
git remote add origin https://github.com/你的用户名/inspection-ci.git

# 推送
git push -u origin main
```

### 3.3 完成认证

推送时会弹出浏览器要求登录 GitHub：

1. 浏览器自动弹出 → 登录 GitHub 账号
2. 点击 **Authorize** 授权
3. 回到 Git Bash，推送自动完成

> 如果浏览器没弹出，手动访问 https://github.com/login/device 输入设备码

---

## 四、配置 Secrets（密钥）

巡检脚本需要 Midscene AI 的 API Key，需要安全地配置到 GitHub。

### 4.1 进入 Secrets 配置页面

方式一：手动导航
```
GitHub 仓库页面 → Settings → Secrets and variables → Actions
```

方式二：直接访问 URL
```
https://github.com/你的用户名/inspection-ci/settings/secrets/actions
```

### 4.2 添加 Secrets

点击 **New repository secret**，添加以下两个：

| Name | Value | 说明 |
|------|-------|------|
| `MIDSCENE_API_KEY` | 你的 API Key | Midscene AI 的密钥 |
| `MIDSCENE_MODEL` | `gpt-4o` | 使用的模型名称 |

> API Key 可以从你的 `.env` 文件中找到，或者从 Midscene 控制台获取

---

## 五、手动触发巡检

### 5.1 进入 Actions 页面

```
GitHub 仓库页面 → 点击顶部 "Actions" 标签
```

### 5.2 选择工作流

左侧列表中找到 **"自动化巡检"**，点击它。

### 5.3 运行工作流

1. 点击右侧 **"Run workflow"** 按钮
2. 填写参数：

```
┌─────────────────────────────────────┐
│ Branch: main                        │
│ Environment: [123 ▼]               │  ← 选择环境
│ Scenario files: [all]              │  ← 场景选择
│ Priority: [全部 ▼]                 │  ← 可选
│ Module filter: []                  │  ← 可选
└─────────────────────────────────────┘
```

3. 点击绿色的 **"Run workflow"** 按钮

### 5.4 参数说明

| 参数 | 说明 | 示例 |
|------|------|------|
| Environment | 巡检哪个环境 | 123、124 |
| Scenario files | 跑哪些场景 | `all`=全部、`1`=第一个、`1,3`=指定 |
| Priority | 按优先级过滤 | 留空=全部、P0、P1、P2、P3 |
| Module filter | 按模块过滤 | 留空=全部、`系统管理` |

---

## 六、查看结果

### 6.1 观察运行过程

点击运行记录，可以看到每个步骤的状态：

```
📥 检出代码              ✅ 5s
🔧 设置 Node.js          ✅ 10s
📦 安装依赖              ✅ 1m
🎭 安装 Playwright       ✅ 2m
🔍 执行自动化巡检         ✅ 10m
📤 上传巡检报告           ✅ 5s
📊 输出巡检摘要           ✅ 3s
```

点击任意步骤可以查看**详细日志**。

### 6.2 下载报告

运行完成后：

1. 页面底部出现 **Artifacts** 区域
2. 点击 `inspection-report-123-1` 下载报告
3. 报告包含：Excel、HTML、Markdown 格式

### 6.3 查看摘要

在运行记录页面，可以看到巡检摘要：

```
## 🔍 巡检完成
- 环境: 123
- 场景选择: all
- 优先级: 全部
- 报告: 请查看 Artifacts 下载
```

---

## 七、常见问题

### Q1: 巡检失败，提示网络错误

**原因**: GitHub Actions 在公网运行，无法访问内网地址（10.x.x.x）

**解决**: 需要在公司内网机器上安装 Self-hosted Runner（后续配置）

### Q2: 巡检失败，提示 API Key 错误

**原因**: Secrets 中的 API Key 配置错误

**解决**: 检查 Settings → Secrets 中的 `MIDSCENE_API_KEY` 是否正确

### Q3: 提示场景文件不存在

**原因**: `scenarios/` 目录下没有 `.yml` 文件

**解决**: 把你的场景文件放到 `scenarios/` 目录下

### Q4: 推送代码时认证失败

**解决**: 
```bash
# 清除旧凭证
git credential reject

# 重新推送，会再次弹出登录
git push
```

---

## 八、后续优化（可选）

### 8.1 定时自动巡检

在 `.github/workflows/inspection.yml` 中添加定时触发：

```yaml
on:
  schedule:
    - cron: '0 1 * * *'  # 每天早上 9 点（UTC+8）
  workflow_dispatch:       # 保留手动触发
    ...
```

### 8.2 内网访问（Self-hosted Runner）

如果巡检目标是内网系统，需要在内网机器上安装 Runner：

1. GitHub 仓库 → Settings → Actions → Runners
2. 点击 **New self-hosted runner**
3. 按指引在公司内网的一台电脑上安装 Runner
4. 修改 workflow 中的 `runs-on: ubuntu-latest` 为 `runs-on: self-hosted`

### 8.3 通知到企业微信

在 workflow 中添加通知步骤：

```yaml
- name: 通知企业微信
  if: always()
  run: |
    curl -X POST "https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=你的KEY" \
      -H 'Content-Type: application/json' \
      -d '{
        "msgtype": "markdown",
        "markdown": {
          "content": "## 🔍 巡检完成\n> 环境: ${{ inputs.environment }}\n> 结果: 请查看 GitHub"
        }
      }'
```

---

## 快速参考

### 常用命令

```bash
# 推送代码
git add .
git commit -m "说明"
git push

# 查看状态
git status

# 查看日志
git log --oneline
```

### 文件结构

```
G:\develop\AI\Qoder\
├── .github/workflows/inspection.yml   # CI/CD 工作流
├── inspection-ci.mjs                  # CI 专用脚本（无交互）
├── inspection.mjs                     # 本地交互式脚本
├── inspection-task.mjs                # Task YAML 执行器
├── package.json                       # 项目依赖
├── data/0 shared/environments.json    # 环境配置
└── scenarios/*.yml                    # 场景文件
```
