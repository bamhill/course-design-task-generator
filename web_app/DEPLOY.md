# Zeabur 部署说明

> 课程设计任务书智能生成 · FastAPI Web App

## 一键部署到 Zeabur（免费）

### 1. 前提

- 一个 GitHub 账号
- 一个 Zeabur 账号（https://zeabur.com，用 GitHub 登录即可）

### 2. 推送到 GitHub

```bash
# 在项目根目录下执行
git init
git add web_app/
git commit -m "add web app for course design task generator"
git remote add origin https://github.com/你的用户名/课程设计任务书生成.git
git push -u origin main
```

### 3. 在 Zeabur 部署

1. 打开 https://zeabur.com → 用 GitHub 登录
2. 点击「Create Project」→ 选择你的 GitHub 仓库
3. Zeabur 自动检测到 `web_app/` 目录下的 `main.py` 和 `requirements.txt`
4. 点击「Deploy」
5. 部署完成后，Zeabur 会自动分配一个域名：`你的项目.zeabur.app`

### 4. 绑定自定义域名（可选）

Zeabur 支持绑定你自己的域名：
- 设置 → Custom Domain → 输入你的域名
- 在域名 DNS 管理处添加 CNAME 记录到 Zeabur

### 5. 访问

教师打开 `https://你的项目.zeabur.app` 即可使用。

## 本地运行

```bash
cd web_app
pip install -r requirements.txt
python main.py
# 打开 http://localhost:8000
```

## 目录结构

```
web_app/
├── main.py              ← FastAPI 后端
├── requirements.txt     ← Python 依赖
├── DEPLOY.md            ← 本文件（部署说明）
├── static/
│   └── style.css        ← 网页样式
├── scripts/
│   ├── generate_task_doc.py      ← 数据库任务书生成
│   ├── generate_cntask_doc.py    ← 计网任务书生成
│   └── generate_optimize_report.py ← 优化报告生成
└── output/              ← 生成的 docx 存放目录（自动创建）
```
