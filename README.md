# 课程设计任务书智能生成

> 同一个固定大纲，不同老师填不同需求，生成属于自己班的任务书。

## 解决的问题

同一门课由多位老师平行授课时，每位老师都有自己的教学风格和侧重点，需要在自己的班级里自由发挥，但大纲是统一的。

**核心流程：**
1. 教师录入个性化需求（学生维度 + 教学倾向 + 资源环境）
2. 自动生成 WORD 版课程设计任务书（初版）
3. 教师提优化需求 → 输出优化总结报告 + 改动对比表 + 任务书定稿

## 快速使用

### 方式一：对话式（推荐，零门槛）

在 Claude Code 中直接说出需求：

```
"帮我分配计算机网络课程设计任务，信管 2024 级 70 人，分 10 个不同题目"
"生成数据库原理与应用的课设任务书，信管 2023-1 班 42 人"
"第二步优化：概念设计阶段拆成两天，平时考核调到 30%"
```

### 方式二：网页版

```bash
pip install fastapi uvicorn python-multipart python-docx
cd web_app
uvicorn main:app --host 0.0.0.0 --port 8000
# 打开 http://localhost:8000
```

### 方式三：命令行

```bash
git clone https://github.com/bamhill/course-design-task-generator.git
pip install python-docx
python scripts/generate_cntask_doc.py    # 计算机网络
python scripts/generate_task_doc.py      # 数据库原理与应用
```

## 项目结构

```
├── SKILL.md                 # Claude Code Skill 主文件
├── scripts/                 # WORD 生成脚本
│   ├── generate_cntask_doc.py       # 计网任务书
│   ├── generate_task_doc.py         # 数据库任务书
│   ├── generate_optimize_report.py  # 优化报告
│   └── generate_cntask_optimize.py  # 计网优化报告
├── web_app/                 # 网页版（FastAPI）
│   ├── main.py
│   ├── static/style.css
│   └── DEPLOY.md            # 部署说明
├── docs/                    # 方案文档 + 检查清单 + 投稿文章
├── 大纲/                    # 固定教学大纲文件
└── 示例输出/                # WORD 样例文件
```

## 输出文档

| 产出 | 格式 | 说明 |
|------|------|------|
| 课程设计任务书（初版） | .docx | 7 章完整版（含思政融入） |
| 优化总结报告 + 对比表 | .docx | 改动对比表（原内容/改后/改因） |
| 课程设计任务书（定稿） | .docx | 融入优化后的最终版 |

## 功能特色

- ✅ **固定大纲不可动**，保证内容权威性
- ✅ **教师自由发挥**，同一门课不同老师出不同任务书
- ✅ **任务自动分配**，支持分层（基础/标准/挑战）
- ✅ **思政自然融入**，正反例对照，不硬贴标签
- ✅ **两步流程**：初版 + 优化报告，教务检查有据可查
- ✅ **三种入口**：对话式 / 网页 / 命令行
- ✅ **WORD 格式规范**，可直接提交教务处

## 投稿

本项目的 Skill 创作分享已投稿至 TRAE 官方中文社区 SOLO 技能创作赛。

---

**仓库地址：** https://github.com/bamhill/course-design-task-generator
