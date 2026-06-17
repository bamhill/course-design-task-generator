---
name: course-design-task-generator
description: This skill should be used when the teacher asks to generate a course design task document (课程设计任务书), optimize a draft task book, or needs a course-syllabus-based design document with civic/political-education integration. Triggers include "课程设计任务书", "课设任务书", "课程设计智能生成", "课设优化", "course design task". Do NOT trigger for general teaching syllabus generation or exam paper generation.
---

# 课程设计任务书智能生成

## 目标

固定教学大纲（理论大纲 + 课设大纲）+ 教师个性化需求 → 
**第一步** 生成课程设计任务书（初版）.docx  
**第二步** 教师提优化需求 → 输出优化总结报告+对比表.docx + 任务书（定稿）.docx

## 输入文件

放置于项目 `00_orgin/` 目录中：

| 文件 | 说明 |
|------|------|
| `理论_2026版...《数据库原理与应用》.docx` | 理论课教学大纲（固定，不可改） |
| `课设_2026版...《数据库原理与应用（SQL-Server）实践》.docx` | 课程设计大纲（固定，不可改） |
| `思政_04040357a《管理信息系统》.pdf` | 思政写作模式参考 |
| `附件7 最近一学期的课程教案_92.pdf` | 最近一学期教案（参考） |

## 教师需求录入方式（三种入口）

| 方式 | 说明 |
|------|------|
| **① 对话式** | 教师直接打字描述需求："信管专业 42 人，SQL Server 环境，重点在实践" |
| **② 文件式** | 教师上传 `.docx` / `.json` 格式的个性化需求文件 |
| **③ 参考式** | 教师给一段文字描述 + 任意参考材料，Claude 提取关键信息 |

### 对话式示例（教师说的）

> "帮我生成数据库原理与应用的任务书，信管专业 2023-1 班 42 人，SQL Server 2022，重点是动手实践，不需要分组。"

### 对话式示例（教师优化）

> "初版我看完了，有两个地方要改：1）概念设计阶段时间太紧，拆成两天；2）平时考核调到 30%，报告降到 50%。"

## 输出

| 产出 | 文件名 | 说明 |
|------|--------|------|
| 任务书（初版） | `示例_课程设计任务书_初版.docx` | WORD 文档，7 章完整版（含思政融入） |
| 优化报告+对比表 | `示例_优化总结报告+对比表.docx` | WORD 文档，含改动对比表 |
| 任务书（定稿） | `示例_课程设计任务书_定稿.docx` | 融入优化后的最终版 |

## 课程设计任务书结构

生成的 .docx 包含以下章节：

```
一、课程设计目标
二、设计任务
  任务一：物资采购和库存管理数据库设计
    阶段1：需求分析
    阶段2：概念结构设计
    阶段3：逻辑结构设计
  任务二：物资采购和库存管理数据库实现
    阶段4：数据库开发
    阶段5：功能调试与维护
三、进度安排
四、思政融合要求
五、考核标准
六、交付物清单
七、参考资料
```

## 思政融入规则

- 每阶段 1 个思政点，1-2 句话
- 融入在教学要求中，不单独设思政章节
- 每条含正例/反例对照
- 融入时机明确（导入/讲解/讨论/实操/检查）
- 融入点参考 `docs/课程设计任务书智能生成方案.md` §思政融入点表

## 专业适配

- 当前学情：**本科**（固定，不分层）
- 专业差异：按「信息管理 / 信息系统 / 数据管理 / 其它文科类」切换预设模板
- 个性化需求三维度：学生情况 / 教学倾向 / 资源环境

## 版本记录

| 版本 | 日期 | 改动 |
|------|------|------|
| v1.0 | 2026-06-17 | 初始 SKILL，对应方案 v1.2 + 原型 v0.9 |

## 网页调用模式

除了对话式调用外，本 SKILL 还支持通过 Web 网页调用：

```
教师打开网页 → 填写表单（三维度需求）
          ↓
Python 后端（FastAPI）→ 调生成脚本
          ↓
生成 .docx → 浏览器直接下载
```

### 部署方式

```
web_app/
├── main.py              # FastAPI 入口
├── templates/
│   └── index.html       # 教师填写的表单页面
├── static/
│   └── style.css        # 样式（复用原型设计）
├── generate_task_doc.py
├── generate_cntask_doc.py
├── generate_optimize_report.py
└── 00_orgin/            # 固定大纲文件
```

启动命令：
```bash
cd web_app
pip install fastapi uvicorn python-multipart python-docx
uvicorn main:app --reload --host 0.0.0.0 --port 8000
```

教师打开 `http://localhost:8000` 即可使用。

## 参考资料

- `docs/课程设计任务书智能生成方案.md` — 完整方案文档（v1.2）
- `docs/要求_Skill开发检查清单.md` — Skill 检查清单（v1.0）
- `docs/prototype/` — 设计原型及版本文件
- `docs/示例_课程设计任务书_初版.md` — 任务书示例
- `docs/generate_task_doc.py` — WORD 文档生成脚本（数据库版）
- `docs/generate_cntask_doc.py` — WORD 文档生成脚本（计网版）
- `docs/generate_optimize_report.py` — 优化报告生成脚本
