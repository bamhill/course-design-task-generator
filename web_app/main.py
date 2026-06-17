# -*- coding: utf-8 -*-
"""课程设计任务书智能生成 · FastAPI 后端"""

import os, sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

from fastapi import FastAPI, Form, File, UploadFile
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
import uvicorn

app = FastAPI(title="课程设计任务书智能生成")

BASE = os.path.dirname(os.path.abspath(__file__))
SCRIPTS = os.path.join(BASE, 'scripts')
OUTPUT = os.path.join(BASE, 'output')
os.makedirs(OUTPUT, exist_ok=True)

# 挂载静态文件
app.mount("/static", StaticFiles(directory=os.path.join(BASE, "static")), name="static")

# HTML 页面
HTML_TEMPLATE = '''
<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>课程设计任务书智能生成</title>
<link href="/static/style.css" rel="stylesheet">
</head>
<body>
<div class="wrap">
  <div class="header">
    <div class="logo">课</div>
    <h1>课程设计任务书智能生成</h1>
    <p class="sub">固定大纲 + 教师需求 → 生成 WORD 版任务书</p>
  </div>

  <form id="taskForm" action="/generate" method="post" enctype="multipart/form-data">
    <div class="card">
      <div class="card-title">① 选择课程</div>
      <div class="radio-group">
        <label class="radio-item">
          <input type="radio" name="course" value="database" checked>
          <span>数据库原理与应用（SQL-Server）</span>
        </label>
        <label class="radio-item">
          <input type="radio" name="course" value="network">
          <span>计算机网络</span>
        </label>
      </div>
    </div>

    <div class="card">
      <div class="card-title">② 填写个性化需求（留空使用默认值）</div>

      <div class="section-label">学生情况</div>
      <div class="field-row">
        <div class="field"><label>专业班级</label><input name="major_class" placeholder="如：信息管理与信息系统 · 2024 级"></div>
        <div class="field"><label>人数</label><input name="student_count" placeholder="如：70"></div>
      </div>
      <div class="field-row">
        <div class="field"><label>知识基础</label><input name="knowledge_base" placeholder="留空默认：已修完理论课"></div>
        <div class="field"><label>是否需要分组</label>
          <select name="needs_grouping">
            <option value="">使用默认</option>
            <option value="yes">是，每组 3-4 人</option>
            <option value="no">否，独立完成</option>
          </select>
        </div>
      </div>

      <div class="section-label">教学倾向</div>
      <div class="field-row">
        <div class="field"><label>重点倾向</label>
          <select name="focus">
            <option value="">使用默认</option>
            <option value="theory">理论深度</option>
            <option value="practice" selected>动手实践</option>
            <option value="project">项目导向</option>
          </select>
        </div>
        <div class="field"><label>项目/设计主题</label><input name="project_theme" placeholder="留空使用大纲默认"></div>
      </div>

      <div class="section-label">资源环境</div>
      <div class="field-row">
        <div class="field"><label>数据库 / 网络环境</label><input name="env" placeholder="如：SQL Server 2022 / Cisco Packet Tracer"></div>
        <div class="field"><label>教材</label><input name="textbook" placeholder="留空使用大纲默认"></div>
      </div>

      <div class="field-row">
        <div class="field"><label>上传补充文件（可选）</label><input type="file" name="extra_file"></div>
      </div>
    </div>

    <button type="submit" class="btn">生成 WORD 任务书 ↓</button>
  </form>

  <div id="loading" style="display:none;text-align:center;padding:48px 0">
    <div class="spinner"></div>
    <p style="margin-top:16px;color:#7D7974">正在生成任务书…</p>
  </div>

  <div class="footer">© 2026 · 课程设计任务书智能生成</div>
</div>

<script>
document.getElementById('taskForm').addEventListener('submit', function(e) {
  document.getElementById('loading').style.display = 'block';
});
</script>
</body>
</html>
'''

@app.get("/", response_class=HTMLResponse)
def index():
    return HTML_TEMPLATE

@app.post("/generate")
async def generate(
    course: str = Form(...),
    major_class: str = Form(""),
    student_count: str = Form(""),
    knowledge_base: str = Form(""),
    needs_grouping: str = Form(""),
    focus: str = Form(""),
    project_theme: str = Form(""),
    env: str = Form(""),
    textbook: str = Form(""),
    extra_file: UploadFile = File(None),
):
    if course == "database":
        script = os.path.join(SCRIPTS, 'generate_task_doc.py')
        out_name = '示例_课程设计任务书_初版.docx'
        display_name = '课程设计任务书_初版.docx'
    else:
        script = os.path.join(SCRIPTS, 'generate_cntask_doc.py')
        out_name = '示例_计网课程设计任务书_初版.docx'
        display_name = '计网课程设计任务书_初版.docx'

    out_file = os.path.join(OUTPUT, out_name)

    import subprocess
    result = subprocess.run(
        [sys.executable, script],
        capture_output=True, text=True,
        cwd=SCRIPTS,
        encoding='utf-8'
    )

    if result.returncode != 0:
        return HTMLResponse(f"<h3>生成失败</h3><pre>{result.stderr}</pre>")

    # 脚本生成的 docx 默认在 SCRIPTS 目录，搬到 OUTPUT
    src = os.path.join(SCRIPTS, out_name)
    if os.path.exists(src):
        import shutil
        shutil.copy2(src, out_file)

    if not os.path.exists(out_file):
        return HTMLResponse(f"<h3>文件未找到</h3><p>生成完成但文件不存在：{out_name}</p>")

    return FileResponse(out_file, filename=display_name, media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=8000)
