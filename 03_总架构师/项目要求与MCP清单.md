# 项目要求与 MCP 清单

> 本文档由总架构师维护，是**系统配置阶段的总依据**。新用户复制总架构师初始提示词后，由总架构师据此引导完成全部配置。
> 目标：**零基础小白也能按步骤跑通**。所有路径均为相对路径（相对系统根目录）。

## 1. 系统依赖一览

| 依赖 | 类别 | 必需/可选 | 说明 |
|---|---|---|---|
| AI 客户端（Trae / Cursor 等） | 运行载体 | 必需 | 本系统按 Trae 设计；其他支持 MCP 与多对话的客户端也可用 |
| 大模型（DeepSeek V4 / Qwen 等） | 推理 | 必需 | 见 README 模型推荐；需自备 API 额度（付费） |
| LibreOffice（≥ 7.0） | 本地软件 | 必需（用 office-mcp 时） | 用于 Word/Excel 文档读写与转换 |
| Python（≥ 3.10）+ openpyxl | 运行环境 | 必需（用可视化看板时） | 运行选岗老师的看板生成脚本（`pip install openpyxl`） |
| office-mcp | MCP | 必需 | 文档处理（简历/表格/PPT），系统核心能力 |
| mcp-jobs | MCP | 可选 | 岗位搜索辅助（偏社招/私企） |
| JobSniper (jobsniper) | MCP | 可选 | 岗位知识库检索（高级） |

## 2. MCP 清单与注册方法

> 以下 JSON 为注册配置模板。`<占位>` 处由总架构师按用户实际安装路径替换。
> 注册入口：客户端 → 设置 → MCP → 手动添加 → 粘贴 JSON。

### 2.1 office-mcp（必需）

- **用途**：读/写 Word、Excel、PPT；生成简历修改稿（.docx）、投递总表（.xlsx）、可视化报表。
- **前置依赖**：安装 LibreOffice 并确认可用；安装 Python（≥ 3.10）。
- **安装步骤**（由总架构师在对话内指导，或按官方仓库说明）：
  1. `git clone https://github.com/gawirable/office-mcp`
  2. 在 office-mcp 目录创建虚拟环境并安装依赖（Windows：`python -m venv .venv` → `.venv\Scripts\activate` → `pip install -r requirements.txt`）
  3. 确认 `server.py` 路径
- **注册配置模板**：

```json
{
  "mcpServers": {
    "office-mcp": {
      "command": "<office-mcp目录>/.venv/Scripts/python.exe",
      "args": ["<office-mcp目录>/server.py"]
    }
  }
}
```

- **注意事项**：中文乱码时配置 `"env": { "PYTHONUTF8": "1" }`；工具缺失/调用报错时，让总架构师检查 LibreOffice 是否安装、server.py 是否可运行。

### 2.2 mcp-jobs（可选）

- **用途**：岗位搜索（关键词+城市），辅助选岗老师补全私企情报。
- **前置依赖**：Node.js ≥ 18。
- **安装**：`npm install -g mcp-jobs`（或本地目录安装）。
- **注册配置模板**（入口文件以实际安装为准，常见为 `dist/mcp.js`）：

```json
{
  "mcpServers": {
    "mcp-jobs": {
      "command": "node",
      "args": ["<mcp-jobs目录>/node_modules/mcp-jobs/dist/mcp.js"]
    }
  }
}
```

> 若报错"could not determine executable to run"，说明入口文件不对，请总架构师检查包内 `dist/` 目录，改用实际 MCP 入口文件。

### 2.3 JobSniper（可选，高级）

- **用途**：岗位知识库检索。
- **注意**：智联自动采集依赖 WSL 环境，默认不启用；普通用户不推荐安装。
- 需要时由总架构师按官方仓库（`github.com/stars-wei/JobSniper`）指导。

## 3. Skill 清单

| Skill | 位置 | 用途 |
|---|---|---|
| 投递状态机 | `02_系统配置/skill_共享/投递状态机.md` | 投递/笔面试状态取值规范（所有角色共用） |
| 投递表更新（选岗老师） | `05_选岗老师/skills/` | 按状态机更新投递总表 |
| 仪表盘生成（选岗老师） | `05_选岗老师/skills/` | 读投递总表 → 生成可视化看板 |
| 待办生成（班主任） | `04_班主任/skills/` | 与用户沟通后生成「今日+本周」待办 |
| 待办看板生成（班主任） | `04_班主任/skills/` | 读待办MD → 生成可视化看板（可勾选、状态存浏览器本地） |
| 面试题生成（学习指导老师） | `06_学习指导老师/skills/` | 生成模拟面试题单 |
| 八股题单（学习指导老师） | `06_学习指导老师/skills/` | 维护笔试题库 |
| 简历优化-按方向（简历优化老师） | `07_简历优化老师/skills/` | 定向优化简历 |

> 各角色 skill 目录已随本仓库分发，无需单独下载。新增/修改 skill 需总架构师审批。

## 4. 可视化看板（选岗老师产出，随仓库分发）

- **文件位置**：`05_选岗老师/可视化/`（`update_dashboard.py` 生成脚本 + `dashboard_template.html` 模板；生成的 `投递看板.html` 由脚本运行时生成，**不随仓库分发**）。
- **运行方式**：数据写入 Excel 后，选岗老师运行 `python "05_选岗老师/可视化/update_dashboard.py"` 即自动生成看板（纯 HTML，双击即开，无需服务器）。
- **依赖**：Python + openpyxl（`pip install openpyxl`，由总架构师确认/安装）。
- **替换姓名（一次性）**：打开 `update_dashboard.py`，把 `USER_NAME = '{{用户姓名}}'` 改为用户姓名（看板标题显示该姓名）。
- **交付用户**：在 `00_交互区/系统输出/` 创建 `.lnk` 快捷方式指向 `投递看板.html`，用户双击即可查看。

## 4.1 待办看板（班主任产出）

- **文件位置**：`04_班主任/可视化/`（`update_todo.py` 生成脚本 + `todo_template.html` 模板；生成的 `待办看板.html` 不随仓库分发）。
- **运行方式**：班主任待办落盘到 `04_班主任/待办历史/` 后，运行 `python "04_班主任/可视化/update_todo.py"` 即自动生成看板（今日+本周两分区、可勾选完成、状态存浏览器本地 localStorage）。
- **替换姓名（一次性）**：同样替换 `update_todo.py` 中 `USER_NAME = '{{用户姓名}}'`。
- **交付用户**：在 `00_交互区/系统输出/` 创建 `.lnk` 快捷方式指向 `待办看板.html`。

## 5. 目录初始化流程（总架构师首次引导用户完成）

1. **检查环境**：确认 AI 客户端、大模型可用；确认 LibreOffice、Python、Node 安装情况。
2. **安装 MCP**：按第 2 节安装 office-mcp（必需）；可选 MCP 由用户决定是否安装。
3. **指导放置材料**：引导用户将简历、成绩单等放入 `00_交互区/用户提供/`（见该目录 README）。
4. **简历入库存档**：引导用户将定稿简历放入 `01_共享区/20_简历库/`。
5. **建立用户档案**：提示班主任在首次对话时向用户收集档案（`01_共享区/90_用户画像/`）。
6. **投递表初始化**：指导选岗老师确认 `05_选岗老师/数据/` 两个模板 Excel（表头已内置）可用；如需多方向可复制 sheet 改名。
7. **清除占位文件**：删除仓库内所有 `.gitkeep` 占位文件（它们仅用于在 GitHub 上保留空目录，本地运行不需要）。
8. **系统自检**：检查各目录 README 是否齐全、模板是否就位，确认后宣布"系统就绪"。

## 5. 环境检查命令（总架构师自查用）

- Python 版本：`python --version`
- Node 版本：`node --version`
- LibreOffice：Windows 检查安装目录；Mac 检查 `/Applications/LibreOffice.app`
- MCP 连接：客户端 MCP 面板查看各服务状态

## 6. 常见问题

| 问题 | 处理 |
|---|---|
| MCP 服务显示启动失败 | 查看报错日志 → 检查路径/入口文件/依赖是否安装 |
| 中文乱码 | MCP 配置加 `env: { PYTHONUTF8: "1" }` |
| 无法读写 Word/Excel | 检查 LibreOffice 是否安装且版本 ≥ 7.0 |
| 提示词里提到 skill 但找不到 | skill 在各角色 `skills/` 目录，随仓库分发；"规划中"表示该对话首次使用时才落地 |
