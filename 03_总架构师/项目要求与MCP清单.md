# 项目要求与 MCP 清单

> 本文档由总架构师维护，是**系统配置阶段的总依据**。新用户复制总架构师初始提示词后，由总架构师据此引导完成全部配置。
> 目标：**零基础小白也能按步骤跑通**。所有路径均为相对路径（相对系统根目录）。
>
> **重要**：本系统同时支持 **Trae** 和 **TraeWork** 两个平台。总架构师须先检测用户使用的客户端，再按对应方案配置。检测方法见第 0 节。

---

## 0. 平台检测与适配（总架构师首条指令）

**首次对话时必须先检测用户平台**，然后走对应配置路线。

### 0.1 检测方法

1. **观察环境线索**：
   - 用户提到"TraeWork"、"插件"、"腾讯文档 MCP" → TraeWork
   - 用户提到"Trae"、"solo_agent"、"pro 模式" → Trae
   - 可用 MCP 列表中有 `mcp_plugin_Tencent_Docs_*` → TraeWork
   - 可用 MCP 列表中只有标准 MCP（office-mcp 等）→ 可能是 Trae，也可能是 TraeWork（需进一步确认）

2. **直接询问**（如果线索不明确）：
   > "请问你使用的是 Trae 还是 TraeWork？我好为你推荐最合适的配置方案。"

### 0.2 适配策略

| 维度 | Trae（基础版） | TraeWork（进阶版） |
|---|---|---|
| 核心 MCP | office-mcp（必需）+ mcp-jobs / jobsniper（可选） | office-mcp（必需）+ 腾讯文档 MCP（强推荐）+ 其他 |
| 情报采集方式 | 浏览器直读 + WebSearch + 可选 jobsniper/mcp-jobs | **sheet-mcp（腾讯文档）** 主力 + 浏览器插件补充 |
| Excel 处理 | office-mcp + Python openpyxl | office-mcp + xlsx skill（内置） |
| Word 处理 | office-mcp | office-mcp + docx skill（内置） |
| PDF 处理 | office-mcp 导出 | office-mcp 导出 + pdf skill（内置） |
| 知识记忆 | 落盘文件（01_共享区/） | 落盘文件 + Knowledge Graph Memory MCP |
| 浏览器能力 | WebSearch / WebFetch（内置） | TRAE-browseruse skill + 浏览器插件 |
| 配置难度 | 中（需手动装 office-mcp） | 低（腾讯文档 MCP 可从市场一键添加） |

**核心原则**：
- 两套方案**业务逻辑完全一致**（角色边界、协作规则、输出格式、看板生成全部相同）
- 差异仅在于**工具层**（用什么 MCP、用什么方式读数据）
- 总架构师负责屏蔽平台差异，给用户统一的体验
- 用户可以随时迁移平台，数据和规则完全兼容

---

## 1. 系统依赖一览

### 1.1 通用依赖（两个平台都需要）

| 依赖 | 类别 | 必需/可选 | 说明 |
|---|---|---|---|
| AI 客户端（Trae 或 TraeWork） | 运行载体 | 必需 | 本系统按 Trae 设计、TraeWork 增强；其他支持 MCP 与多对话的客户端也可用 |
| 大模型（DeepSeek V4 / Qwen 等） | 推理 | 必需 | 先用客户端免费额度，不够再自备 API |
| Python（≥ 3.10）+ openpyxl | 运行环境 | 必需 | 运行可视化看板生成脚本（`pip install openpyxl`） |
| office-mcp | MCP | 必需 | 文档处理（简历/表格/PPT），系统核心能力 |

### 1.2 Trae 平台可选依赖

| 依赖 | 类别 | 说明 |
|---|---|---|
| LibreOffice（≥ 7.0） | 本地软件 | office-mcp 需要它来读写 Word/Excel |
| mcp-jobs | MCP | 岗位搜索辅助（偏社招/私企） |
| JobSniper (jobsniper) | MCP | 岗位知识库检索（高级，智联采集依赖 WSL） |

### 1.3 TraeWork 平台推荐依赖

| 依赖 | 类别 | 说明 |
|---|---|---|
| 腾讯文档 MCP（4个） | MCP | 从 MCP 市场一键添加。sheet-mcp 用于情报采集，doc-mcp 用于在线文档，slide-mcp 备用 |
| Knowledge Graph Memory | MCP | 从 MCP 市场添加，跨会话结构化记忆 |
| 浏览器控制插件 | 插件 | TraeWork 内置，用于非腾讯文档的网页情报 |
| xlsx / docx / pdf skill | 内置 Skill | TraeWork 自带，复杂格式化时使用 |

> **注意**：TraeWork 用户的 office-mcp 仍需手动配置（本地文件读写刚需），腾讯文档 MCP 从市场一键添加即可。

---

## 2. MCP 清单与注册方法

> 以下 JSON 为注册配置模板。`<占位>` 处由总架构师按用户实际安装路径替换。
> 注册入口：客户端 → 设置 → MCP → 手动添加 → 粘贴 JSON。

### 2.1 office-mcp（必需，两个平台都要）

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
      "args": ["<office-mcp目录>/server.py"],
      "env": {
        "PYTHONUTF8": "1",
        "PYTHONIOENCODING": "utf-8",
        "PYTHONDONTWRITEBYTECODE": "1",
        "OFFICE_MCP_DEFAULT_FOLDER": "<系统根目录>"
      }
    }
  }
}
```

- **注意事项**：中文乱码时配置 `env.PYTHONUTF8=1`；工具缺失/调用报错时，检查 LibreOffice 是否安装、server.py 是否可运行。

### 2.2 腾讯文档 MCP（TraeWork 用户强推荐）

- **用途**：在线表格/文档/幻灯片读写。核心价值是**直接读取校招汇总腾讯文档**（API 级，比浏览器方案快数倍）。
- **安装方式**：TraeWork → MCP 市场 → 搜索"腾讯文档"→ 添加（4 个相关 MCP 都加上）。
- **核心工具**：
  - `sheet-mcp`：`get_sheet_info`（表结构）、`get_cell_data`（读数据，支持 return_csv）、`set_cell_value`/`set_range_value`（写数据）
  - `doc-mcp`：在线文档读写
  - `slide-mcp`：在线幻灯片（备用）
  - `tencent-docs（综合）`：空间管理、网页剪藏 `scrape_url`、OCR
- **情报源使用**（选岗老师）：见 `05_选岗老师/README.md` 第 5.1 节。

### 2.3 Knowledge Graph Memory（TraeWork 用户可选）

- **用途**：跨会话知识图谱记忆，沉淀用户画像、系统规则、决策记录。
- **安装方式**：TraeWork → MCP 市场 → 搜索"Knowledge Graph Memory"→ 添加。
- **配置项**：
  - `MEMORY_FILE_PATH`：`<系统根目录>/01_共享区/90_知识图谱/knowledge_graph.jsonl`
  - （先手动创建 `90_知识图谱` 目录）
- **使用场景**：用户画像沉淀、重要决策记录、规则变更追踪。

### 2.4 mcp-jobs（可选，两个平台通用）

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

### 2.5 JobSniper（可选，高级，两个平台通用）

- **用途**：岗位知识库检索。
- **注意**：智联自动采集依赖 WSL 环境，默认不启用；普通用户不推荐安装。
- 需要时由总架构师按官方仓库（`github.com/stars-wei/JobSniper`）指导。

---

## 3. Skill 清单

### 3.1 通用 Skill（随仓库分发）

| Skill | 位置 | 用途 |
|---|---|---|
| 投递状态机 | `02_系统配置/skill_共享/投递状态机.md` | 投递/笔面试状态取值规范（所有角色共用） |
| 投递表更新（选岗老师） | `05_选岗老师/skills/` | 按状态机更新投递总表 |
| 仪表盘生成（选岗老师） | `05_选岗老师/skills/` | 读投递总表 → 生成可视化看板 |
| 待办生成（班主任） | `04_班主任/skills/` | 与用户沟通后生成「今日+本周」待办 |
| 待办看板生成（班主任） | `04_班主任/skills/` | 读待办MD → 生成可视化看板 |
| 面试题生成（学习指导老师） | `06_学习指导老师/skills/` | 生成模拟面试题单 |
| 八股题单（学习指导老师） | `06_学习指导老师/skills/` | 维护笔试题库 |
| 简历优化-按方向（简历优化老师） | `07_简历优化老师/skills/` | 定向优化简历 |

### 3.2 TraeWork 内置增强 Skill

> 以下为 TraeWork 内置 Skill，总架构师根据需要引导各角色使用。

| Skill | 用途 | 适用角色 |
|---|---|---|
| xlsx | 专业 Excel 处理（格式化/图表/复杂操作） | 选岗老师、总架构师 |
| docx | Word 文档精细排版 | 简历优化老师、学习指导老师 |
| pdf | PDF 读取/生成/合并 | 全体 |
| html-report | 专业 HTML 报告生成 | 总架构师、选岗老师 |
| dynamic-ui | 对话内直接渲染图表/流程图 | 总架构师、选岗老师 |
| TRAE-browseruse | 浏览器自动化（网页情报） | 选岗老师 |
| research-guide | 研究报告撰写 | 选岗老师（行业/岗位分析） |
| skill-creator | 自定义 Skill 创建 | 总架构师（体系搭建） |

---

## 4. 可视化看板（选岗老师产出，随仓库分发）

- **文件位置**：`05_选岗老师/可视化/`（`update_dashboard.py` 生成脚本 + `dashboard_template.html` 模板；生成的 `投递看板.html` 由脚本运行时生成，**不随仓库分发**）。
- **运行方式**：数据写入 Excel 后，选岗老师运行 `python "05_选岗老师/可视化/update_dashboard.py"` 即自动生成看板（纯 HTML，双击即开，无需服务器）。
- **依赖**：Python + openpyxl（`pip install openpyxl`，由总架构师确认/安装）。
- **替换姓名（一次性）**：打开 `update_dashboard.py`，把 `USER_NAME = '{{用户姓名}}'` 改为用户姓名（看板标题显示该姓名）。
- **交付用户**：在 `00_交互区/系统输出/` 创建 `.lnk` 快捷方式指向 `投递看板.html`，用户双击即可查看。

### 4.1 待办看板（班主任产出）

- **文件位置**：`04_班主任/可视化/`（`update_todo.py` 生成脚本 + `todo_template.html` 模板；生成的 `待办看板.html` 不随仓库分发）。
- **运行方式**：班主任待办落盘到 `04_班主任/待办历史/` 后，运行 `python "04_班主任/可视化/update_todo.py"` 即自动生成看板。
- **替换姓名（一次性）**：同样替换 `update_todo.py` 中 `USER_NAME = '{{用户姓名}}'`。
- **交付用户**：在 `00_交互区/系统输出/` 创建 `.lnk` 快捷方式指向 `待办看板.html`。

---

## 5. 目录初始化流程（总架构师首次引导用户完成）

> **第一步先检测平台**（第 0 节），确认是 Trae 还是 TraeWork，然后走对应流程。

1. **检查环境**：确认 AI 客户端、大模型可用；确认 Python 安装情况；Trae 用户额外检查 LibreOffice。
2. **安装核心 MCP**：
   - Trae 用户：安装 office-mcp（必需）+ 可选 mcp-jobs / jobsniper
   - TraeWork 用户：安装 office-mcp（必需）+ 腾讯文档 MCP（市场一键添加）+ 可选 Knowledge Graph Memory
3. **指导放置材料**：引导用户将简历、成绩单等放入 `00_交互区/用户提供/`（见该目录 README）。
4. **简历入库存档**：引导用户将定稿简历放入 `01_共享区/20_简历库/`。
5. **情报源配置**：
   - Trae 用户：询问用户是否有腾讯文档情报源链接，如有则用浏览器直读方案
   - TraeWork 用户：推荐添加两个腾讯文档情报源（V5学长、毕业帮），用 sheet-mcp 读取
6. **建立用户档案**：提示班主任在首次对话时向用户收集档案（`01_共享区/90_用户画像/`）。
7. **投递表初始化**：指导选岗老师确认 `05_选岗老师/数据/` 两个模板 Excel（表头已内置）可用；如需多方向可复制 sheet 改名。
8. **清除占位文件**：删除仓库内所有 `.gitkeep` 占位文件（它们仅用于在 GitHub 上保留空目录，本地运行不需要）。
9. **系统自检**：检查各目录 README 是否齐全、模板是否就位、MCP 是否可调用，确认后宣布"系统就绪"。

---

## 6. 环境检查命令（总架构师自查用）

| 检查项 | 命令 | 平台 |
|---|---|---|
| Python 版本 | `python --version` | 通用 |
| Node 版本 | `node --version` | 通用（安装 mcp-jobs 时） |
| LibreOffice | Windows 检查安装目录；Mac 检查 `/Applications/LibreOffice.app` | Trae（必需）/ TraeWork（可选） |
| MCP 连接 | 客户端 MCP 面板查看各服务状态 | 通用 |
| 腾讯文档 MCP | 调用 `get_sheet_info` 读一个公开表格验证 | TraeWork |
| office-mcp | 调用 `excel_list_sheets` 读一个本地 Excel 验证 | 通用 |

---

## 7. 常见问题

| 问题 | 处理 |
|---|---|
| MCP 服务显示启动失败 | 查看报错日志 → 检查路径/入口文件/依赖是否安装 |
| 中文乱码 | MCP 配置加 `env: { PYTHONUTF8: "1" }` |
| 无法读写 Word/Excel | 检查 LibreOffice 是否安装且版本 ≥ 7.0 |
| 提示词里提到 skill 但找不到 | skill 在各角色 `skills/` 目录，随仓库分发；"规划中"表示该对话首次使用时才落地 |
| TraeWork 腾讯文档 MCP 读不到数据 | 检查文档链接是否公开（无需登录）、sheet_id 是否正确 |
| 从 Trae 迁移到 TraeWork 数据会丢吗？ | 不会。所有数据都在本地文件夹里，平台只是工具。把同一个文件夹在 TraeWork 中打开即可 |
