# 电子表格公式影响分析（Spreadsheet Formula Impact Analysis）

一个可自托管的 Web 应用：上传 Excel（`.xlsx`）文件后，系统解析工作表中的公式，
将**跨工作表引用**、**矩形区域引用**与**名称管理器中的定义名称**展开为单元格级别的
依赖图，自动检测**循环引用**，并可针对任意单元格列出其**所有直接 / 间接受影响单元格
及完整影响路径**。

- 后端：FastAPI + openpyxl
- 前端：原生 HTML / CSS / JavaScript 单页（无需构建，无需 Node）
- 部署：Dockerfile + docker-compose 一键启动

## 功能一览

- 📤 拖拽 / 点选上传 `.xlsx`（上限 25MB，内存中保留最近 5 个文件）
- 🔗 解析单元格引用、矩形区域（`A1:B10`）、整列（`B:B`）、整行（`2:5`）、
  绝对引用（`$A$1`）、跨表引用（`参数!B2`）、带空格的工作表名（`'月报 2026'!B2`）
- 🏷 **名称管理器（定义名称）支持**：
  - 解析公式中引用的**工作簿级**与**工作表级**定义名称，把名称指向的单元格 / 区域
    展开为依赖边（区域名称逐格展开，支持 `A1,C1` 形式的多区域并集）；
  - 同名名称按 Excel 作用域规则解析：裸名先取公式所在工作表的工作表级名称再回退
    工作簿级，`工作表!名称` 形式的限定名先取该表局部名称再回退工作簿级；
  - 无法解析的名称、指向不存在工作表 / `#REF!` / 外部工作簿的名称生成
    **可定位（定位到公式单元格）的审计问题**，且**无效目标不会进入影响面**；
  - 常量名称（如 `3.14`）不产生依赖边，完整导出与审计导出均保留名称清单、
    每次使用的解析状态与证据
- 🔁 Tarjan 强连通分量算法检测循环引用（含经名称形成的环与自环），页面顶部高亮展示
- 🔍 选定单元格后展示：
  - **下游影响**：哪些单元格直接 / 间接引用了它（改了它会影响谁）
  - **上游依赖**：该单元格直接 / 间接依赖谁（它的计算需要谁）
  - 每个受影响单元格的**层级（深度）**和一条**最短影响路径**
- 🗂 可按工作表筛选、按单元格 / 公式文本搜索、切换是否只看公式单元格
- 📄 查看每个单元格的公式原文与 Excel 缓存计算值
- 🩺 **依赖健康审计**：扫描依赖图并按规则定级，支持按级别 / 工作表筛选
  - `error` —— 跨表公式指向**不存在的工作表**、公式中的**定义名称无法解析**、
    名称定义指向**不存在的工作表 / `#REF!` / 外部工作簿**
  - `warning` —— 被公式引用却**为空的单元格**、引用**超出目标工作表已用范围**、
    **扇出（直接依赖数）超过阈值**的公式节点
  - `info` —— **未被任何公式引用的孤立公式**
  - 每条问题给出级别、单元格位置、问题说明与**证据边**（名称类问题附名称、
    作用域与定义原文），并附带按级别 / 类型的汇总计数；
    页面提供审计面板与「导出审计 JSON」
- 💾 一键导出完整分析 JSON（工作表统计、公式清单、依赖边、循环环、依赖健康审计）

## 一键启动（Docker Compose）

```bash
docker compose up --build -d
```

然后浏览器打开：

> **http://localhost:8000**

停止：

```bash
docker compose down
```

如需修改端口，编辑 `docker-compose.yml` 中的 `"8000:8000"`（左为主机端口）。

容器启动后可直接验证健康状态与审计接口：

```bash
curl http://localhost:8000/api/health
# 首页下载示例文件上传后，用返回的 workbook_id 查看审计结果
curl -F "file=@samples/sample_impact.xlsx" http://localhost:8000/api/upload
curl "http://localhost:8000/api/workbooks/$WID/audit?level=warning"
```

> 上传的工作簿仅保存在容器内存中（最近 5 个），容器重启即清空；
> 镜像构建时会预生成示例文件 `samples/sample_impact.xlsx`。

## 本地开发（不使用 Docker）

需要 Python 3.11+：

```bash
pip install -r requirements.txt
python scripts/make_sample.py          # 生成示例文件（可选，仓库已附带）
uvicorn app.main:app --host 0.0.0.0 --port 8000 --reload
```

运行测试：

```bash
pip install pytest httpx
python -m pytest tests/ -v
```

## 端口与 API

服务监听 **8000** 端口，全部接口：

| 方法 | 路径 | 说明 |
|---|---|---|
| `GET` | `/` | 前端单页 |
| `POST` | `/api/upload` | 上传 xlsx（multipart 字段名 `file`），返回工作簿 id、统计与循环信息 |
| `GET` | `/api/workbooks/{id}` | 工作簿统计摘要 |
| `GET` | `/api/workbooks/{id}/cells?sheet=&q=&formulas_only=` | 单元格列表（支持工作表筛选 / 搜索） |
| `GET` | `/api/workbooks/{id}/impact?cell=参数!B2&direction=downstream` | 影响面分析；`direction` 取 `downstream`（默认，受影响）或 `upstream`（依赖链） |
| `GET` | `/api/workbooks/{id}/audit?level=&sheet=&fanout_threshold=` | 依赖健康审计；支持按级别（`error`/`warning`/`info`，可重复）、工作表、扇出阈值筛选 |
| `GET` | `/api/workbooks/{id}/audit/export` | 导出依赖健康审计 JSON（证据边不截断） |
| `GET` | `/api/workbooks/{id}/export` | 导出完整分析 JSON（含 `audit` 段） |
| `GET` | `/samples/sample_impact.xlsx` | 下载示例文件 |
| `GET` | `/api/health` | 健康检查 |

### 接口示例

```bash
# 上传并取得 workbook_id
curl -F "file=@samples/sample_impact.xlsx" http://localhost:8000/api/upload

# 全量审计（默认扇出阈值 20）
curl "http://localhost:8000/api/workbooks/$WID/audit"

# 只看错误与警告、限定「审计示例」工作表、扇出阈值改为 5
curl "http://localhost:8000/api/workbooks/$WID/audit?level=error&level=warning&sheet=审计示例&fanout_threshold=5"

# 下载审计 JSON（证据边完整，不做 100 条截断）
curl -OJ "http://localhost:8000/api/workbooks/$WID/audit/export"
```

错误响应均为 4xx：工作簿不存在 / 已被淘汰返回 `404`；
`level` 非法、`fanout_threshold` 不是 1–10000 的整数返回 `400`
（响应体形如 `{"detail": "fanout_threshold 必须是 1-10000 的整数"}`）。

### 审计 JSON 结构

```json
{
  "workbook_id": "ab12cd34ef56",
  "issues": [
    {
      "code": "missing_sheet",
      "level": "error",
      "sheet": "审计示例",
      "cell": "审计示例!B1",
      "cell_display": "审计示例!B1",
      "message": "引用了不存在的工作表 '不存在的表'（不存在的表!A1）",
      "formula": "=不存在的表!A1+1",
      "target_sheet": "不存在的表",
      "target_display": "不存在的表!A1",
      "evidence": [
        { "from": "审计示例!B1", "from_display": "审计示例!B1",
          "to_display": "不存在的表!A1", "ref": "不存在的表!A1" }
      ]
    },
    {
      "code": "empty_ref",
      "level": "warning",
      "sheet": "参数",
      "cell": "参数!B8",
      "cell_display": "参数!B8",
      "message": "单元格 参数!B8 为空，但被 2 个公式引用",
      "formula": null,
      "referenced_by": ["汇总!B7", "审计示例!B3"],
      "evidence": [
        { "from": "汇总!B7", "from_display": "汇总!B7",
          "to": "参数!B8", "to_display": "参数!B8" }
      ]
    },
    {
      "code": "undefined_name",
      "level": "error",
      "sheet": "名称",
      "cell": "名称!B11",
      "cell_display": "名称!B11",
      "message": "名称 '从未定义的名称' 无法解析：工作簿中不存在工作表级（名称）同名名称（记号 从未定义的名称）",
      "formula": "=从未定义的名称+1",
      "name": "从未定义的名称",
      "name_display": "从未定义的名称",
      "name_scope": null,
      "name_qualifier": null,
      "name_target": null,
      "name_status": "undefined",
      "target_display": null,
      "evidence": [
        { "from": "名称!B11", "from_display": "名称!B11", "ref": "从未定义的名称" }
      ]
    },
    {
      "code": "high_fanout",
      "level": "warning",
      "cell": "审计示例!B4", "fanout": 200, "threshold": 20,
      "message": "直接引用 200 个单元格，超过扇出阈值 20",
      "evidence": [ { "from": "审计示例!B4", "to": "数据!B2" } ]
    }
  ],
  "total": 18,
  "counts_by_level": { "error": 1, "warning": 4, "info": 13 },
  "counts_by_code": { "missing_sheet": 1, "empty_ref": 2, "out_of_range": 1,
                      "high_fanout": 1, "orphan_formula": 13 },
  "fanout_threshold": 20,
  "filters": { "levels": ["error", "warning", "info"], "sheet": null }
}
```

问题类型（`code`）与级别：

| code | 级别 | 含义 |
|---|---|---|
| `missing_sheet` | error | 跨表公式指向不存在的工作表（证据边只含 `to_display` 文本） |
| `undefined_name` | error | 公式中的定义名称按作用域无法解析（无工作簿级或对应工作表级同名名称） |
| `invalid_name_target` | error | 名称定义损坏（`#REF!` / 外部工作簿）或指向不存在的工作表；无效目标不展开为边 |
| `empty_ref` | warning | 目标单元格无公式也无值，按目标单元格聚合所有引用它的公式为证据边 |
| `out_of_range` | warning | 显式单元格 / 区域引用超出目标工作表已用范围（整行 / 整列引用不判越界） |
| `high_fanout` | warning | 公式直接依赖的去重单元格数 **> 阈值**（默认 20，可用 `fanout_threshold` 调整） |
| `orphan_formula` | info | 公式结果未被任何其他公式引用 |

> 接口返回的单条问题最多附带 100 条证据边（超出时带 `evidence_truncated` /
> `evidence_total`）；`/audit/export` 与完整导出不截断。

### 导出 JSON 结构

```json
{
  "filename": "sample_impact.xlsx",
  "summary": { "sheets": [...], "formula_count": 33, "edge_count": 253,
               "cycle_count": 2, "defined_name_count": 7,
               "cycles": [["循环!A1","循环!B1"],["循环!C1"]] },
  "formulas": { "计算!B2": { "formula": "=B2*参数!B4", "value": null, "in_cycle": false } },
  "edges": [ { "from": "参数!B2", "to": "计算!B2" } ],
  "cycles": [ ["循环!A1", "循环!B1"], ["循环!C1"] ],
  "defined_names": {
    "definitions": [
      { "name": "税率", "scope": null, "scope_label": "工作簿级",
        "target": "参数!$B$5", "status": "ok",
        "refs": [ { "ref": "参数!$B$5", "sheet": "参数", "sheet_exists": true, "kind": "cell" } ] }
    ],
    "uses": [
      { "source": "名称!B4", "source_display": "名称!B4", "formula": "=税率*100",
        "name": "税率", "name_display": "税率", "qualifier": null,
        "scope": null, "target": "参数!$B$5", "status": "ok", "resolved": true }
    ],
    "counts_by_status": { "ok": 5, "constant": 1, "undefined": 1,
                          "bad_target": 1, "missing_sheet": 1 }
  },
  "audit": { "issues": [ ... ], "total": 29, "counts_by_level": { ... } }
}
```

`defined_names` 段在完整导出（`/export`）与审计导出（`/audit/export`）中均会输出：
`definitions` 为工作簿中全部定义名称（名称、作用域、定义原文、解析状态与目标引用），
`uses` 为公式中每一次名称使用及其命中作用域 / 目标 / 解析状态（同公式内重复使用去重）。
名称解析状态：`ok`（已展开为依赖边）、`constant`（常量 / 表达式，无边）、
`undefined`（按作用域找不到）、`bad_target`（`#REF!` / 外部工作簿 / 空定义）、
`missing_sheet`（指向不存在的工作表）；后三者只生成审计问题，不产生影响面节点。

> 说明：`value` 取自 Excel 保存时的缓存计算结果。openpyxl 不执行公式计算，
> 由程序新建、从未在 Excel 中打开保存的文件其缓存值为 `null`，属正常现象。

## 示例文件

`samples/sample_impact.xlsx`（由 `scripts/make_sample.py` 生成，也可在首页直接下载）
包含 8 个工作表，覆盖所有分析场景：

| 工作表 | 内容 |
|---|---|
| `参数` | 纯输入项（单价、数量、折扣率、税率、月数），并含一个工作表级名称「折扣」 |
| `计算` | 跨表引用链：单价×数量 → 折扣 → 税额 → 月合计 → 年合计 |
| `汇总` | 跨表矩形区域 `SUM(计算!B2:B5)`、整列引用 `SUM(参数!B:B)` |
| `月报 2026` | 带空格的工作表名（公式里写作 `='月报 2026'!...`） |
| `循环` | 故意制造的循环引用：`A1↔B1` 互引、`C1=C1*2` 自环、`D1` 为环的下游 |
| `数据` | 200 行数据表，供审计中的高扇出公式 `SUM(数据!B2:B201)` 引用 |
| `名称` | 名称管理器：工作簿级 / 工作表级同名名称按作用域解析、区域 / 常量名称、`#REF!`、指向不存在工作表、未定义名称 |
| `审计示例` | 覆盖全部审计问题：缺失工作表、越界引用、空单元格引用、高扇出、孤立公式与一条正常公式 |

推荐体验路径：

1. 首页下载并上传 `sample_impact.xlsx`，顶部可看到红色循环引用提示与审计问题计数；
2. 左侧选择 `参数!B2`（单价），右侧显示经直接引用与名称展开的受影响单元格与逐层路径；
3. 切换到「上游依赖」，选择 `'月报 2026'!B3`，可看到完整回溯链直到 `参数!B2`；
4. 用工作表筛选定位 `循环!A1`，观察循环标记与影响面的截断行为；
5. 点击右上角「🩺 依赖健康审计」，在审计面板按级别 / 工作表筛选，
   或定位到「名称」工作表查看 `undefined_name` / `invalid_name_target` 问题，
   点击证据边中的单元格可跳转到其影响面；
6. 点击「导出分析 JSON」或「导出审计 JSON」获得机器可读的完整结果（含 `defined_names` 段）。

## 实现要点

- `app/formula_parser.py`：词法式正则扫描公式，区分字符串字面量 / `#REF!` 错误 /
  函数名 / 表结构化引用，避免把 `LOG10`、`"A1"`、`表1[列名]` 误判为单元格或名称；
  输出未展开的 `RawRef` 与 `RawName`（含 `工作表!名称` 限定形式，单元格引用优先匹配）。
- `app/graph.py`：
  - 加载两次工作簿（`data_only=False` 取公式、`data_only=True` 取缓存值）；
  - 收集工作簿级（`wb.defined_names`）与工作表级（`ws.defined_names`）定义名称并预解析
    目标（仅整个定义为纯引用 / 引用并集时才视为指向单元格，`A1+1` 等表达式按常量处理）；
  - 区域展开为具体单元格，整行 / 整列按工作表已用范围截断，单个区域上限 5000 格
    （`MAX_RANGE_CELLS`，防 `SUM(A:XFD)` 之类打爆内存）；名称指向的区域同样受此上限约束；
  - 名称按作用域解析（裸名：本表局部 → 工作簿级；限定名：被限定表局部 → 工作簿级），
    解析出的边与直接引用边统一入图，因此筛选、扇出、环检测 / 影响面对名称透明生效；
  - 迭代版 Tarjan SCC 检测循环引用；
  - BFS 求影响面，记录最短路径、层级与环标记，默认最大深度 50 / 5000 节点；
  - 构图时记录不入图的异常引用（缺失工作表、超出已用范围）与名称解析记录，
    `audit()` 汇总七类依赖健康问题并按级别 / 工作表筛选，证据边超过 100 条时截断。
- `app/main.py`：FastAPI 路由与内存 LRU 式工作簿缓存（不落盘，重启即清空）。

## 目录结构

```
.
├── app/
│   ├── main.py              # FastAPI 入口与 API
│   ├── formula_parser.py    # 公式引用解析
│   ├── graph.py             # 依赖图 / 环检测 / 影响面 BFS
│   └── static/index.html    # 前端单页
├── samples/sample_impact.xlsx
├── scripts/make_sample.py   # 示例文件生成脚本
├── tests/                   # pytest 用例（67 个）
├── Dockerfile
├── docker-compose.yml
└── requirements.txt
```
