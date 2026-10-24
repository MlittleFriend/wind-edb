# wind-edb

让 AI Agent 直接操作 **Wind 金融终端 Excel 插件的 EDB 经济数据库连接**：写入、刷新、按日期精确取数、核验保存结果。

Wind 的 EDB 函数在这台机器上只接受无参数调用 `='WindFunc.xla'!edb()`，指标、日期、导出格式全部藏在**单元格的传统批注**里（Base64 → 无中央目录的 ZIP local header → `EDB\0` + 版本号 + UTF-8 JSON）。这个技能把这套非公开格式讲清楚，并提供离线检查、编码/解码、指标池、原生刷新和 PowerShell 写入工具。

## 能做什么

- **离线盘点**底稿里所有 EDB / WSET 锚点，解码批注配置，不启动 Excel、不改文件
- **跨底稿建指标池**：批量解码多个工作簿，把 code / name / unit / freq 池化成可检索的指标库
- **按指定统计期取数**：单期、区间、最近 N 期，三种日期模式
- **原生写入并刷新**：写批注、写公式、计算、读回实际返回区，保存后复查连接是否完好
- **故障定位**：区分"函数没加载"、"批注丢了"、"代码不对"这三类完全不同的失败

## 安装

把 `wind-edb` 文件夹放进你的 Agent 技能目录：

- Codex：`~/.codex/skills/wind-edb`
- 其他支持 `SKILL.md` 的 Agent：按其技能安装方式放置

也可以直接让 Agent 读 `wind-edb/SKILL.md`。

## 环境要求

- Windows
- Excel
- 已登录且有权限的 Wind 金融终端（Wind 26.6.1 实测）
- Python 3（离线工具只依赖标准库）
- `pywin32`（仅 `probe_excel.py` 需要；没有就用附带的 `edb_write.ps1`）

## 快速开始

```powershell
# 1. 看看底稿里有什么
python scripts/edb_tools.py inspect 底稿.xlsx --out 检查.json

# 2. 从多本底稿里建指标池，找你要的代码
python scripts/pool_metrics.py --roots "<指标库目录>" --terms 消费 零售 --out pool.json

# 3. 从选定锚点生成指定日期的查询配置
python scripts/workbook_query.py prepare 底稿.xlsx --sheet EDB_raw --anchor A4 `
    --codes M0001427 --start 2024-01-01 --end 2024-12-31 --output query.json

# 4. 原生刷新并核验
pwsh -File scripts/edb_write.ps1 -CommentFile cmt.txt -Anchor A4 -Workbook 底稿.xlsx -Sheet EDB_raw -Save
```

`cmt.txt` 是 `edb_tools.pack_config(config)` 的输出，即那个 Base64 批注。

## 几个容易踩的坑

| 现象 | 真相 |
|---|---|
| 公式返回 `Wind` | 那是**表头**，不是错误码 |
| `=EDB("代码","起始","截止")` 返回 `#VALUE!` | 该插件版本不认带参数写法，与代码对错无关 |
| 改了代码但仍显示原指标名 | 返回区的"指标名称"行回显的是你传入的 `name`，**不能用来校验代码真伪**，只能看数值量级 |
| 批注在，但刷新无结果 | `WindFunc.xla` 没作为加载项打开，宏没执行 |
| 用 openpyxl 改完底稿，连接全没了 | 通用表格工具会丢批注和外链，必须走 Excel COM |

## 找指标代码

指标代码**不在 Wind 安装目录里**。`etc/edb/*.chs` 是分词索引，`Indicator.xml` 加密，`wind_IndicatorTree.dat` 压缩，`formulaBasicIndicator.json` 只含财务指标，`WindFunc.xla` 是纯加载项——都提不出宏观 EDB 代码。

真正好用的是用户自己的历史底稿：`pool_metrics.py` 扫一遍，几百个指标几秒钟到手。剩下的才去 Wind 指标浏览器查。

## 目录结构

```
wind-edb/
├── SKILL.md                      技能主指令
├── agents/openai.yaml            平台接口描述
├── references/
│   ├── native-workflow.md        连接格式、写入顺序、故障表
│   ├── specific-data.md          按统计期取数、日期语义、报告单元格匹配
│   └── metric-pool.md            指标池、常用代码参考、探测与交叉校验
└── scripts/
    ├── edb_tools.py              离线检查 / 配置编解码 / CRC 校验
    ├── pool_metrics.py           跨底稿建指标池
    ├── prepare_query.py          从模板生成日期查询配置
    ├── workbook_query.py         底稿内搜索 + 生成查询
    ├── probe_excel.py            新建小样探针（需 pywin32）
    └── edb_write.ps1             PowerShell 等价写入路径
```

## 边界

- 连接格式是**实测兼容格式**，不是 Wind 官方公开稳定协议；插件升级或换机器请先做探针。
- 不包含 Wind 授权、商业数据或任何用户底稿。
- 不含本地路径、账号或机器信息，示例路径均以 `<指标库目录>`、`<Wind安装目录>` 占位。
- 业务交付仍需自行核对统计日期、单位、频率、指标口径与最新已发布期次。

## 许可

仅供个人学习与研究使用。Wind 是万得信息技术股份有限公司的商标，本项目与其无隶属关系。
