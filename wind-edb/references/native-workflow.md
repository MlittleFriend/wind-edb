# 原生连接操作

## 格式来源与适用范围

此实现依据既有底稿的实测读取核实：无参数 edb() 从锚点传统批注读取配置。它是实测兼容格式，不是 Wind 官方公开稳定协议，插件升级后先做探针。

批注：Base64 → ZIP local header（文件名 `0`，raw DEFLATE，末尾 data descriptor，无 central directory）→ `EDB\0` + little-endian uint32 版本 `10` + UTF-8 JSON。不能直接用普通 zipfile 读取这一非完整 ZIP 片段。脚本检查 CRC、长度和版本；不支持的版本不自动重写。

JSON 至少包含 `metric`、`exportConfig`、`exportTitleOption`、`timeFrameSetting`。保留未知字段；不要从零猜测整份配置。

- metric 中 `code`、`name`、`unit`、`freq`、`beginDate` 等应来自核验过的指标元数据；复制完整条目后如需新连接可生成新的 `key` UUID，保留实际需要的转换和 changeRecord，不能盲目清空变换记录。
- timeFrameSetting 中 `periodType` 决定 obsFrom/obsTo、lastObs 等字段是否生效；全历史设置中的旧日期字段不一定是有效截止日期。具体语义以当前插件界面和返回序列核验为准。
- exportConfig / exportTitleOption 影响空值、顺序与标题行，改变这些会改变数据范围。示例脚本不强行规定用户的格式。

## 离线命令

```powershell
python scripts/edb_tools.py inspect '底稿.xlsx' --out '检查.json' --configs-dir 'configs'
python scripts/edb_tools.py pack 'configs/connection-1.json' '批注.txt'
python scripts/edb_tools.py unpack '批注.txt' '配置.json'
```

inspect 导出的连接文件编号对应 JSON 中 anchors 的 config_file。结果含业务配置/指标元数据，按用户底稿的访问权限保管，不对外上传。

## Excel 小样

```powershell
python scripts/probe_excel.py --config '配置.json' --output '工作目录/probe.xlsx' --addin '<Wind安装目录>/DataBrowse/XLA/WindFunc.xla' --timeout 90
```

若 pywin32 不在当前 Python 环境中，使用已安装它的解释器；也可指定 `--dependency-dir` 到已知可信的现有包目录。不要从未知来源安装 DLL。插件路径是本机历史位置示例，执行前必须检查实际路径。

探针创建新工作簿、加载指定插件、在 A4 写入批注和公式，等待返回区域出现数值且计算完成，再保存。输出旁边的 `.evidence.json` 保存公式、返回区域和值样本，以及离线重新读取结果。这个条件仅证明技术返回；业务日期、值、权限、单位与自动更新能力须继续核验。探针不会自动延长日期或猜测指标。

## 正式写入核心顺序（pywin32）

以下是操作片段，需先定位目标工作簿/表/锚点并备份，不能直接在未知活动单元格运行：

```python
from edb_tools import pack_config
cell = sheet.Range(anchor)
# 先保存原有批注与公式；只对已确认要替换的 EDB 锚点操作。
cell.ClearComments()
cell.AddComment(pack_config(verified_config))
cell.Formula = "='WindFunc.xla'!edb()"
excel.Calculation = -4105  # xlCalculationAutomatic
excel.CalculateBeforeSave = True
sheet.Calculate()
# 有界轮询 + 消息泵；核对目标返回范围和期次后才 Save。
```

共享 Excel 会话的 Calculation 是应用级状态。不要随意切成手动或在整个实例全量重算；优先专用实例与目标表。专用实例不一定继承当前登录插件状态，探针失败需检查插件初始化和 Wind 登录，不应因此终止用户现有实例。若使用当前会话，保存前记录/处理全局状态及用户未保存内容，不 Quit。

## 常见故障

| 症状 | 核查与处理 |
|---|---|
| 公式存在但不返回 | 检查传统批注、配置解码、插件路径、登录与数据权限 |
| #NAME? / 找不到函数 | 插件未载入或外部链接错误；不要把 [1] 当固定索引 |
| 数值划线/旧缓存 | 检查自动计算、计算完成、保存重读；不要清除字体掩盖问题 |
| 长时间等待 | 保存错误/公式/计算状态，达到上限停止；COM 单次调用也可能卡住，脚本软超时不能强制打断 COM，外部监督仅可终止确认归属本次探针的进程 |
| 数据期次错误 | 检查实际有效的时间模式、发布日期与频率，不仅看 obsTo |
| 图片图表丢失 | 避免非原生整本重存；回到备份并用 Excel 最小编辑 |
| WSET 转 EDB 后口径变了 | 恢复正确 WSET；函数统一不应凌驾于数据口径 |

`calcPr` 缺失某个可选字段不自动表示失败；manual、显式 calcCompleted=0 或 calcOnSave=0 是需要结合 Excel 状态进一步检查的信号。仅修改 XML 标志不能完成重算。
