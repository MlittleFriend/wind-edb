# wind-edb 原生写入（PowerShell COM 路径）
#
# 用途：本机没有 pywin32、probe_excel.py 无法运行时，用等价流程写入并核验 EDB 连接。
# 与 probe_excel.py 的唯一差别是宿主语言；协议完全一致：
#   1) Workbooks.Open(WindFunc.xla)      加载宏，否则公式只返回占位文本
#   2) cell.AddComment(<base64>)          配置必须先写
#   3) cell.Formula = "='WindFunc.xla'!edb()"
#   4) Calculate + 有限轮询
#   5) 读实际 UsedRange，而不是假定布局
#
# 示例：
#   pwsh -File edb_write.ps1 -CommentFile cmt.txt -Anchor A4
#   pwsh -File edb_write.ps1 -CommentFile cmt.txt -Anchor A4 -Workbook 底稿.xlsx -Sheet 'EDB_raw' -Save
#
# 安全约束：只 Quit 自己创建/自己打开的实例；不给 -Workbook 时写新建工作簿，永不覆盖用户文件。

param(
    [Parameter(Mandatory = $true)][string]$CommentFile,
    [string]$Anchor = "A4",
    [string]$Workbook = "",
    [string]$Sheet = "",
    [string]$Xla = $env:WIND_EDB_XLA,
    [int]$TimeoutSec = 30,
    [switch]$Save,
    [switch]$KeepVisible
)

$ErrorActionPreference = "Continue"
if (-not (Test-Path $CommentFile)) { throw "批注文件不存在: $CommentFile" }
if (-not $Xla) {
    throw "请用 -Xla 指定 WindFunc.xla 路径，或设置环境变量 WIND_EDB_XLA（通常位于 <Wind安装目录>/DataBrowse/XLA/WindFunc.xla）"
}
if (-not (Test-Path $Xla)) { throw "找不到 WindFunc.xla: $Xla" }

$b64 = (Get-Content $CommentFile -Raw).Trim()
if ($b64.Length -lt 100) { throw "批注内容过短，可能不是 pack_config 的输出" }

$excel = New-Object -ComObject Excel.Application
$excel.Visible = [bool]$KeepVisible
$excel.DisplayAlerts = $false
$excel.AskToUpdateLinks = $false
try { $excel.AutomationSecurity = 1 } catch { }   # 强制启用宏，否则 XLA 不执行

$book = $null
$addin = $null
$ownedBook = $false
try {
    if ($Workbook) {
        if (-not (Test-Path $Workbook)) { throw "底稿不存在: $Workbook" }
        $book = $excel.Workbooks.Open((Resolve-Path $Workbook).Path)
    } else {
        $book = $excel.Workbooks.Add()
        $ownedBook = $true
    }
    $addin = $excel.Workbooks.Open((Resolve-Path $Xla).Path, 0, $true)

    if ($Sheet) {
        $ws = $book.Worksheets.Item($Sheet)
    } else {
        $ws = $book.Worksheets.Item(1)
    }
    if ($null -eq $ws) { throw "找不到工作表: $Sheet" }

    $cell = $ws.Range($Anchor)
    $cell.ClearComments()
    $cell.AddComment($b64) | Out-Null
    $cell.Formula = "='WindFunc.xla'!edb()"

    $ws.Calculate()
    $deadline = (Get-Date).AddSeconds($TimeoutSec)
    $lastAddr = ""
    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Milliseconds 700
        try { $lastAddr = $ws.UsedRange.Address() } catch { }
        # 表头 + 至少一行为数据即认为返回
        $used = $ws.UsedRange
        if ($used.Rows.Count -ge 3 -and $used.Columns.Count -ge 1) {
            $probe = $ws.Cells.Item(6, 2).Text
            if ($probe -ne "") { break }
        }
    }

    Write-Output "anchor=$Anchor usedRange=$lastAddr"

    $used = $ws.UsedRange
    $rowCount = [Math]::Min($used.Rows.Count, 25)
    $colCount = [Math]::Min($used.Columns.Count, 10)
    $startRow = $used.Row
    for ($r = $startRow; $r -lt ($startRow + $rowCount); $r++) {
        $parts = @()
        for ($c = 1; $c -le $colCount; $c++) {
            $t = $ws.Cells.Item($r, $c).Text
            if ($t -ne "") { $parts += ("{0}{1}={2}" -f [char](64 + $c), $r, $t) }
        }
        if ($parts.Count -gt 0) { Write-Output ("  " + ($parts -join "  ")) }
    }

    if ($Save) {
        $book.Save()
        Write-Output "saved: $($book.FullName)"
    }
}
finally {
    if ($book -and $ownedBook -and -not $Save) { $book.Close($false) | Out-Null }
    elseif ($book -and -not $ownedBook) { $book.Close($false) | Out-Null }
    if ($addin) { $addin.Close($false) | Out-Null }
    $excel.Quit() | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    [GC]::Collect()
}
