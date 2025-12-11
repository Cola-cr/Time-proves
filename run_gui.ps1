# 自动为当前项目设置 Tcl/Tk 环境并启动 GUI
$ErrorActionPreference = 'Stop'

# 1) 使用项目虚拟环境的 Python
$venvPython = Join-Path $PSScriptRoot '.venv\Scripts\python.exe'
if (-not (Test-Path $venvPython)) {
  Write-Host '[错误] 未找到虚拟环境 Python：.\\.venv\\Scripts\\python.exe' -ForegroundColor Red
  Write-Host '请先创建虚拟环境并安装依赖：'
  Write-Host '  py -m venv .venv'
  Write-Host '  .\\.venv\\Scripts\\python -m pip install -r requirements.txt'
  exit 1
}

# 2) 解析基础安装目录（非 venv 目录），一般用于定位 tcl 目录
$basePrefix = & $venvPython -c "import sys; print(sys.base_prefix)"
if (-not $basePrefix) { $basePrefix = & $venvPython -c "import sys; print(sys.prefix)" }

# 3) 推断 Tcl/Tk 库目录（常见于 <base>\tcl\tcl8.6 与 <base>\tcl\tk8.6）
$tclPath = Join-Path $basePrefix 'tcl\tcl8.6'
$tkPath  = Join-Path $basePrefix 'tcl\tk8.6'

if (-not (Test-Path $tclPath)) {
  # 兼容某些发行版路径：<base>\lib\tcl8.6 / <base>\lib\tk8.6
  $altTcl = Join-Path $basePrefix 'lib\tcl8.6'
  $altTk  = Join-Path $basePrefix 'lib\tk8.6'
  if (Test-Path $altTcl -and Test-Path $altTk) {
    $tclPath = $altTcl
    $tkPath  = $altTk
  }
}

if (-not (Test-Path $tclPath) -or -not (Test-Path $tkPath)) {
  Write-Host '[警告] 未能在以下路径找到 Tcl/Tk 库：' -ForegroundColor Yellow
  Write-Host "  TCL: $tclPath"
  Write-Host "  TK : $tkPath"
  Write-Host '你仍可继续尝试启动；若报 init.tcl 错误，请确认你的 Python 安装已包含 Tcl/Tk。' -ForegroundColor Yellow
}

# 4) 设置当前会话环境变量（仅对本进程与子进程生效，不污染系统全局）
$env:TCL_LIBRARY = $tclPath
$env:TK_LIBRARY  = $tkPath

# 5) 启动应用（GUI）
Write-Host "[启动] TCL_LIBRARY=$($env:TCL_LIBRARY)" -ForegroundColor Cyan
Write-Host "[启动] TK_LIBRARY=$($env:TK_LIBRARY)" -ForegroundColor Cyan
& $venvPython (Join-Path $PSScriptRoot 'main.py') @args
