param(
    [string]$WorkbookPath = "上层产品净值数据库.xlsm",
    [string]$ModuleGroups = "data,chart,optional_panel,tool,weekly,one_page",
    [string]$RepoRoot = ""
)

# ---------- 路径解析 ----------

if (-not $RepoRoot) {
    $scriptPath = if ($MyInvocation.MyCommand.Path) {
        Split-Path $MyInvocation.MyCommand.Path -Parent
    } else {
        Get-Location
    }
    $dir = $scriptPath
    while ($dir -and -not (Test-Path (Join-Path $dir ".git"))) {
        $dir = Split-Path $dir -Parent
    }
    if (-not $dir) { throw "找不到仓库根目录（.git）" }
    $RepoRoot = $dir
}

$resolvedWorkbook = if ([System.IO.Path]::IsPathRooted($WorkbookPath)) {
    $WorkbookPath
} else {
    Join-Path $RepoRoot $WorkbookPath
}

if (-not (Test-Path $resolvedWorkbook)) {
    Write-Error "工作簿不存在: $resolvedWorkbook"
    exit 1
}

# ---------- 编码检测 ----------

# 源文件编码（git 仓库中统一使用 UTF-8）
$sourceEncoding = [System.Text.UTF8Encoding]::new($false)  # UTF-8 without BOM

# 系统默认编码（中文 Windows 为 GB2312/CP936），VBE Import 以此编码读取文件
$systemEncoding = [System.Text.Encoding]::Default

# 创建带容错 fallback 的编码器：无法转换的字符用 ? 替代，避免抛异常
$safeSystemEncoding = [System.Text.Encoding]::GetEncoding(
    $systemEncoding.CodePage,
    [System.Text.EncoderFallback]::ReplacementFallback,
    [System.Text.DecoderFallback]::ReplacementFallback
)

Write-Host "仓库根目录: $RepoRoot"
Write-Host "目标工作簿: $resolvedWorkbook"
Write-Host "系统编码: $($systemEncoding.EncodingName) (CP$($systemEncoding.CodePage))"

# ---------- 模块清单 ----------

$moduleMap = @{
    data           = @{ path = "scripts/vba/data";                  type = "bas" }
    chart          = @{ path = "scripts/vba/chart";                 type = "bas" }
    optional_panel = @{ path = "scripts/vba/optional_panel";        type = "mixed" }
    tool           = @{ path = "scripts/vba/tool";                  type = "bas" }
    weekly         = @{ path = "scripts/vba/weekly_recommendation"; type = "bas" }
    one_page       = @{ path = "scripts\vba\product_one_page";      type = "bas" }
}

$groups = $ModuleGroups -split ',' | ForEach-Object { $_.Trim() }

# ---------- 收集待导入文件 ----------

$filesToImport = @()
foreach ($g in $groups) {
    if (-not $moduleMap.ContainsKey($g)) {
        Write-Warning "未知模块组: $g，跳过"
        continue
    }
    $info = $moduleMap[$g]
    $dir = Join-Path $RepoRoot $info.path
    if (-not (Test-Path $dir)) {
        Write-Warning "目录不存在: $dir，跳过"
        continue
    }
    switch ($info.type) {
        "bas"   { $filesToImport += Get-ChildItem -Path $dir -Filter *.bas }
        "mixed" { $filesToImport += Get-ChildItem -Path $dir -Filter *.bas
                  $filesToImport += Get-ChildItem -Path $dir -Filter *.cls
                  $filesToImport += Get-ChildItem -Path $dir -Filter *.frm }
    }
}

if ($filesToImport.Count -eq 0) {
    Write-Host "没有找到需要导入的模块文件。"
    exit 0
}

Write-Host "找到 $($filesToImport.Count) 个模块待导入："
$filesToImport | ForEach-Object { Write-Host "  - $($_.FullName)" }

# ---------- 临时目录 ----------

$tempDir = Join-Path $env:TEMP "vba-sync"
New-Item -ItemType Directory -Path $tempDir -Force | Out-Null

# ---------- 辅助函数 ----------

# 模块组 → VBA 模块名前缀（VBA 模块名不能以数字开头，需加字母前缀）
$groupPrefixMap = @{
    data             = "D"
    chart            = "C"
    tool             = "T"
    weekly           = "R"
    optional_panel   = "P"
    one_page         = "O"
}

# 代码中存在跨模块类型引用时，名称必须与代码声明一致，不能仅由文件名推导。
# .frm 的值是传给 Import-VbaModule 的基础名，函数会自动加 frm 前缀。
$componentNameMap = @{
    "optional_panel/00_operation_panel_form.frm"   = "OperationPanel"
    "optional_panel/00_operation_panel_button.cls" = "clsOperationPanelButton"
}

# VBA 模块名限制 31 字符
$MAX_VBA_NAME_LENGTH = 31

# 将文件 baseName 转为合法的 VBA 模块名：不以数字开头、不含非法字符、不超31字符
function Get-VbaModuleName {
    param(
        [string]$BaseName,
        [string]$GroupName
    )
    # 如果 baseName 以字母开头则直接使用，否则加组前缀
    $name = if ($BaseName -match '^[A-Za-z]') {
        $BaseName
    } else {
        $prefix = if ($groupPrefixMap.ContainsKey($GroupName)) {
            $groupPrefixMap[$GroupName]
        } else {
            "M"
        }
        "$prefix$BaseName"
    }
    # 截断到 31 字符（VBA 模块名限制）
    if ($name.Length -gt $MAX_VBA_NAME_LENGTH) {
        $name = $name.Substring(0, $MAX_VBA_NAME_LENGTH)
    }
    return $name
}

# 编码检测与读取：尝试 UTF-8 读取，失败时回退到系统编码
function Read-VbaModuleFile {
    param([string]$FilePath)
    try {
        # 优先按 UTF-8 读取（匹配 git 仓库中的源文件编码）
        return [System.IO.File]::ReadAllText($FilePath, $sourceEncoding)
    } catch {
        try {
            # UTF-8 失败则尝试系统默认编码
            return [System.IO.File]::ReadAllText($FilePath, $systemEncoding)
        } catch {
            throw "无法读取模块文件，已尝试 UTF-8 和系统默认编码: $FilePath"
        }
    }
}

# 统一导入：将内容写入系统编码临时文件，用 VBComponents.Import 导入
# 这样 VBE 以原生编码读取文件，避免 AddFromString 的 COM 编码问题
function Import-VbaModule {
    param(
        $VBProject,
        [string]$Content,
        [string]$BaseName,
        [string]$Extension,    # .bas / .cls / .frm
        [string]$OriginalName  # 原始文件名，仅用于日志
    )

    $ext = $Extension.ToLower()

    # ---- 确保 CRLF 行尾（统一归一化） ----
    $normalized = $Content -replace "`r`n", "`n" -replace "`n", "`r`n"

    if ($ext -eq ".frm") {
        # UserForm：不能直接 Import，优先复用已有窗体原地更新代码
        $formName = "frm$BaseName"
        $existingForm = $null
        foreach ($c in $VBProject.VBComponents) {
            if ($c.Name -eq $formName) { $existingForm = $c; break }
        }
        if ($existingForm) {
            # 已存在：清空代码后重新注入
            $codeMod = $existingForm.CodeModule
            $totalLines = $codeMod.CountOfLines
            if ($totalLines -gt 0) { $codeMod.DeleteLines(1, $totalLines) }
            $codeMod.AddFromString($normalized)
            $importedComponent = $existingForm
        } else {
            # 不存在：创建新窗体（带重试，COM 可能竞争）
            $maxRetries = 3
            $success = $false
            for ($retry = 0; $retry -lt $maxRetries -and -not $success; $retry++) {
                try {
                    if ($retry -gt 0) {
                        Start-Sleep -Milliseconds 500
                        foreach ($c in $VBProject.VBComponents) {
                            if ($c.Name -eq $formName) { $VBProject.VBComponents.Remove($c); break }
                        }
                        Write-Host "    重试 UserForm ($($retry+1)/$maxRetries)..." -ForegroundColor DarkYellow
                    }
                    $importedComponent = $VBProject.VBComponents.Add(3)
                    Start-Sleep -Milliseconds 100
                    $importedComponent.Name = $formName
                    Start-Sleep -Milliseconds 100
                    $importedComponent.CodeModule.AddFromString($normalized)
                    $success = $true
                } catch {
                    if ($retry -eq $maxRetries - 1) { throw $_ }
                }
            }
        }
    } elseif ($ext -eq ".cls") {
        # 类模块：需要完整 VBE 类模块导出头，Import 才能识别为 Class 而非 Standard
        $fileContent = "VERSION 1.0 CLASS`r`n"
        $fileContent += "BEGIN`r`n"
        $fileContent += "  MultiUse = -1  'True`r`n"
        $fileContent += "END`r`n"
        $fileContent += "Attribute VB_Name = `"$BaseName`"`r`n"
        $fileContent += "Attribute VB_GlobalNameSpace = False`r`n"
        $fileContent += "Attribute VB_Creatable = False`r`n"
        $fileContent += "Attribute VB_PredeclaredId = False`r`n"
        $fileContent += "Attribute VB_Exposed = False`r`n"
        $fileContent += $normalized
        $tempFile = Join-Path $tempDir "$BaseName$Extension"
        [System.IO.File]::WriteAllText($tempFile, $fileContent, $safeSystemEncoding)
        try {
            $importedComponent = $VBProject.VBComponents.Import($tempFile)
        } catch {
            [System.IO.File]::WriteAllText($tempFile, $fileContent, $sourceEncoding)
            $importedComponent = $VBProject.VBComponents.Import($tempFile)
        }
        Remove-Item $tempFile -Force -ErrorAction SilentlyContinue
    } else {
        # 标准模块：添加 Attribute VB_Name，写入系统编码临时文件后 Import
        $fileContent = "Attribute VB_Name = `"$BaseName`"`r`n" + $normalized
        $tempFile = Join-Path $tempDir "$BaseName$Extension"
        [System.IO.File]::WriteAllText($tempFile, $fileContent, $safeSystemEncoding)
        try {
            $importedComponent = $VBProject.VBComponents.Import($tempFile)
        } catch {
            [System.IO.File]::WriteAllText($tempFile, $fileContent, $sourceEncoding)
            $importedComponent = $VBProject.VBComponents.Import($tempFile)
        }
        Remove-Item $tempFile -Force -ErrorAction SilentlyContinue
    }

    return $importedComponent
}

# ---------- Excel COM 导入 ----------

Write-Host "`n正在打开工作簿..." -NoNewline

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $wb = $excel.Workbooks.Open((Get-Item $resolvedWorkbook).FullName)
    $vbProject = $wb.VBProject
    Write-Host " 完成"

    # const: VBE component types
    $vbext_ct_StdModule   = 1
    $vbext_ct_ClassModule = 2
    $vbext_ct_MSForm      = 3
    $vbext_ct_Document    = 100   # Sheet/ThisWorkbook — 不操作

    $imported = 0
    $errors = @()

    # 清理用户模块（避免旧名遗留），保留内置模块和 UserForm（UserForm 将被原地更新）
    Write-Host "  清理旧模块..." -NoNewline
    $removedCount = 0
    $builtInNames = @("ThisWorkbook")
    foreach ($comp in $vbProject.VBComponents) {
        if ($comp.Type -eq $vbext_ct_Document) { $builtInNames += $comp.Name }
    }
    $toRemove = @()
    foreach ($comp in $vbProject.VBComponents) {
        # 跳过内置模块和 UserForm（UserForm 删除后 VBE 资源释放慢，改为后续原地更新）
        if ($comp.Name -notin $builtInNames -and $comp.Type -ne $vbext_ct_MSForm) {
            $toRemove += $comp
        }
    }
    for ($i = $toRemove.Count - 1; $i -ge 0; $i--) {
        try {
            $vbProject.VBComponents.Remove($toRemove[$i])
            $removedCount++
        } catch { }
    }
    Write-Host " $removedCount 个已清除"

    # 批量删除后短暂等待，让 VBE 完全释放 COM 资源（避免 UserForm 创建时的 CTL_E_PATHFILEACCESSERROR）
    if ($removedCount -gt 0) {
        Start-Sleep -Seconds 2
    }

    # 排序：UserForm 优先导入（VBE 创建窗体需要更多 COM 资源，趁早处理）
    $sortedFiles = @($filesToImport | Sort-Object { if ($_.Extension -eq ".frm") { return 0 } else { return 1 } }, { $_.Name })
    foreach ($file in $sortedFiles) {
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
        $ext = [System.IO.Path]::GetExtension($file.Name)

        # 确定模块所属 group（用于生成合法的 VBA 模块名）
        $fileGroup = ""
        foreach ($g in $groups) {
            if ($file.FullName -match [regex]::Escape((Join-Path $RepoRoot $moduleMap[$g].path))) {
                $fileGroup = $g
                break
            }
        }
        # 生成合法 VBA 模块名（不能以数字开头）
        # 某些组件在其他代码中以固定名称引用，优先使用显式映射。
        $componentKey = "$fileGroup/$($file.Name)"
        if ($componentNameMap.ContainsKey($componentKey)) {
            $vbaName = $componentNameMap[$componentKey]
        } elseif ($ext -eq ".frm") {
            # .frm 文件导入后 VBE 自动加 "frm" 前缀，已满足字母开头要求
            $vbaName = $baseName
        } else {
            $vbaName = Get-VbaModuleName -BaseName $baseName -GroupName $fileGroup
        }

        try {
            # ---- 读取源文件 ----
            $content = Read-VbaModuleFile -FilePath $file.FullName

            # ---- 统一通过临时文件 + Import 导入（避免 AddFromString 编码问题） ----
            $label = if ($ext -eq ".frm") { "[UserForm]" } else { "" }
            $null = Import-VbaModule -VBProject $vbProject -Content $content `
                -BaseName $vbaName -Extension $ext -OriginalName $file.Name
            Write-Host "  导入: $($file.Name) → $vbaName $label" -ForegroundColor Green

            $imported++
        } catch {
            Write-Host "  失败: $($file.Name) — $($_.Exception.Message)" -ForegroundColor Red
            $errors += "$($file.Name): $($_.Exception.Message)"
        }
    }

    if ($errors.Count -eq 0) {
        $wb.Save()
        Write-Host "`n成功导入 $imported 个模块，工作簿已保存。"
    } else {
        Write-Host "`n已导入 $imported 个模块，$($errors.Count) 个失败："
        $errors | ForEach-Object { Write-Host "  - $_" -ForegroundColor Red }
        Write-Host "工作簿未保存，请修复后重试。" -ForegroundColor Yellow
    }
} catch {
    Write-Error "导入过程出错: $($_.Exception.Message)"
    Write-Host "`n常见原因："
    Write-Host "  1. Excel 已打开目标工作簿 — 请先关闭"
    Write-Host "  2. 未开启 VBA 对象模型信任 — 检查 Excel 信任中心设置"
    Write-Host "  3. Excel 以管理员权限运行 — 改用普通用户权限"
} finally {
    # 清理临时目录
    if (Test-Path $tempDir) {
        Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
    }
    if ($wb) { try { $wb.Close($false) } catch {} }
    if ($excel) { try { $excel.Quit() } catch {} }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
