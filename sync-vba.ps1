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

# PowerShell 7 中 Encoding.Default 固定为 UTF-8，不能代表 VBE 实际使用的 Windows ANSI code page。
# 直接调用 Win32 GetACP，并且只在中文 CP936 环境中执行，避免 VBE 对临时文件二次误解码。
if (-not ("VbaSync.NativeMethods" -as [type])) {
    Add-Type -TypeDefinition @"
using System.Runtime.InteropServices;
namespace VbaSync {
    public static class NativeMethods {
        [DllImport("kernel32.dll")]
        public static extern uint GetACP();
    }
}
"@
}

$windowsAnsiCodePage = [VbaSync.NativeMethods]::GetACP()
if ($windowsAnsiCodePage -ne 936) {
    throw "当前 Windows ANSI code page 为 CP$windowsAnsiCodePage，VBE 中文源码同步要求 CP936。请关闭 Windows 的[使用 Unicode UTF-8 提供全球语言支持]选项，或改在 CP936 中文 Windows 上运行。"
}

# 严格 UTF-8：允许有无 BOM，但非法 UTF-8 字节必须报错。
$sourceEncoding = [System.Text.UTF8Encoding]::new($false, $true)
# 严格 CP936：任何不可表示字符必须报错，绝不静默替换为 ?。
$vbeEncoding = [System.Text.Encoding]::GetEncoding(
    936,
    [System.Text.EncoderFallback]::ExceptionFallback,
    [System.Text.DecoderFallback]::ExceptionFallback
)

Write-Host "仓库根目录: $RepoRoot"
Write-Host "目标工作簿: $resolvedWorkbook"
Write-Host "VBE 导入编码: $($vbeEncoding.EncodingName) (CP$windowsAnsiCodePage)"

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

function Get-ImportComponentInfo {
    param($File)

    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($File.Name)
    $ext = [System.IO.Path]::GetExtension($File.Name)
    $fileGroup = ""
    foreach ($g in $groups) {
        if ($File.FullName -match [regex]::Escape((Join-Path $RepoRoot $moduleMap[$g].path))) {
            $fileGroup = $g
            break
        }
    }

    $componentKey = "$fileGroup/$($File.Name)"
    if ($componentNameMap.ContainsKey($componentKey)) {
        $vbaName = $componentNameMap[$componentKey]
    } elseif ($ext -eq ".frm") {
        $vbaName = $baseName
    } else {
        $vbaName = Get-VbaModuleName -BaseName $baseName -GroupName $fileGroup
    }

    return @{
        BaseName = $baseName
        Extension = $ext
        Group = $fileGroup
        VbaName = $vbaName
        ActualComponentName = if ($ext -eq ".frm") { "frm$vbaName" } else { $vbaName }
    }
}

function Read-VbaModuleFile {
    param([string]$FilePath)
    try {
        return [System.IO.File]::ReadAllText($FilePath, $sourceEncoding)
    } catch [System.Text.DecoderFallbackException] {
        throw "模块文件不是有效的 UTF-8: $FilePath"
    }
}

function Get-UnencodableCharacters {
    param([string]$Content)

    $issues = @()
    $line = 1
    for ($i = 0; $i -lt $Content.Length; $i++) {
        $charLength = 1
        $codePoint = [int][char]$Content[$i]
        if ([char]::IsHighSurrogate($Content[$i]) -and
            $i + 1 -lt $Content.Length -and
            [char]::IsLowSurrogate($Content[$i + 1])) {
            $codePoint = [char]::ConvertToUtf32($Content[$i], $Content[$i + 1])
            $charLength = 2
        }
        $character = $Content.Substring($i, $charLength)
        try {
            $null = $vbeEncoding.GetBytes($character)
        } catch [System.Text.EncoderFallbackException] {
            $issues += "第 $line 行 '$character' (U+$($codePoint.ToString('X4')))"
            if ($issues.Count -ge 10) { break }
        }
        if ($Content[$i] -eq "`n") { $line++ }
        if ($charLength -eq 2) { $i++ }
    }
    return $issues
}

function Get-CanonicalCode {
    param([string]$Content)

    $canonical = ($Content -replace "`r`n", "`n" -replace "`r", "`n").TrimEnd([char[]]"`n")
    # VBE 会自动统一标识符大小写并删除行尾空白；这些变化与编码无关。
    $lines = @($canonical.Split([char]"`n"))
    $normalizedLines = @($lines | ForEach-Object { $_.TrimEnd([char[]]" `t") })
    return ($normalizedLines -join "`n").ToLowerInvariant()
}

function Assert-ImportedCodeMatches {
    param(
        $Component,
        [string]$ExpectedContent,
        [string]$OriginalName
    )

    $codeModule = $Component.CodeModule
    $actualContent = if ($codeModule.CountOfLines -gt 0) {
        $codeModule.Lines(1, $codeModule.CountOfLines)
    } else {
        ""
    }
    $expected = Get-CanonicalCode -Content $ExpectedContent
    $actual = Get-CanonicalCode -Content $actualContent
    if ($actual -ceq $expected) { return }

    $expectedLines = @($expected.Split([char]"`n"))
    $actualLines = @($actual.Split([char]"`n"))
    $maxLines = [Math]::Max($expectedLines.Count, $actualLines.Count)
    $differentLine = 1
    for ($i = 0; $i -lt $maxLines; $i++) {
        $expectedLine = if ($i -lt $expectedLines.Count) { $expectedLines[$i] } else { "<缺失>" }
        $actualLine = if ($i -lt $actualLines.Count) { $actualLines[$i] } else { "<缺失>" }
        if ($expectedLine -cne $actualLine) {
            $differentLine = $i + 1
            break
        }
    }
    $expectedPreview = if ($expectedLine.Length -gt 120) { $expectedLine.Substring(0, 120) + "..." } else { $expectedLine }
    $actualPreview = if ($actualLine.Length -gt 120) { $actualLine.Substring(0, 120) + "..." } else { $actualLine }
    throw "导入后代码校验失败: $OriginalName，第 $differentLine 行与 UTF-8 源码不一致；期望=[$expectedPreview]；实际=[$actualPreview]"
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
        # UserForm 复用已有设计器，只通过 CP936 临时文件更新代码，避免 AddFromString 转码。
        $formName = "frm$BaseName"
        $tempFile = Join-Path $tempDir "$formName.code.txt"
        [System.IO.File]::WriteAllText($tempFile, $normalized, $vbeEncoding)
        try {
            $existingForm = $null
            foreach ($c in $VBProject.VBComponents) {
                if ($c.Name -eq $formName) { $existingForm = $c; break }
            }
            if ($existingForm) {
                $codeMod = $existingForm.CodeModule
                $totalLines = $codeMod.CountOfLines
                if ($totalLines -gt 0) { $codeMod.DeleteLines(1, $totalLines) }
                $null = $codeMod.AddFromFile($tempFile)
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
                        $null = $importedComponent.CodeModule.AddFromFile($tempFile)
                        $success = $true
                    } catch {
                        if ($retry -eq $maxRetries - 1) { throw $_ }
                    }
                }
            }
        } finally {
            Remove-Item $tempFile -Force -ErrorAction SilentlyContinue
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
        [System.IO.File]::WriteAllText($tempFile, $fileContent, $vbeEncoding)
        try {
            $importedComponent = $VBProject.VBComponents.Import($tempFile)
        } finally {
            Remove-Item $tempFile -Force -ErrorAction SilentlyContinue
        }
    } else {
        # 标准模块：添加 Attribute VB_Name，写入 CP936 临时文件后 Import
        $fileContent = "Attribute VB_Name = `"$BaseName`"`r`n" + $normalized
        $tempFile = Join-Path $tempDir "$BaseName$Extension"
        [System.IO.File]::WriteAllText($tempFile, $fileContent, $vbeEncoding)
        try {
            $importedComponent = $VBProject.VBComponents.Import($tempFile)
        } finally {
            Remove-Item $tempFile -Force -ErrorAction SilentlyContinue
        }
    }

    return $importedComponent
}

# ---------- 导入前严格预检 ----------

Write-Host "`n正在预检 UTF-8、CRLF 与 CP936 兼容性..."
$validatedContent = @{}
$validationErrors = @()
foreach ($file in $filesToImport) {
    try {
        $content = Read-VbaModuleFile -FilePath $file.FullName
        $withoutCrLf = $content.Replace("`r`n", "")
        if ($withoutCrLf.Contains("`r") -or $withoutCrLf.Contains("`n")) {
            throw "存在非 CRLF 行尾"
        }
        $encodingIssues = @(Get-UnencodableCharacters -Content $content)
        if ($encodingIssues.Count -gt 0) {
            throw "包含 CP936 无法表示的字符: $($encodingIssues -join '; ')"
        }
        # 强制执行一次完整转换；与逐字符诊断相互独立，防止遗漏编码器状态问题。
        $null = $vbeEncoding.GetBytes($content)
        $validatedContent[$file.FullName] = $content
        Write-Host "  通过: $($file.Name)" -ForegroundColor DarkGreen
    } catch {
        $validationErrors += "$($file.FullName): $($_.Exception.Message)"
        Write-Host "  失败: $($file.Name) — $($_.Exception.Message)" -ForegroundColor Red
    }
}
if ($validationErrors.Count -gt 0) {
    Write-Host "`n预检失败，尚未打开或修改工作簿：" -ForegroundColor Red
    $validationErrors | ForEach-Object { Write-Host "  - $_" -ForegroundColor Red }
    exit 1
}

# ---------- Excel COM 导入 ----------

Write-Host "`n正在打开工作簿..." -NoNewline

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $wb = $excel.Workbooks.Open((Get-Item $resolvedWorkbook).FullName)
    $vbProject = $wb.VBProject
    if ($null -eq $vbProject) {
        throw "无法访问 VBA 工程。请关闭所有 Excel 窗口，在信任中心启用[信任对 VBA 工程对象模型的访问]，然后重新打开 Excel。"
    }
    Write-Host " 完成"

    # const: VBE component types
    $vbext_ct_StdModule   = 1
    $vbext_ct_ClassModule = 2
    $vbext_ct_MSForm      = 3
    $vbext_ct_Document    = 100   # Sheet/ThisWorkbook — 不操作

    $imported = 0
    $errors = @()

    # 清理待替换模块。全量同步时清理所有用户模块；按组同步时不得删除其他组。
    Write-Host "  清理旧模块..." -NoNewline
    $removedCount = 0
    $builtInNames = @("ThisWorkbook")
    foreach ($comp in $vbProject.VBComponents) {
        if ($comp.Type -eq $vbext_ct_Document) { $builtInNames += $comp.Name }
    }

    $recognizedGroups = @($groups | Where-Object { $moduleMap.ContainsKey($_) } | Select-Object -Unique)
    $isFullSync = ($recognizedGroups.Count -eq $moduleMap.Count)
    $selectedComponentNames = @{}
    if (-not $isFullSync) {
        foreach ($sourceFile in $filesToImport) {
            $info = Get-ImportComponentInfo -File $sourceFile
            $selectedComponentNames[$info.ActualComponentName] = $true
        }
    }

    $toRemove = @()
    foreach ($comp in $vbProject.VBComponents) {
        # UserForm 原地更新；内置模块永不删除。
        if ($comp.Name -in $builtInNames -or $comp.Type -eq $vbext_ct_MSForm) { continue }
        if ($isFullSync -or $selectedComponentNames.ContainsKey($comp.Name)) {
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
        $componentInfo = Get-ImportComponentInfo -File $file
        $baseName = $componentInfo.BaseName
        $ext = $componentInfo.Extension
        $fileGroup = $componentInfo.Group
        $vbaName = $componentInfo.VbaName

        try {
            # 复用预检时严格解码的内容，避免预检与实际导入之间再次读取。
            $content = $validatedContent[$file.FullName]

            # 统一通过 CP936 临时文件导入，并从 VBE 读回逐字校验。
            $label = if ($ext -eq ".frm") { "[UserForm]" } else { "" }
            $importedComponent = Import-VbaModule -VBProject $vbProject -Content $content `
                -BaseName $vbaName -Extension $ext -OriginalName $file.Name
            Assert-ImportedCodeMatches -Component $importedComponent `
                -ExpectedContent $content -OriginalName $file.Name
            Write-Host "  导入并校验: $($file.Name) → $vbaName $label" -ForegroundColor Green

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
