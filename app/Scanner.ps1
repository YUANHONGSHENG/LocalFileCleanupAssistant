param(
    [Parameter(Mandatory = $true)]
    [string]$ConfigPath
)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

# ---------------------------------------------------------------------------
# Native helper: file identity (volume + file index) for hard-link detection.
# This is read-only. It opens the file briefly with shared access.
# ---------------------------------------------------------------------------
$script:HardLinkDetectionAvailable = $false
$script:HardLinkError = ''
if (-not ([System.Management.Automation.PSTypeName]'LocalFileId').Type) {
    try {
        Add-Type -Language CSharp -TypeDefinition @'
using System;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;

public static class LocalFileId
{
    [StructLayout(LayoutKind.Sequential)]
    private struct FILETIME
    {
        public uint dwLowDateTime;
        public uint dwHighDateTime;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct BY_HANDLE_FILE_INFORMATION
    {
        public uint dwFileAttributes;
        public FILETIME ftCreationTime;
        public FILETIME ftLastAccessTime;
        public FILETIME ftLastWriteTime;
        public uint dwVolumeSerialNumber;
        public uint nFileSizeHigh;
        public uint nFileSizeLow;
        public uint nNumberOfLinks;
        public uint nFileIndexHigh;
        public uint nFileIndexLow;
    }

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool GetFileInformationByHandle(
        SafeFileHandle hFile,
        out BY_HANDLE_FILE_INFORMATION lpFileInformation);

    public static string GetIdentity(string path)
    {
        try
        {
            using (var fs = new FileStream(
                path,
                FileMode.Open,
                FileAccess.Read,
                FileShare.ReadWrite | FileShare.Delete,
                4096,
                FileOptions.SequentialScan))
            {
                BY_HANDLE_FILE_INFORMATION info;
                if (!GetFileInformationByHandle(fs.SafeFileHandle, out info))
                    return null;

                return info.dwVolumeSerialNumber.ToString("X8") + ":" +
                       info.nFileIndexHigh.ToString("X8") +
                       info.nFileIndexLow.ToString("X8");
            }
        }
        catch
        {
            return null;
        }
    }
}
'@
        $script:HardLinkDetectionAvailable = $true
    }
    catch {
        $script:HardLinkError = $_.Exception.Message
    }
}
else {
    $script:HardLinkDetectionAvailable = $true
}

# ---------------------------------------------------------------------------
# Configuration and state
# ---------------------------------------------------------------------------
$configPathFull = [IO.Path]::GetFullPath($ConfigPath)
$config = Get-Content -LiteralPath $configPathFull -Raw -Encoding UTF8 | ConvertFrom-Json

$script:ScanId = [string]$config.scanId
$script:Roots = @($config.roots)
$script:RootSet = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
$script:ExcludePaths = @($config.excludePaths)
$script:Thresholds = $config.thresholds
$script:OutputDir = [string]$config.outputDir
$script:StatusPath = [string]$config.statusPath
$script:CancelPath = [string]$config.cancelPath
$script:ResultPath = [string]$config.resultPath
$script:MaxErrors = [int]$config.maxErrors
if ($script:MaxErrors -le 0) { $script:MaxErrors = 2000 }
$script:CacheEnabled = [bool]$config.cacheEnabled
$script:CachePath = [string]$config.cachePath

[void][IO.Directory]::CreateDirectory($script:OutputDir)
$script:StartedAt = (Get-Date).ToUniversalTime()
$script:StartLocal = Get-Date
$script:CancelRequested = $false
$script:LastStatusWrite = [DateTime]::MinValue
$script:AllFiles = [System.Collections.Generic.List[object]]::new()
$script:Errors = [System.Collections.Generic.List[object]]::new()
$script:Notes = [System.Collections.Generic.List[string]]::new()
if (-not $script:HardLinkDetectionAvailable) {
    $script:Notes.Add('硬链接检测未启用；同一物理文件的多个路径会按独立副本显示。原因：' + $script:HardLinkError)
}
$script:FileSeq = 0
$script:HashedBytes = [long]0
$script:Cache = @{}
$script:CacheHits = 0
$script:CacheDirty = $false

function Load-HashCache {
    if (-not $script:CacheEnabled -or [string]::IsNullOrWhiteSpace($script:CachePath)) { return }
    try {
        if (-not [IO.File]::Exists($script:CachePath)) { return }
        $data = Get-Content -LiteralPath $script:CachePath -Raw -Encoding UTF8 | ConvertFrom-Json
        foreach ($e in @($data)) {
            if ($null -eq $e -or [string]::IsNullOrWhiteSpace([string]$e.path)) { continue }
            try {
                $key = ([IO.Path]::GetFullPath([string]$e.path)).TrimEnd('\').ToLowerInvariant()
                $script:Cache[$key] = [ordered]@{
                    size       = [long]$e.size
                    mtimeTicks = [long]$e.mtimeTicks
                    hash       = [string]$e.hash
                }
            }
            catch { }
        }
    }
    catch { $script:Cache = @{} }
}

function Get-CacheHash {
    param([string]$Path, [long]$Size, $Mtime)
    if (-not $script:CacheEnabled -or $script:Cache.Count -eq 0) { return $null }
    try {
        $key = ([IO.Path]::GetFullPath($Path)).TrimEnd('\').ToLowerInvariant()
        if (-not $script:Cache.ContainsKey($key)) { return $null }
        $entry = $script:Cache[$key]
        if ([long]$entry.size -ne $Size) { return $null }
        $ticks = [long]0
        try { $ticks = ([DateTime]$Mtime).ToUniversalTime().Ticks } catch { }
        if ($ticks -le 0 -or [long]$entry.mtimeTicks -ne $ticks) { return $null }
        if ([string]::IsNullOrWhiteSpace([string]$entry.hash)) { return $null }
        return [string]$entry.hash
    }
    catch { return $null }
}

function Set-CacheHash {
    param([string]$Path, [long]$Size, $Mtime, [string]$Hash)
    if (-not $script:CacheEnabled -or [string]::IsNullOrWhiteSpace($script:CachePath)) { return }
    if ([string]::IsNullOrWhiteSpace($Hash)) { return }
    try {
        $key = ([IO.Path]::GetFullPath($Path)).TrimEnd('\').ToLowerInvariant()
        $ticks = [long]0
        try { $ticks = ([DateTime]$Mtime).ToUniversalTime().Ticks } catch { }
        $script:Cache[$key] = [ordered]@{ size = $Size; mtimeTicks = $ticks; hash = $Hash }
        $script:CacheDirty = $true
    }
    catch { }
}

function Save-HashCache {
    if (-not $script:CacheEnabled -or -not $script:CacheDirty -or [string]::IsNullOrWhiteSpace($script:CachePath)) { return }
    try {
        $list = [System.Collections.Generic.List[object]]::new()
        foreach ($k in $script:Cache.Keys) {
            $e = $script:Cache[$k]
            if ([string]::IsNullOrWhiteSpace([string]$e.hash)) { continue }
            $list.Add([pscustomobject]@{
                path       = $k
                size       = [long]$e.size
                mtimeTicks = [long]$e.mtimeTicks
                hash       = [string]$e.hash
            })
        }
        $json = ($list | ConvertTo-Json -Depth 4 -Compress)
        $tmp = "$($script:CachePath).tmp"
        [IO.File]::WriteAllText($tmp, $json, [Text.UTF8Encoding]::new($false))
        if ([IO.File]::Exists($script:CachePath)) { [IO.File]::Delete($script:CachePath) }
        [IO.File]::Move($tmp, $script:CachePath)
        $script:CacheDirty = $false
    }
    catch {
        $script:Notes.Add('哈希缓存保存失败：' + $_.Exception.Message)
    }
}

$script:Status = [ordered]@{
    scanId            = $script:ScanId
    phase             = 'starting'
    message           = '正在初始化扫描'
    currentPath       = ''
    filesScanned      = 0
    dirsScanned       = 0
    bytesScanned      = [long]0
    filesIndexed      = 0
    hashedFiles       = 0
    hashTotal         = 0
    hashedBytes       = [long]0
    errors            = 0
    candidateFiles    = 0
    duplicateGroups   = 0
    duplicateCopies   = 0
    largeFiles        = 0
    tempFiles         = 0
    sameNameGroups    = 0
    versionGroups     = 0
    elapsedSeconds    = 0
    cancelRequested   = $false
    partial           = $false
    resultPath        = ''
    finishedAt        = ''
}

function Write-Status {
    try {
        $script:Status.elapsedSeconds = [math]::Round(((Get-Date) - $script:StartLocal).TotalSeconds, 1)
        $json = $script:Status | ConvertTo-Json -Depth 8 -Compress
        $tmp = "$($script:StatusPath).tmp"
        [IO.File]::WriteAllText($tmp, $json, [Text.UTF8Encoding]::new($false))
        if ([IO.File]::Exists($script:StatusPath)) { [IO.File]::Delete($script:StatusPath) }
        [IO.File]::Move($tmp, $script:StatusPath)
        $script:LastStatusWrite = Get-Date
    }
    catch {
        # Status reporting must never abort a read-only scan.
    }
}

function Test-Cancel {
    if ($script:CancelRequested) { return $true }
    try {
        if ([IO.File]::Exists($script:CancelPath)) {
            $script:CancelRequested = $true
            $script:Status.cancelRequested = $true
            $script:Status.message = '已收到取消请求，正在安全停止'
            Write-Status
            return $true
        }
    }
    catch { }
    return $false
}

function Add-ScanError {
    param(
        [string]$Path,
        [string]$Message,
        [string]$FileId = ''
    )
    $script:Status.errors++
    if ($script:Errors.Count -lt $script:MaxErrors) {
        $script:Errors.Add([pscustomobject]@{
            path   = $Path
            error  = $Message
            time   = (Get-Date).ToUniversalTime().ToString('o')
            fileId = $FileId
        })
    }
}

function Get-LongPath {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return $Path }
    if ($Path.StartsWith('\\?\')) { return $Path }
    if ($Path.StartsWith('\\')) { return '\\?\UNC\' + $Path.Substring(2) }
    return '\\?\' + $Path
}

function Get-AgeDays {
    param($Date)
    if ($null -eq $Date) { return $null }
    try {
        $d = [DateTime]$Date
        if ($d.Year -lt 2000) { return $null }
        $days = ((Get-Date).ToUniversalTime() - $d.ToUniversalTime()).TotalDays
        if ($days -lt 0) { return 0 }
        return [math]::Floor($days)
    }
    catch { return $null }
}

$script:ExcludeDirNames = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
@(
    '.git', '.svn', '.hg', '.bzr',
    'node_modules', 'vendor', 'Pods', 'DerivedData',
    '.venv', 'venv', 'env', '__pycache__', '.pytest_cache', '.mypy_cache', '.ruff_cache', '.tox', '.nox',
    'target', 'build', 'dist', 'out', '.next', '.nuxt', '.gradle', '.idea', '.vs', '.vscode', 'coverage',
    'AppData', 'Application Data', 'Cookies', 'Local Settings', 'Recent', 'SendTo', 'Templates',
    'Favorites', 'Links', 'NetHood', 'PrintHood', 'Searches', 'Contacts', 'Saved Games', 'Library',
    'OneDrive', 'iCloudDrive', 'Dropbox', 'Google Drive', 'Google 云端硬盘',
    'BaiduNetdisk', '百度网盘', '阿里云盘', '坚果云', 'Nutstore', 'Seafile', 'Nextcloud', 'MEGA',
    'WPS Cloud', 'WPS Cloud Files',
    'JianyingPro', 'JianyingPro Drafts', 'IdeaProjects', 'My project', 'PyCharmMiscProject', 'vscode-remote-wsl',
    'Unreal Projects', 'UnrealEngine', 'Unity', 'Unity Projects',
    'Intermediate', 'Saved', 'Binaries', 'DerivedDataCache',
    'Virtual Machines', 'VirtualBox VMs', 'VMware', 'Hyper-V',
    'WeChat Files', 'Tencent Files', 'WhatsApp', 'Telegram Desktop', 'Discord', 'Slack', 'Teams',
    'Thunderbird', 'Outlook', 'Mail', 'My Games', 'SaveGames', 'Saves',
    '1Password', 'KeePass', 'Bitwarden', '.ssh', '.gnupg'
) | ForEach-Object { [void]$script:ExcludeDirNames.Add($_) }

$script:ExcludeDirPatterns = @(
    '*备份*', '*backup*', '*Backup*', '*Backups*',
    '*VirtualBox*', '*VMware*', '*Hyper-V*',
    '*Chrome*', '*Edge*', '*Firefox*', '*BraveSoftware*', '*Opera Software*', '*Vivaldi*',
    '*WeChat*', '*Tencent*', '*WhatsApp*', '*Telegram*', '*Discord*', '*Slack*',
    '*OneDrive*', '*iCloud*', '*Dropbox*', '*Google Drive*', '*百度网盘*', '*阿里云盘*', '*坚果云*', '*WPS Cloud*',
    '*Unreal*', '*Unity*'
)

$script:ExcludeFilePatterns = @(
    '*.pst', '*.ost', '*.mbox', '*.eml', '*.msg',
    '*.kdbx', '*.key', '*.pem', '*.pfx', '*.p12', '*.cer', '*.crt', '*.gpg', '*.asc',
    '*.keystore', '*.jks', '*.vmdk', '*.vdi', '*.vhdx', '*.vhd', '*.qcow2', '*.ova', '*.ovf',
    '*.uproject', '*.uprojectdirs', '*.uasset', '*.umap', '*.ubulk', '*.uexp',
    '*.bak', '*.backup',
    'Bookmarks', 'History', 'Login Data', 'Web Data', 'Cookies'
)

$script:ExcludeExtensions = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
@('.vmdk', '.vdi', '.vhdx', '.vhd', '.qcow2', '.ova', '.ovf', '.pst', '.ost', '.mbox', '.eml', '.msg',
  '.kdbx', '.key', '.pem', '.pfx', '.p12', '.cer', '.crt', '.gpg', '.asc', '.keystore', '.jks', '.bak',
  '.backup', '.uproject', '.uprojectdirs', '.uasset', '.umap', '.ubulk', '.uexp') |
  ForEach-Object { [void]$script:ExcludeExtensions.Add($_) }

$script:FinancialRegex = '(财务|发票|税务|银行|账单|报销|会计|工资|存款|投资|保险|贷款|预算|Finance|Invoice|Tax|Bank|Statement|Payroll|Accounting|Budget)'
$script:SecretRegex = '(?i)(^|[._\-\s])(id_rsa|id_ed25519|credentials?|secrets?|password|passwd|token)([._\-\s]|$)|^\.env(\.|$)'

function Test-ExcludedPathPrefix {
    param([string]$Path)
    $p = $Path.TrimEnd('\')
    foreach ($ex in $script:ExcludePaths) {
        if ([string]::IsNullOrWhiteSpace($ex)) { continue }
        $e = ([string]$ex).TrimEnd('\')
        if ($p.Equals($e, [StringComparison]::OrdinalIgnoreCase)) { return $true }
        if ($p.StartsWith($e + '\', [StringComparison]::OrdinalIgnoreCase)) { return $true }
    }
    return $false
}

function Test-ProjectDirectory {
    param([string]$Path)
    foreach ($marker in @('.git', '.svn', '.hg')) {
        try { if ([IO.Directory]::Exists((Join-Path $Path $marker))) { return $true } } catch { }
    }
    foreach ($marker in @('package.json', 'pyproject.toml', 'Cargo.toml', 'go.mod', 'pom.xml', 'build.gradle',
                          'build.gradle.kts', '*.sln', '*.csproj', '*.uproject', '*.uprojectdirs', '*.unity',
                          'project.godot', '*.rproj')) {
        try {
            if ($marker.Contains('*')) {
                if ([IO.Directory]::EnumerateFiles($Path, $marker, [IO.SearchOption]::TopDirectoryOnly).GetEnumerator().MoveNext()) { return $true }
            }
            elseif ([IO.File]::Exists((Join-Path $Path $marker))) { return $true }
        } catch { }
    }
    return $false
}

function Test-ExcludedDirectory {
    param(
        [string]$Path,
        [bool]$IsRoot = $false
    )

    if (Test-ExcludedPathPrefix $Path) { return $true }

    $name = [IO.Path]::GetFileName($Path.TrimEnd('\'))
    if ([string]::IsNullOrWhiteSpace($name)) { return $false }

    if (-not $IsRoot) {
        if ($name.StartsWith('.')) { return $true }
        if ($script:ExcludeDirNames.Contains($name)) { return $true }
        foreach ($pat in $script:ExcludeDirPatterns) {
            if ($name -like $pat) { return $true }
        }
        if ($name -match $script:FinancialRegex) { return $true }
        if (Test-ProjectDirectory $Path) { return $true }
    }

    try {
        $attrs = [IO.File]::GetAttributes((Get-LongPath $Path))
        if (($attrs -band [IO.FileAttributes]::ReparsePoint) -ne 0) { return $true }
        if (($attrs -band [IO.FileAttributes]::Hidden) -ne 0 -and -not $IsRoot) { return $true }
        if (($attrs -band [IO.FileAttributes]::System) -ne 0 -and -not $IsRoot) { return $true }
        if (($attrs -band [IO.FileAttributes]::Offline) -ne 0) { return $true }
        if (([int]$attrs -band 0x400000) -ne 0) { return $true } # RECALL_ON_DATA_ACCESS
    }
    catch { }

    return $false
}

function Test-ExcludedFile {
    param([string]$Path)

    $name = [IO.Path]::GetFileName($Path)
    $ext = [IO.Path]::GetExtension($Path).ToLowerInvariant()

    if ($script:ExcludeExtensions.Contains($ext)) { return '排除的扩展名' }
    foreach ($pat in $script:ExcludeFilePatterns) {
        if ($name -like $pat) { return '排除的文件类型' }
    }
    if ($name -match $script:FinancialRegex) { return '财务/敏感名称' }
    if ($name -match $script:SecretRegex) { return '密钥/凭据名称' }
    return $null
}

function Test-FileUnlocked {
    param([string]$Path)
    try {
        $fs = [IO.File]::Open((Get-LongPath $Path), [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::None)
        $fs.Dispose()
        return $true
    }
    catch { return $false }
}

function Get-FileIdentitySafe {
    param([string]$Path)
    if (-not $script:HardLinkDetectionAvailable) { return $null }
    try {
        $id = [LocalFileId]::GetIdentity((Get-LongPath $Path))
        if ([string]::IsNullOrWhiteSpace($id)) { return $null }
        return $id
    }
    catch { return $null }
}

function ConvertTo-HexLower {
    param([byte[]]$Bytes)
    $sb = [Text.StringBuilder]::new($Bytes.Length * 2)
    foreach ($b in $Bytes) { [void]$sb.Append($b.ToString('x2')) }
    return $sb.ToString()
}

function Get-Sha256Hex {
    param(
        [string]$Path,
        [long]$Length = 0
    )
    $sha = $null
    $fs = $null
    try {
        $sha = [Security.Cryptography.SHA256]::Create()
        $fs = [IO.FileStream]::new(
            (Get-LongPath $Path),
            [IO.FileMode]::Open,
            [IO.FileAccess]::Read,
            ([IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete),
            1048576,
            [IO.FileOptions]::SequentialScan)

        $remaining = $Length
        $buffer = New-Object byte[] 1048576
        $lastProgress = Get-Date
        while ($true) {
            if (Test-Cancel) { throw 'CANCELED' }
            $toRead = $buffer.Length
            if ($Length -gt 0) {
                if ($remaining -le 0) { break }
                if ($remaining -lt $toRead) { $toRead = [int]$remaining }
            }
            $read = $fs.Read($buffer, 0, $toRead)
            if ($read -le 0) { break }
            [void]$sha.TransformBlock($buffer, 0, $read, $null, 0)
            $script:HashedBytes += $read
            if ($Length -gt 0) { $remaining -= $read }
            $now = Get-Date
            if (($now - $lastProgress).TotalMilliseconds -ge 700) {
                $script:Status.hashedBytes = $script:HashedBytes
                $script:Status.currentPath = $Path
                Write-Status
                $lastProgress = $now
            }
        }
        [void]$sha.TransformFinalBlock((New-Object byte[] 0), 0, 0)
        return (ConvertTo-HexLower -Bytes $sha.Hash)
    }
    finally {
        if ($fs) { $fs.Dispose() }
        if ($sha) { $sha.Dispose() }
    }
}

function Get-TypeText {
    param([string]$Ext)
    switch ($Ext.ToLowerInvariant()) {
        '.pdf'  { 'PDF' }
        '.doc'  { 'Word' }
        '.docx' { 'Word' }
        '.xls'  { 'Excel' }
        '.xlsx' { 'Excel' }
        '.ppt'  { 'PowerPoint' }
        '.pptx' { 'PowerPoint' }
        '.txt'  { '文本' }
        '.md'   { 'Markdown' }
        '.rtf'  { 'RTF' }
        '.csv'  { 'CSV' }
        '.jpg'  { '图片' }
        '.jpeg' { '图片' }
        '.png'  { '图片' }
        '.gif'  { '图片' }
        '.bmp'  { '图片' }
        '.webp' { '图片' }
        '.heic' { '图片' }
        '.tif'  { '图片' }
        '.tiff' { '图片' }
        '.zip'  { '压缩包' }
        '.7z'   { '压缩包' }
        '.rar'  { '压缩包' }
        '.tar'  { '压缩包' }
        '.gz'   { '压缩包' }
        '.mp4'  { '视频' }
        '.mov'  { '视频' }
        '.mkv'  { '视频' }
        '.avi'  { '视频' }
        '.wmv'  { '视频' }
        '.mp3'  { '音频' }
        '.wav'  { '音频' }
        '.flac' { '音频' }
        '.m4a'  { '音频' }
        '.exe'  { '安装程序/可执行文件' }
        '.msi'  { '安装包' }
        '.iso'  { '光盘镜像' }
        '.dmg'  { '磁盘镜像' }
        default {
            if ([string]::IsNullOrWhiteSpace($Ext)) { '无扩展名' }
            else { $Ext.TrimStart('.').ToUpperInvariant() + ' 文件' }
        }
    }
}

function Get-PurposeText {
    param([string]$Ext)
    switch ($Ext.ToLowerInvariant()) {
        '.pdf'  { 'PDF 文档；具体用途需打开确认' }
        '.doc'  { 'Word 文档；具体用途需打开确认' }
        '.docx' { 'Word 文档；具体用途需打开确认' }
        '.xls'  { 'Excel 表格；具体用途需打开确认' }
        '.xlsx' { 'Excel 表格；具体用途需打开确认' }
        '.ppt'  { '演示文稿；具体用途需打开确认' }
        '.pptx' { '演示文稿；具体用途需打开确认' }
        '.txt'  { '文本文件；内容无法确认' }
        '.md'   { 'Markdown 文本；内容无法确认' }
        '.rtf'  { 'RTF 文档；具体用途需打开确认' }
        '.csv'  { 'CSV 表格；具体用途需打开确认' }
        '.jpg'  { '图片；内容无法确认' }
        '.jpeg' { '图片；内容无法确认' }
        '.png'  { '图片；内容无法确认' }
        '.gif'  { '图片；内容无法确认' }
        '.bmp'  { '图片；内容无法确认' }
        '.webp' { '图片；内容无法确认' }
        '.heic' { '图片；内容无法确认' }
        '.tif'  { '图片；内容无法确认' }
        '.tiff' { '图片；内容无法确认' }
        '.zip'  { '压缩包；内容无法确认' }
        '.7z'   { '压缩包；内容无法确认' }
        '.rar'  { '压缩包；内容无法确认' }
        '.tar'  { '压缩包；内容无法确认' }
        '.gz'   { '压缩包；内容无法确认' }
        '.mp4'  { '视频；内容无法确认' }
        '.mov'  { '视频；内容无法确认' }
        '.mkv'  { '视频；内容无法确认' }
        '.avi'  { '视频；内容无法确认' }
        '.wmv'  { '视频；内容无法确认' }
        '.mp3'  { '音频；内容无法确认' }
        '.wav'  { '音频；内容无法确认' }
        '.flac' { '音频；内容无法确认' }
        '.m4a'  { '音频；内容无法确认' }
        '.exe'  { '安装程序/可执行文件；来源和用途需确认' }
        '.msi'  { '安装包；可能可重新下载，需确认' }
        '.iso'  { '光盘镜像；用途无法确认' }
        '.dmg'  { '磁盘镜像；用途无法确认' }
        default { '无法确认' }
    }
}

function Normalize-Stem {
    param([string]$Name)
    $stem = [IO.Path]::GetFileNameWithoutExtension($Name).ToLowerInvariant()
    $patterns = @(
        '\s*[\(\[]\d+[\)\]]$',
        '\s*[-_ ]?(copy|副本|备份|final|final\d*|v\d+|version\d+|rev\d*|修订|修改版)$',
        '\s*[-_ ]?20\d{2}[-_.]?\d{2}[-_.]?\d{2}$',
        '\s*[-_ ]?\d{8}$'
    )
    do {
        $changed = $false
        foreach ($p in $patterns) {
            $new = [regex]::Replace($stem, $p, '', [Text.RegularExpressions.RegexOptions]::IgnoreCase)
            if ($new -ne $stem) { $stem = $new; $changed = $true }
        }
    } while ($changed)
    return $stem.Trim(' ', '-', '_', '.')
}

# ---------------------------------------------------------------------------
# Enumerate files without following links or entering excluded trees.
# ---------------------------------------------------------------------------
Write-Status
$script:Status.phase = 'enumerating'
$script:Status.message = '正在枚举文件（只读取元数据）'
Write-Status

$largeBytes = [long]$script:Thresholds.largeBytes
$modifiedDays = [int]$script:Thresholds.modifiedDays
$accessedDays = [int]$script:Thresholds.accessedDays
$tempDays = [int]$script:Thresholds.tempDays

$stack = [System.Collections.Generic.Stack[string]]::new()
foreach ($root in $script:Roots) {
    if ([string]::IsNullOrWhiteSpace($root)) { continue }
    try {
        $r = [IO.Path]::GetFullPath([string]$root).TrimEnd('\')
        if (-not [IO.Directory]::Exists((Get-LongPath $r))) {
            Add-ScanError -Path $r -Message '目录不存在或不可访问'
            continue
        }
        if (Test-ExcludedPathPrefix $r) {
            $script:Notes.Add("根目录被排除：$r")
            continue
        }
        [void]$script:RootSet.Add($r)
        $stack.Push($r)
    }
    catch {
        Add-ScanError -Path ([string]$root) -Message ("无法解析根目录：" + $_.Exception.Message)
    }
}

while ($stack.Count -gt 0) {
    if (Test-Cancel) { break }

    $dir = $stack.Pop()
    $isRootDir = $script:RootSet.Contains($dir)
    if (Test-ExcludedDirectory -Path $dir -IsRoot $isRootDir) { continue }

    $script:Status.currentPath = $dir
    $items = $null
    try {
        $di = [IO.DirectoryInfo]::new((Get-LongPath $dir))
        $items = $di.EnumerateFileSystemInfos()
    }
    catch {
        Add-ScanError -Path $dir -Message ("目录无法读取：" + $_.Exception.Message)
        continue
    }

    try {
        foreach ($item in $items) {
            if (Test-Cancel) { break }
            try {
                $attrs = $item.Attributes

                if (($attrs -band [IO.FileAttributes]::ReparsePoint) -ne 0) { continue }
                if (($attrs -band [IO.FileAttributes]::Offline) -ne 0) { continue }
                if (([int]$attrs -band 0x400000) -ne 0) { continue }

                $isDir = (($attrs -band [IO.FileAttributes]::Directory) -ne 0)
                $full = $item.FullName
                if ($full.StartsWith('\\?\')) { $full = $full.Substring(4) }

                if ($isDir) {
                    if (Test-ExcludedDirectory -Path $full -IsRoot $false) { continue }
                    $stack.Push($full)
                    continue
                }

                $name = $item.Name
                $ext = [IO.Path]::GetExtension($name).ToLowerInvariant()
                $isTempPattern = ($name -like '*.tmp' -or $name -like '*.temp' -or $name -like '*.part' -or
                                  $name -like '*.partial' -or $name -like '*.crdownload' -or
                                  $name -like '*.download' -or $name -like '~$*')

                # Hidden/system files are skipped, except stale Office lock files
                # which we may want to show under the temp category.
                if (-not $isTempPattern) {
                    if (($attrs -band [IO.FileAttributes]::Hidden) -ne 0) { continue }
                    if (($attrs -band [IO.FileAttributes]::System) -ne 0) { continue }
                }

                $excluded = Test-ExcludedFile -Path $full
                if ($excluded) { continue }

                $fi = [IO.FileInfo]::new((Get-LongPath $full))
                $size = [long]$fi.Length
                if ($size -le 0) { continue }

                $script:FileSeq++
                $f = [pscustomobject]@{
                    Id                = ('F{0:D6}' -f $script:FileSeq)
                    Path              = $full
                    Name              = $name
                    Ext               = $ext
                    Size              = $size
                    ModifiedUtc       = $fi.LastWriteTimeUtc
                    AccessedUtc       = $fi.LastAccessTimeUtc
                    Hash              = ''
                    Identity          = ''
                    IsHardLink        = $false
                    IsDuplicate       = $false
                    IsLarge           = $false
                    IsTemp            = $false
                    IsSameNameDiff    = $false
                    IsVersion         = $false
                    IsError           = $false
                    CanSelect         = $true
                    GroupId           = ''
                    GroupSize         = [long]0
                    GroupPaths        = 0
                    GroupDistinct     = 0
                    Error             = ''
                    Purpose           = ''
                    TypeText          = ''
                    Categories        = ''
                    Reason            = ''
                    Risk              = ''
                    LargeAccessRecent = $false
                }

                $script:AllFiles.Add($f)
                $script:Status.filesScanned++
                $script:Status.filesIndexed++
                $script:Status.bytesScanned += $size

                if (($script:Status.filesScanned % 500) -eq 0) {
                    $script:Status.hashedBytes = $script:HashedBytes
                    Write-Status
                }
            }
            catch {
                Add-ScanError -Path ([string]$item.FullName) -Message ("读取文件元数据失败：" + $_.Exception.Message)
            }
        }
    }
    catch {
        Add-ScanError -Path $dir -Message ("枚举目录时失败：" + $_.Exception.Message)
    }

    $script:Status.dirsScanned++
    if (((Get-Date) - $script:LastStatusWrite).TotalMilliseconds -ge 700) { Write-Status }
}

if ($script:CancelRequested) {
    if ($script:CacheEnabled) { Save-HashCache }
    $script:Status.phase = 'canceled'
    $script:Status.message = '扫描已取消，未生成完整候选清单'
    $script:Status.partial = $true
    $result = [ordered]@{
        schema     = 1
        scanId     = $script:ScanId
        startedAt  = $script:StartedAt.ToString('o')
        finishedAt = (Get-Date).ToUniversalTime().ToString('o')
        canceled   = $true
        partial    = $true
        roots      = $script:Roots
        thresholds = $script:Thresholds
        stats      = $script:Status
        groups     = @()
        files      = @()
        errors     = $script:Errors
        notes      = $script:Notes
    }
    [IO.File]::WriteAllText($script:ResultPath, ($result | ConvertTo-Json -Depth 10 -Compress), [Text.UTF8Encoding]::new($false))
    $script:Status.resultPath = $script:ResultPath
    $script:Status.finishedAt = (Get-Date).ToUniversalTime().ToString('o')
    Write-Status
    exit 0
}

# ---------------------------------------------------------------------------
# Hash candidates (only same-size groups and same-name groups).
# ---------------------------------------------------------------------------
Load-HashCache
if ($script:CacheEnabled) {
    $script:Notes.Add("哈希缓存已启用：$($script:CachePath)")
}
$script:Status.phase = 'hashing'
$script:Status.message = '正在按大小初筛并计算 SHA-256'
Write-Status

$bySize = @{}
foreach ($f in $script:AllFiles) {
    if ($f.IsError -or $f.Size -le 0) { continue }
    $key = [string]$f.Size
    if (-not $bySize.ContainsKey($key)) {
        $bySize[$key] = [System.Collections.Generic.List[object]]::new()
    }
    $bySize[$key].Add($f)
}

$sizeGroups = @($bySize.Values | Where-Object { $_.Count -ge 2 })
$hashTotal = 0
foreach ($g in $sizeGroups) { $hashTotal += $g.Count }
$script:Status.hashTotal = $hashTotal

$prefixBytes = 1048576
foreach ($g in $sizeGroups) {
    $prefixByF = @{}
    foreach ($f in $g) {
        if (Test-Cancel) { break }
        if ($f.IsError -or $f.Hash) { continue }
        $cached = Get-CacheHash -Path $f.Path -Size $f.Size -Mtime $f.ModifiedUtc
        if ($cached) {
            $f.Hash = $cached
            $script:CacheHits++
            continue
        }
        try {
            $prefixByF[$f.Id] = Get-Sha256Hex -Path $f.Path -Length $prefixBytes
            $script:Status.hashedBytes = $script:HashedBytes
        }
        catch {
            if ($script:CancelRequested) { break }
            $f.IsError = $true
            $f.CanSelect = $false
            $f.Error = $_.Exception.Message
            Add-ScanError -Path $f.Path -Message ("SHA-256 前缀计算失败：" + $_.Exception.Message) -FileId $f.Id
        }
    }
    if ($script:CancelRequested) { break }

    $prefixGroups = @(
        $g |
        Where-Object { -not $_.IsError -and [string]::IsNullOrWhiteSpace([string]$_.Hash) } |
        Group-Object -Property { $prefixByF[$_.Id] } |
        Where-Object { $_.Count -ge 2 -and -not [string]::IsNullOrWhiteSpace([string]$_.Name) }
    )
    foreach ($pg in $prefixGroups) {
        foreach ($f in $pg.Group) {
            if (Test-Cancel) { break }
            if ($f.IsError -or $f.Hash) { continue }
            try {
                $f.Hash = Get-Sha256Hex -Path $f.Path
                Set-CacheHash -Path $f.Path -Size $f.Size -Mtime $f.ModifiedUtc -Hash $f.Hash
                $script:Status.hashedFiles++
                $script:Status.hashedBytes = $script:HashedBytes
            }
            catch {
                if ($script:CancelRequested) { break }
                $f.IsError = $true
                $f.CanSelect = $false
                $f.Error = $_.Exception.Message
                Add-ScanError -Path $f.Path -Message ("SHA-256 计算失败：" + $_.Exception.Message) -FileId $f.Id
            }
        }
        if ($script:CancelRequested) { break }
    }
    if ($script:CancelRequested) { break }
    Write-Status
}

# Same-name groups may have different sizes; hash them too.
$nameGroups = @($script:AllFiles | Where-Object { -not $_.IsError } | Group-Object -Property { $_.Name.ToLowerInvariant() } | Where-Object { $_.Count -ge 2 })
foreach ($g in $nameGroups) {
    foreach ($f in $g.Group) {
        if (Test-Cancel) { break }
        if ($f.IsError -or $f.Hash) { continue }
        $cached = Get-CacheHash -Path $f.Path -Size $f.Size -Mtime $f.ModifiedUtc
        if ($cached) {
            $f.Hash = $cached
            $script:CacheHits++
            continue
        }
        try {
            $f.Hash = Get-Sha256Hex -Path $f.Path
            Set-CacheHash -Path $f.Path -Size $f.Size -Mtime $f.ModifiedUtc -Hash $f.Hash
            $script:Status.hashedFiles++
            $script:Status.hashedBytes = $script:HashedBytes
        }
        catch {
            if ($script:CancelRequested) { break }
            $f.IsError = $true
            $f.CanSelect = $false
            $f.Error = $_.Exception.Message
            Add-ScanError -Path $f.Path -Message ("SHA-256 计算失败：" + $_.Exception.Message) -FileId $f.Id
        }
    }
    if ($script:CancelRequested) { break }
    Write-Status
}

if ($script:CancelRequested) {
    if ($script:CacheEnabled) { Save-HashCache }
    $script:Status.phase = 'canceled'
    $script:Status.message = '扫描已取消，未生成完整候选清单'
    $script:Status.partial = $true
    $result = [ordered]@{
        schema     = 1
        scanId     = $script:ScanId
        startedAt  = $script:StartedAt.ToString('o')
        finishedAt = (Get-Date).ToUniversalTime().ToString('o')
        canceled   = $true
        partial    = $true
        roots      = $script:Roots
        thresholds = $script:Thresholds
        stats      = $script:Status
        groups     = @()
        files      = @()
        errors     = $script:Errors
        notes      = $script:Notes
    }
    [IO.File]::WriteAllText($script:ResultPath, ($result | ConvertTo-Json -Depth 10 -Compress), [Text.UTF8Encoding]::new($false))
    $script:Status.resultPath = $script:ResultPath
    $script:Status.finishedAt = (Get-Date).ToUniversalTime().ToString('o')
    Write-Status
    exit 0
}

# ---------------------------------------------------------------------------
# Duplicate groups: same size + same SHA-256.
# Hard links are identified and flagged; they are not silently treated as
# independent copies.
# ---------------------------------------------------------------------------
$script:Status.phase = 'analyzing'
$script:Status.message = '正在分析重复组、大文件和时间阈值'
Write-Status

$byHash = @{}
foreach ($f in $script:AllFiles) {
    if ($f.IsError -or [string]::IsNullOrWhiteSpace($f.Hash)) { continue }
    $key = ([string]$f.Size) + '|' + $f.Hash
    if (-not $byHash.ContainsKey($key)) {
        $byHash[$key] = [System.Collections.Generic.List[object]]::new()
    }
    $byHash[$key].Add($f)
}

$dupGroups = [System.Collections.Generic.List[object]]::new()
foreach ($kv in $byHash.GetEnumerator()) {
    $list = $kv.Value
    if ($list.Count -lt 2) { continue }

    $idCounts = @{}
    foreach ($f in $list) {
        $id = Get-FileIdentitySafe -Path $f.Path
        if ([string]::IsNullOrWhiteSpace($id)) { $id = 'UNKNOWN:' + $f.Id }
        $f.Identity = $id
        if (-not $idCounts.ContainsKey($id)) { $idCounts[$id] = 0 }
        $idCounts[$id]++
    }
    foreach ($f in $list) {
        $f.IsHardLink = ($idCounts[$f.Identity] -gt 1)
    }

    $distinct = $idCounts.Count
    $dupGroups.Add([pscustomobject]@{
        Files      = $list
        Size       = [long]$list[0].Size
        Hash       = [string]$list[0].Hash
        Distinct   = [int]$distinct
        TotalPaths = [int]$list.Count
    })
}

$dupGroupsSorted = @($dupGroups | Sort-Object -Property Size -Descending)
$groupSeq = 0
foreach ($g in $dupGroupsSorted) {
    $groupSeq++
    $gid = ('G{0:D4}' -f $groupSeq)
    foreach ($f in $g.Files) {
        $f.IsDuplicate = $true
        $f.GroupId = $gid
        $f.GroupSize = $g.Size
        $f.GroupPaths = $g.TotalPaths
        $f.GroupDistinct = $g.Distinct
    }
    $script:Status.duplicateGroups++
    $script:Status.duplicateCopies += $g.TotalPaths
}

# ---------------------------------------------------------------------------
# Same name, different content.
# ---------------------------------------------------------------------------
$sameNameCount = 0
foreach ($g in $nameGroups) {
    $hashes = @($g.Group | Where-Object { -not $_.IsError -and -not [string]::IsNullOrWhiteSpace($_.Hash) } | ForEach-Object { $_.Hash } | Sort-Object -Unique)
    if ($hashes.Count -ge 2) {
        $sameNameCount++
        foreach ($f in $g.Group) {
            if (-not $f.IsError) { $f.IsSameNameDiff = $true }
        }
    }
}
$script:Status.sameNameGroups = $sameNameCount

# ---------------------------------------------------------------------------
# Possible historical versions (heuristic only, never a deletion suggestion).
# ---------------------------------------------------------------------------
$versionGroups = @($script:AllFiles | Where-Object { -not $_.IsError } | Group-Object -Property { (Normalize-Stem $_.Name) + '|' + $_.Ext } | Where-Object { $_.Count -ge 2 })
$versionCount = 0
foreach ($g in $versionGroups) {
    if ($g.Count -lt 2) { continue }
    $hashes = @($g.Group | Where-Object { -not [string]::IsNullOrWhiteSpace($_.Hash) } | ForEach-Object { $_.Hash } | Sort-Object -Unique)
    if ($hashes.Count -le 1) { continue }

    $sorted = @($g.Group | Sort-Object -Property ModifiedUtc)
    $looksLikeVersion = $false
    for ($i = 1; $i -lt $sorted.Count; $i++) {
        $a = $sorted[$i - 1]
        $b = $sorted[$i]
        $days = [math]::Abs((([DateTime]$b.ModifiedUtc) - ([DateTime]$a.ModifiedUtc)).TotalDays)
        if ($days -gt 180) { continue }
        $minSize = [math]::Min([double]$a.Size, [double]$b.Size)
        if ($minSize -le 0) { continue }
        $ratio = [math]::Max([double]$a.Size, [double]$b.Size) / $minSize
        if ($ratio -le 1.25) { $looksLikeVersion = $true; break }
    }
    if ($looksLikeVersion) {
        $versionCount++
        foreach ($f in $g.Group) {
            if (-not $f.IsError) { $f.IsVersion = $true }
        }
    }
}
$script:Status.versionGroups = $versionCount

# ---------------------------------------------------------------------------
# Large stale files and recognizable temp files.
# ---------------------------------------------------------------------------
foreach ($f in $script:AllFiles) {
    if ($f.IsError) { continue }
    $ageMod = Get-AgeDays -Date $f.ModifiedUtc
    $ageAcc = Get-AgeDays -Date $f.AccessedUtc

    if ($f.Size -ge $largeBytes -and $null -ne $ageMod -and $ageMod -ge $modifiedDays) {
        $f.IsLarge = $true
        if ($null -ne $ageAcc -and $ageAcc -lt $accessedDays) {
            $f.LargeAccessRecent = $true
        }
    }

    $name = $f.Name
    $isTempPattern = ($name -like '*.tmp' -or $name -like '*.temp' -or $name -like '*.part' -or
                      $name -like '*.partial' -or $name -like '*.crdownload' -or
                      $name -like '*.download' -or $name -like '~$*')
    if ($isTempPattern -and $null -ne $ageMod -and $ageMod -ge $tempDays) {
        if (Test-FileUnlocked -Path $f.Path) {
            $f.IsTemp = $true
        }
        else {
            $f.IsError = $true
            $f.CanSelect = $false
            $f.Error = '文件被占用，未列入临时文件'
            Add-ScanError -Path $f.Path -Message $f.Error -FileId $f.Id
        }
    }
}

# ---------------------------------------------------------------------------
# Build candidate output rows.
# ---------------------------------------------------------------------------
$candidateRows = [System.Collections.Generic.List[object]]::new()
$largeCount = 0
$tempCount = 0

foreach ($f in $script:AllFiles) {
    $cats = [System.Collections.Generic.List[string]]::new()
    $reasons = [System.Collections.Generic.List[string]]::new()
    $riskScore = 1

    if ($f.IsError) {
        $cats.Add('跳过/错误')
        $reasons.Add('扫描/读取失败：' + $f.Error)
        $riskScore = 4
    }

    if ($f.IsDuplicate) {
        $cats.Add('完全重复')
        if ($f.GroupDistinct -ge 2) {
            $reasons.Add("完全重复（SHA-256 一致），同组 $($f.GroupPaths) 个路径 / $($f.GroupDistinct) 个物理副本")
            $riskScore = [math]::Max($riskScore, 1)
        }
        else {
            $reasons.Add("同一物理文件的多个硬链接路径（同组 $($f.GroupPaths) 个路径），仅删除部分路径不会释放空间")
            $riskScore = [math]::Max($riskScore, 3)
        }
    }

    if ($f.IsLarge) {
        $cats.Add('长期未使用的大文件')
        $largeCount++
        $ageMod = Get-AgeDays -Date $f.ModifiedUtc
        $ageAcc = Get-AgeDays -Date $f.AccessedUtc
        if ($f.LargeAccessRecent) {
            $reasons.Add("长期未修改（$ageMod 天），但最近访问过（$ageAcc 天）；可能仍在使用")
            $riskScore = [math]::Max($riskScore, 3)
        }
        elseif ($null -eq $ageAcc) {
            $reasons.Add("长期未修改（$ageMod 天），最近访问时间不可用；访问时间仅供参考")
            $riskScore = [math]::Max($riskScore, 2)
        }
        else {
            $reasons.Add("长期未修改且未访问（修改 $ageMod 天，访问 $ageAcc 天）；建议考虑删除，但请先确认用途")
            $riskScore = [math]::Max($riskScore, 2)
        }
    }

    if ($f.IsTemp) {
        $cats.Add('临时文件')
        $tempCount++
        $reasons.Add('超过 30 天的可识别临时文件/下载残留，且当前未被占用')
        $riskScore = [math]::Max($riskScore, 1)
    }

    if ($f.IsSameNameDiff) {
        $cats.Add('同名不同内容')
        $reasons.Add('同名但内容不同（SHA-256 不同），不是重复文件，请人工确认')
        $riskScore = [math]::Max($riskScore, 3)
    }

    if ($f.IsVersion) {
        $cats.Add('可能的历史版本')
        $reasons.Add('可能的历史版本（仅按名称/时间/大小推测，未做内容相似度判断），不建议自动删除')
        $riskScore = [math]::Max($riskScore, 3)
    }

    if ($cats.Count -eq 0) { continue }

    $f.Purpose = Get-PurposeText -Ext $f.Ext
    $f.TypeText = Get-TypeText -Ext $f.Ext
    $f.Categories = ($cats | Select-Object -Unique) -join '、'
    $f.Reason = ($reasons | Select-Object -Unique) -join '；'

    if ($f.IsError -or -not $f.CanSelect) {
        $riskScore = 4
    }
    $f.Risk = switch ($riskScore) {
        4 { '不建议' }
        3 { '高' }
        2 { '中' }
        default { '低' }
    }

    if ($f.IsSameNameDiff -or $f.IsVersion) {
        $f.CanSelect = $true
    }
    if ($f.IsError) {
        $f.CanSelect = $false
    }

    $candidateRows.Add([ordered]@{
        id               = $f.Id
        path             = $f.Path
        name             = $f.Name
        ext              = $f.Ext
        typeText         = $f.TypeText
        size             = [long]$f.Size
        modifiedUtc      = $f.ModifiedUtc.ToString('o')
        accessedUtc      = $f.AccessedUtc.ToString('o')
        purpose          = $f.Purpose
        categories       = $f.Categories
        reason           = $f.Reason
        risk             = $f.Risk
        canSelect        = [bool]$f.CanSelect
        isDuplicate      = [bool]$f.IsDuplicate
        isLarge          = [bool]$f.IsLarge
        isTemp           = [bool]$f.IsTemp
        isSameNameDiff   = [bool]$f.IsSameNameDiff
        isVersion        = [bool]$f.IsVersion
        isError          = [bool]$f.IsError
        isHardLink       = [bool]$f.IsHardLink
        groupId          = $f.GroupId
        groupSize        = [long]$f.GroupSize
        groupPaths       = [int]$f.GroupPaths
        groupDistinct    = [int]$f.GroupDistinct
        hash             = $f.Hash
        identity         = $f.Identity
        error            = $f.Error
    })
}

$script:Status.largeFiles = $largeCount
$script:Status.tempFiles = $tempCount
$script:Status.candidateFiles = $candidateRows.Count

$groupsOut = [System.Collections.Generic.List[object]]::new()
foreach ($g in $dupGroupsSorted) {
    $groupsOut.Add([ordered]@{
        groupId    = $g.Files[0].GroupId
        hash       = $g.Hash
        size       = [long]$g.Size
        distinct   = [int]$g.Distinct
        totalPaths = [int]$g.TotalPaths
        fileIds    = @($g.Files | ForEach-Object { $_.Id })
        paths      = @($g.Files | ForEach-Object { $_.Path })
    })
}

if ($script:CacheEnabled) {
    Save-HashCache
    $script:Notes.Add("哈希缓存：本次命中 $($script:CacheHits) 个，缓存条目 $($script:Cache.Count) 条")
}

$finishedAt = (Get-Date).ToUniversalTime()
$script:Status.phase = 'done'
$script:Status.message = '扫描与分析完成'
$script:Status.partial = $false
$script:Status.finishedAt = $finishedAt.ToString('o')
$script:Status.resultPath = $script:ResultPath
$script:Status.elapsedSeconds = [math]::Round(($finishedAt - $script:StartedAt).TotalSeconds, 1)

$result = [ordered]@{
    schema     = 1
    scanId     = $script:ScanId
    startedAt  = $script:StartedAt.ToString('o')
    finishedAt = $finishedAt.ToString('o')
    canceled   = $false
    partial    = $false
    roots      = $script:Roots
    thresholds = $script:Thresholds
    stats      = $script:Status
    groups     = $groupsOut
    files      = $candidateRows
    errors     = $script:Errors
    notes      = $script:Notes
}

[IO.File]::WriteAllText($script:ResultPath, ($result | ConvertTo-Json -Depth 12 -Compress), [Text.UTF8Encoding]::new($false))
Write-Status
exit 0




























