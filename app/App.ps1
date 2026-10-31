param(
    [switch]$SmokeTest,
    [string]$TestScanDir,
    [switch]$TestDialogs,
    [switch]$LoadLatest,
    [switch]$PreselectDuplicates,
    [string]$LoadScanDir,
    [switch]$TestPreview,
    [switch]$TestBulk,
    [switch]$TestCleanup,
    [switch]$TestPermanent,
    [switch]$TestTrace,
    [switch]$TestFolder
)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName Microsoft.VisualBasic

$script:AppRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$script:BaseDir = Split-Path -Parent $script:AppRoot
$script:DataDir = Join-Path $script:BaseDir 'data'
$script:LogDir = Join-Path $script:BaseDir 'logs'
$script:QuarantineDir = Join-Path $script:BaseDir 'quarantine'
$script:ScannerPath = Join-Path $script:AppRoot 'Scanner.ps1'
$script:PwshPath = (Get-Process -Id $PID).Path
if ([string]::IsNullOrWhiteSpace($script:PwshPath)) { $script:PwshPath = Join-Path $PSHOME 'powershell.exe' }
$script:XamlPath = Join-Path $script:AppRoot 'MainWindow.xaml'

foreach ($d in @($script:DataDir, $script:LogDir, $script:QuarantineDir)) {
    [void][IO.Directory]::CreateDirectory($d)
}

$script:Window = [Windows.Markup.XamlReader]::Parse((Get-Content -LiteralPath $script:XamlPath -Raw -Encoding UTF8))

function Ui([string]$Name) {
    return $script:Window.FindName($Name)
}

$script:RootCheckBoxes = [System.Collections.Generic.List[object]]::new()
$script:Table = $null
$script:View = $null
$script:CurrentScanDir = $null
$script:ScannerProcess = $null
$script:ScanTimer = $null
$script:SelectionTimer = $null
$script:FilterTimer = $null
$script:DialogResult = $false
$script:SuppressSelectionSummary = $false
$script:StrictHashCheck = $false
$script:StatusCheckTimer = $null
$script:StatusCheckRows = $null
$script:StatusCheckIndex = 0
$script:StatusCheckNewMissing = 0
$script:StatusCheckShowMessage = $false
$script:SummaryTimer = $null
$script:CleanupProgressWindow = $null
$script:CleanupProgressBar = $null
$script:CleanupProgressText = $null
$script:CleanupCancelButton = $null
$script:CleanupCancel = $false
$script:CleanupDone = $false
$script:UiPumpFrame = $null
$script:DeleteShortcuts = $true
$script:SecureErase = $false
$script:ManifestPermanentButton = $null
$script:PermanentDialog = $null
$script:PermanentDialogOK = $null
$script:PermanentDialogText = $null
$script:PendingDuplicateTab = $false

function Format-Bytes {
    param([long]$Bytes)
    if ($Bytes -lt 0) { return '0 B' }
    if ($Bytes -ge 1TB) { return ('{0:N2} TiB' -f ($Bytes / 1TB)) }
    if ($Bytes -ge 1GB) { return ('{0:N2} GiB' -f ($Bytes / 1GB)) }
    if ($Bytes -ge 1MB) { return ('{0:N2} MiB' -f ($Bytes / 1MB)) }
    if ($Bytes -ge 1KB) { return ('{0:N1} KiB' -f ($Bytes / 1KB)) }
    return "$Bytes B"
}

function Show-Info {
    param([string]$Text, [string]$Title = '本地文件清理助手')
    if ($script:Window.IsVisible) {
        [void][Windows.MessageBox]::Show($script:Window, $Text, $Title,
            [Windows.MessageBoxButton]::OK, [Windows.MessageBoxImage]::Information)
    }
    else {
        [void][Windows.MessageBox]::Show($Text, $Title,
            [Windows.MessageBoxButton]::OK, [Windows.MessageBoxImage]::Information)
    }
}

function Show-Warn {
    param([string]$Text, [string]$Title = '请注意')
    if ($script:Window.IsVisible) {
        [void][Windows.MessageBox]::Show($script:Window, $Text, $Title,
            [Windows.MessageBoxButton]::OK, [Windows.MessageBoxImage]::Warning)
    }
    else {
        [void][Windows.MessageBox]::Show($Text, $Title,
            [Windows.MessageBoxButton]::OK, [Windows.MessageBoxImage]::Warning)
    }
}

function Show-Error {
    param([string]$Text, [string]$Title = '操作未完成')
    if ($script:Window.IsVisible) {
        [void][Windows.MessageBox]::Show($script:Window, $Text, $Title,
            [Windows.MessageBoxButton]::OK, [Windows.MessageBoxImage]::Error)
    }
    else {
        [void][Windows.MessageBox]::Show($Text, $Title,
            [Windows.MessageBoxButton]::OK, [Windows.MessageBoxImage]::Error)
    }
}

function Set-Page {
    param([string]$Name)
(Ui 'RangePage').Visibility = if ($Name -eq 'Range') { 'Visible' } else { 'Collapsed' }
(Ui 'ScanPage').Visibility = if ($Name -eq 'Scan') { 'Visible' } else { 'Collapsed' }
(Ui 'ResultsPage').Visibility = if ($Name -eq 'Results') { 'Visible' } else { 'Collapsed' }
}

function Add-RootCheckBoxes {
$panel = (Ui 'RootsPanel')
    $panel.Children.Clear()
    $script:RootCheckBoxes.Clear()

    $names = @('Downloads', 'Desktop', 'Documents', 'Pictures', 'Videos', 'Music')
    foreach ($name in $names) {
        $path = Join-Path $env:USERPROFILE $name
        $cb = [Windows.Controls.CheckBox]::new()
        $cb.Margin = [Windows.Thickness]::new(0, 5, 0, 5)
        $cb.Tag = $path
        $cb.Content = "$name  —  $path"
        $cb.ToolTip = $path
        $cb.IsChecked = ($name -ne 'Music')
        $cb.IsEnabled = [IO.Directory]::Exists($path)
        if (-not $cb.IsEnabled) {
            $cb.Content = "$name  —  $path（不存在）"
        }
        [void]$panel.Children.Add($cb)
        $script:RootCheckBoxes.Add($cb)
    }
}

function Set-ExclusionText {
    $workspace = Join-Path $env:USERPROFILE 'Documents\deepseek-harness'
    $lines = @(
        '系统与程序：',
        'C:\Windows；C:\Program Files；C:\Program Files (x86)；C:\ProgramData；',
        'C:\$Recycle.Bin；C:\System Volume Information；C:\Recovery；C:\PerfLogs；C:\Users\Public。',
        '',
        '用户目录中的非扫描项：',
        'AppData；Application Data；Cookies；Local Settings；Recent；SendTo；Templates；',
        '「开始」菜单；Favorites；Links；NetHood；PrintHood；Searches；Contacts；Saved Games；',
        'Library；ansel；所有点号开头的配置目录。',
        '',
        '代码/版本库：',
        '.git；.svn；.hg；node_modules；vendor；Pods；DerivedData；.venv；venv；__pycache__；',
        'target；build；dist；out；.next；.nuxt；.gradle；.idea；.vs；.vscode；coverage。',
        '检测到 package.json、pyproject.toml、Cargo.toml、go.mod、*.sln、*.uproject、*.unity、project.godot 等项目标记时，整目录跳过。',
        'Unreal Projects、UnrealEngine、Unity、Unity Projects、Intermediate、Saved、Binaries、DerivedDataCache 默认排除。',
        '',
        '备份/虚拟机/游戏：',
        '名称含 Backup、Backups、备份；Virtual Machines；VirtualBox VMs；VMware；Hyper-V；',
        '*.vmdk；*.vdi；*.vhdx；*.vhd；*.qcow2；*.ova；*.ovf；Saved Games；SaveGames；Saves；My Games。',
        '',
        '浏览器/邮件/聊天：',
        'Chrome；Edge；Firefox；Brave；Opera；Vivaldi 配置目录；Bookmarks；History；',
        'Login Data；Web Data；Cookies；*.pst；*.ost；*.mbox；*.eml；*.msg；',
        'WeChat Files；Tencent Files；WhatsApp；Telegram；Discord；Slack；Teams；Outlook；Mail。',
        '',
        '密码/密钥/证书/财务：',
        '.ssh；.gnupg；1Password；KeePass；Bitwarden；*.kdbx；*.key；*.pem；*.pfx；*.p12；',
        '*.cer；*.crt；*.gpg；*.asc；id_rsa*；id_ed25519*；*.keystore；*.jks；.env；.env.*；',
        '名称含 财务、发票、税务、银行、账单、报销、会计、工资、存款、投资、保险、贷款、预算 的目录或文件。',
        '',
        '云同步：',
        'OneDrive；iCloudDrive；Dropbox；Google Drive；Google 云端硬盘；BaiduNetdisk；',
        '百度网盘；阿里云盘；坚果云；Nutstore；WPS Cloud；WPS Cloud Files；Seafile；Nextcloud；MEGA。',
        '',
        '工具自身：',
        $workspace,
        '',
        '链接/占位符：',
        '符号链接、junction、重解析点、云占位文件、离线文件不会被进入或触发下载。'
    )
(Ui 'ExclusionsText').Text = ($lines -join [Environment]::NewLine)
}

function Reset-ScanLabels {
(Ui 'ScanFilesText').Text = '0'
(Ui 'ScanDirsText').Text = '0'
(Ui 'ScanHashedText').Text = '0'
(Ui 'ScanErrorsText').Text = '0'
(Ui 'ScanElapsedText').Text = '0s'
(Ui 'ScanPhaseText').Text = '正在初始化扫描'
(Ui 'ScanCurrentPathText').Text = ''
(Ui 'CancelScanButton').IsEnabled = $true
(Ui 'CancelScanButton').Content = '取消扫描（不会修改文件）'
}

function Start-FilterDebounce {
    if ($null -eq $script:FilterTimer) {
        $script:FilterTimer = [Windows.Threading.DispatcherTimer]::new()
        $script:FilterTimer.Interval = [TimeSpan]::FromMilliseconds(260)
        $script:FilterTimer.Add_Tick({
            $script:FilterTimer.Stop()
            Apply-Filter
        })
    }
    $script:FilterTimer.Stop()
    $script:FilterTimer.Start()
}

function Commit-GridEdits {
    try { [void](Ui 'FilesGrid').CommitEdit([Windows.Controls.DataGridEditingUnit]::Cell, $true) } catch { }
    try { [void](Ui 'FilesGrid').CommitEdit([Windows.Controls.DataGridEditingUnit]::Row, $true) } catch { }
}

function Apply-Filter {
    if ($null -eq $script:View) { return }
    Commit-GridEdits
    try {
        $tab = $null
if ($null -ne (Ui 'CategoryTabs').SelectedItem) { $tab = [string]((Ui 'CategoryTabs').SelectedItem.Tag) }
        if ([string]::IsNullOrWhiteSpace($tab)) { $tab = 'CANDS' }

        $filters = [System.Collections.Generic.List[string]]::new()
        switch ($tab) {
            'CANDS'    { $filters.Add('IsError = false') }
            'DUP'      { $filters.Add('IsDuplicate = true') }
            'LARGE'    { $filters.Add('IsLarge = true') }
            'TEMP'     { $filters.Add('IsTemp = true') }
            'SAMENAME' { $filters.Add('IsSameNameDiff = true') }
            'VERSION'  { $filters.Add('IsVersion = true') }
            'ERROR'    { $filters.Add('IsError = true') }
            default    { $filters.Add('IsError = false') }
        }

        $q = ([string]((Ui 'SearchBox').Text)).Trim()
        if ($q.Length -gt 0) {
            $qe = $q.Replace("'", "''")
            $filters.Add("(Name LIKE '%$qe%' OR Path LIKE '%$qe%')")
        }

        if ($null -ne (Ui 'DirectoryFilter').SelectedItem -and [string]((Ui 'DirectoryFilter').SelectedItem) -ne '全部目录') {
            $d = ([string]((Ui 'DirectoryFilter').SelectedItem)).Replace("'", "''")
            $filters.Add("Path LIKE '$d%'")
        }
        if ($null -ne (Ui 'TypeFilter').SelectedItem -and [string]((Ui 'TypeFilter').SelectedItem) -ne '全部类型') {
            $t = ([string]((Ui 'TypeFilter').SelectedItem)).Replace("'", "''")
            $filters.Add("TypeText = '$t'")
        }
        if ($null -ne (Ui 'RiskFilter').SelectedItem -and [string]((Ui 'RiskFilter').SelectedItem) -ne '全部风险') {
            $r = ([string]((Ui 'RiskFilter').SelectedItem)).Replace("'", "''")
            $filters.Add("Risk = '$r'")
        }

        $sizeFilter = [string]((Ui 'SizeFilter').SelectedItem)
        switch ($sizeFilter) {
            '≥500 MiB'    { $filters.Add('Size >= 524288000') }
            '100–500 MiB' { $filters.Add('Size >= 104857600 AND Size < 524288000') }
            '10–100 MiB'  { $filters.Add('Size >= 10485760 AND Size < 104857600') }
            '<10 MiB'     { $filters.Add('Size < 10485760') }
        }

        $timeFilter = [string]((Ui 'TimeFilter').SelectedItem)
        switch ($timeFilter) {
            '修改≥1年' {
                $cut = [DateTime]::Today.AddYears(-1).ToString('yyyy-MM-dd')
                $filters.Add("Modified < #$cut#")
            }
            '修改≥180天' {
                $cut = [DateTime]::Today.AddDays(-180).ToString('yyyy-MM-dd')
                $filters.Add("Modified < #$cut#")
            }
            '修改≥90天' {
                $cut = [DateTime]::Today.AddDays(-90).ToString('yyyy-MM-dd')
                $filters.Add("Modified < #$cut#")
            }
            '最近30天修改' {
                $cut = [DateTime]::Today.AddDays(-30).ToString('yyyy-MM-dd')
                $filters.Add("Modified >= #$cut#")
            }
        }

        if ([bool]((Ui 'SelectedOnlyCheckBox').IsChecked)) {
            $filters.Add('IsSelected = true')
        }

        $script:View.RowFilter = ($filters -join ' AND ')

$sort = [string]((Ui 'SortFilter').SelectedItem)
        switch ($sort) {
            '大小从大到小'     { $script:View.Sort = 'Size DESC' }
            '修改时间从旧到新' { $script:View.Sort = 'Modified ASC' }
            '访问时间从旧到新' { $script:View.Sort = 'Accessed ASC' }
            '名称'             { $script:View.Sort = 'Name ASC' }
            default {
                if ($tab -eq 'DUP') { $script:View.Sort = 'Exists DESC, GroupId ASC, Size DESC' }
                else { $script:View.Sort = 'Exists DESC, RiskRank DESC, Size DESC' }
            }
        }

(Ui 'VisibleCountText').Text = "当前显示 $($script:View.Count) 个"
        Update-SelectionSummary
    }
    catch {
(Ui 'VisibleCountText').Text = '筛选表达式无法应用'
    }
}

function Update-SelectionSummaryNow {
    if ($null -eq $script:Table) { return }
    if ($script:SuppressSelectionSummary) { return }
    $count = 0
    $dirs = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $bytes = [long]0
    $seenPhysical = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)

    foreach ($r in $script:Table.Rows) {
        if (-not [bool]$r['IsSelected']) { continue }
        $count++
        try { [void]$dirs.Add([IO.Path]::GetDirectoryName([string]$r['Path'])) } catch { }
        $identity = [string]$r['Identity']
        if (-not [string]::IsNullOrWhiteSpace($identity)) {
            if (-not $seenPhysical.Add($identity)) { continue }
        }
        $bytes += [long]$r['Size']
    }

(Ui 'SelectionSummaryText').Text = "已勾选 $count 个文件 | $($dirs.Count) 个目录 | 预计移出 $(Format-Bytes $bytes)"
(Ui 'GenerateManifestButton').IsEnabled = ($count -gt 0)
}

function Update-SelectionSummary {
    if ($null -eq $script:Table) { return }
    if ($script:SuppressSelectionSummary) { return }
    if ($null -eq $script:SummaryTimer) {
        $script:SummaryTimer = [Windows.Threading.DispatcherTimer]::new()
        $script:SummaryTimer.Interval = [TimeSpan]::FromMilliseconds(120)
        $script:SummaryTimer.Add_Tick({
            $script:SummaryTimer.Stop()
            Update-SelectionSummaryNow
        })
    }
    $script:SummaryTimer.Stop()
    $script:SummaryTimer.Start()
}

function Update-RowDetails {
    $item = Ui('FilesGrid').SelectedItem
    if ($null -eq $item -or -not ($item -is [Data.DataRowView])) {
        Ui('RowDetailsText').Text = '选中一行查看完整详情。'
        return
    }
    $r = $item.Row
    $lines = [System.Collections.Generic.List[string]]::new()
    $lines.Add("文件名：$($r['Name'])")
    $lines.Add("完整路径：$($r['Path'])")
    $lines.Add("类型：$($r['TypeText'])")
    $lines.Add("用途：$($r['Purpose'])")
    $lines.Add("大小：$($r['SizeText'])    状态：$($r['StatusText'])")
    $lines.Add("修改时间：$($r['ModifiedText'])    最近访问：$($r['AccessedText'])")
    $lines.Add("分类：$($r['Categories'])")
    $lines.Add("推荐理由：$($r['Reason'])")
    $lines.Add("风险：$($r['Risk'])")
    if (-not [string]::IsNullOrWhiteSpace([string]$r['GroupId'])) {
        $lines.Add("重复组：$($r['GroupId'])    同组路径：$($r['GroupPaths'])    物理副本：$($r['GroupDistinct'])")
    }
    if (-not [string]::IsNullOrWhiteSpace([string]$r['Hash'])) {
        $lines.Add("SHA-256：$($r['Hash'])")
    }
    if ([bool]$r['IsHardLink']) {
        $lines.Add("注意：该路径属于同一物理文件的多个硬链接之一；只删除部分路径不会释放空间。")
    }
    if (-not [string]::IsNullOrWhiteSpace([string]$r['Error'])) {
        $lines.Add("错误/跳过原因：$($r['Error'])")
    }
    Ui('RowDetailsText').Text = ($lines -join [Environment]::NewLine)
}

function Update-FileStatusSync {
    param([switch]$ShowMessage)
    if ($null -eq $script:Table) { return }
    $missingNow = 0
    $alreadyMissing = 0
    $checked = 0
    foreach ($r in $script:Table.Rows) {
        if ([bool]$r['IsError']) { continue }
        if (-not [bool]$r['Exists']) { $alreadyMissing++; continue }
        $checked++
        $path = [string]$r['Path']
        if ([IO.File]::Exists($path)) { continue }

        $r['Exists'] = $false
        $r['StatusText'] = '已不在原位置'
        $r['CanSelect'] = $false
        $r['IsSelected'] = $false
        $r['Risk'] = '不建议'
        $r['RiskRank'] = 4
        $reason = [string]$r['Reason']
        if ($reason -notlike '*已不在原位置*') {
            $r['Reason'] = $reason + '；文件已不在原位置（可能已被删除或移动到回收站）'
        }
        $missingNow++
    }

    if ($missingNow -gt 0) {
        Update-SelectionSummary
        if ([bool]((Ui 'SelectedOnlyCheckBox').IsChecked)) { Apply-Filter }
        Ui('FooterText').Text = "检测到 $($alreadyMissing + $missingNow) 个文件已不在原位置；已自动取消这些行的勾选。"
    }
    elseif ($ShowMessage) {
        Ui('FooterText').Text = "状态检查完成：检查 $checked 个文件，未发现缺失。"
    }
    if ($ShowMessage) {
        Show-Info "状态检查完成。`n检查文件：$checked`n新发现缺失：$missingNow`n累计缺失：$($alreadyMissing + $missingNow)"
    }
}

function Ensure-StatusCheckTimer {
    if ($null -ne $script:StatusCheckTimer) { return }
    $script:StatusCheckTimer = [Windows.Threading.DispatcherTimer]::new()
    $script:StatusCheckTimer.Interval = [TimeSpan]::FromMilliseconds(35)
    $script:StatusCheckTimer.Add_Tick({ Step-FileStatusCheck })
}

function Start-FileStatusCheck {
    param([switch]$ShowMessage)
    if ($null -eq $script:Table) { return }
    Ensure-StatusCheckTimer
    $script:StatusCheckTimer.Stop()
    $script:StatusCheckRows = @($script:Table.Rows | Where-Object { -not [bool]$_['IsError'] })
    $script:StatusCheckIndex = 0
    $script:StatusCheckNewMissing = 0
    $script:StatusCheckShowMessage = [bool]$ShowMessage
    $script:StatusCheckTimer.Start()
    Ui('FooterText').Text = "正在检查文件状态：0 / $($script:StatusCheckRows.Count)…（界面可继续操作）"
}

function Step-FileStatusCheck {
    if ($null -eq $script:StatusCheckRows) { $script:StatusCheckTimer.Stop(); return }
    $total = $script:StatusCheckRows.Count
    if ($script:StatusCheckIndex -ge $total) {
        Complete-FileStatusCheck
        return
    }

    $batch = 120
    $end = [Math]::Min($script:StatusCheckIndex + $batch, $total)
    $script:SuppressSelectionSummary = $true
    try {
        for ($i = $script:StatusCheckIndex; $i -lt $end; $i++) {
            $r = $script:StatusCheckRows[$i]
            if ($null -eq $r) { continue }
            if (-not [bool]$r['Exists']) { continue }
            $path = [string]$r['Path']
            if ([IO.File]::Exists($path)) { continue }

            $r['Exists'] = $false
            $r['StatusText'] = '已不在原位置'
            $r['CanSelect'] = $false
            $r['IsSelected'] = $false
            $r['Risk'] = '不建议'
            $r['RiskRank'] = 4
            $reason = [string]$r['Reason']
            if ($reason -notlike '*已不在原位置*') {
                $r['Reason'] = $reason + '；文件已不在原位置（可能已被删除或移动到回收站）'
            }
            $script:StatusCheckNewMissing++
        }
    }
    finally {
        $script:SuppressSelectionSummary = $false
    }

    $script:StatusCheckIndex = $end
    if ($script:StatusCheckIndex -lt $total) {
        Ui('FooterText').Text = "正在检查文件状态：$($script:StatusCheckIndex) / $total…（界面可继续操作）"
        return
    }
    Complete-FileStatusCheck
}

function Complete-FileStatusCheck {
    $script:StatusCheckTimer.Stop()
    $total = if ($null -eq $script:StatusCheckRows) { 0 } else { $script:StatusCheckRows.Count }
    $newMissing = [int]$script:StatusCheckNewMissing
    $allMissing = @($script:Table.Rows | Where-Object { -not [bool]$_['Exists'] -and -not [bool]$_['IsError'] }).Count
    Update-SelectionSummary
    if ([bool]((Ui 'SelectedOnlyCheckBox').IsChecked)) { Apply-Filter }
    if ($allMissing -gt 0) {
        Ui('FooterText').Text = "状态检查完成：检查 $total 个文件，本次新发现缺失 $newMissing 个；累计缺失 $allMissing 个。"
    }
    else {
        Ui('FooterText').Text = "状态检查完成：检查 $total 个文件，未发现缺失。"
    }
    if ($script:StatusCheckShowMessage) {
        Show-Info "状态检查完成。`n检查文件：$total`n本次新发现缺失：$newMissing`n累计缺失：$allMissing"
    }
    $script:StatusCheckShowMessage = $false
}

function Update-FileStatus {
    param([switch]$ShowMessage)
    Start-FileStatusCheck -ShowMessage:$ShowMessage
}

function Get-VisualChildren {
    param(
        [Windows.DependencyObject]$Parent,
        [Type]$Type
    )
    $result = [System.Collections.Generic.List[object]]::new()
    $count = 0
    try { $count = [Windows.Media.VisualTreeHelper]::GetChildrenCount($Parent) } catch { return $result }
    for ($i = 0; $i -lt $count; $i++) {
        $child = $null
        try { $child = [Windows.Media.VisualTreeHelper]::GetChild($Parent, $i) } catch { continue }
        if ($null -eq $child) { continue }
        if ($Type.IsInstanceOfType($child)) { $result.Add($child) }
        foreach ($grand in (Get-VisualChildren -Parent $child -Type $Type)) { $result.Add($grand) }
    }
    return $result
}

function Find-ButtonFromEvent {
    param($Source)
    $cur = $Source
    while ($null -ne $cur -and -not ($cur -is [Windows.Controls.Button])) {
        try { $cur = [Windows.Media.VisualTreeHelper]::GetParent($cur) } catch { return $null }
    }
    return $cur
}

function Open-Path {
    param([string]$Path)
    if (-not [IO.File]::Exists($Path)) {
        Show-Warn "该文件已不在原位置，无法打开：`n$Path`n`n它可能已被删除或移动到回收站。请点击「刷新状态」更新清单。"
        return
    }
    $ext = [IO.Path]::GetExtension($Path).ToLowerInvariant()
    if ($ext -in @('.exe', '.msi', '.bat', '.cmd', '.ps1', '.com', '.scr')) {
        $ans = [Windows.MessageBox]::Show($script:Window,
            "该文件是可执行文件或安装程序：`n$Path`n`n「打开」会运行它，而不只是查看内容。继续？",
            '可执行文件确认', [Windows.MessageBoxButton]::YesNo, [Windows.MessageBoxImage]::Warning)
        if ($ans -ne [Windows.MessageBoxResult]::Yes) { return }
    }
    $errors = [System.Collections.Generic.List[string]]::new()
    try {
        Start-Process -FilePath $Path -ErrorAction Stop | Out-Null
        return
    }
    catch {
        $errors.Add('Start-Process：' + $_.Exception.Message)
    }

    try {
        $shell = New-Object -ComObject Shell.Application
        $item = $shell.Namespace(0).ParseName($Path)
        if ($null -ne $item) {
            $item.InvokeVerb('open')
            return
        }
        $errors.Add('Shell.Application：无法解析文件')
    }
    catch {
        $errors.Add('Shell.Application：' + $_.Exception.Message)
    }

    try {
        Start-Process -FilePath 'explorer.exe' -ArgumentList "`"$Path`"" -ErrorAction Stop | Out-Null
        return
    }
    catch {
        $errors.Add('explorer.exe：' + $_.Exception.Message)
    }

    Show-Error ("无法用系统默认程序打开：`n$Path`n`n" + ($errors -join "`n") + "`n`n可以改用「预览选中」或「查看文件夹」。")
}

function Show-FolderContents {
    param([string]$Path)

    $folder = if ([IO.Directory]::Exists($Path)) { $Path } else { [IO.Path]::GetDirectoryName($Path) }
    if ([string]::IsNullOrWhiteSpace($folder) -or -not [IO.Directory]::Exists($folder)) {
        Show-Warn "文件夹不存在：`n$Path"
        return
    }
    try { [Windows.Clipboard]::SetText($folder) } catch { }

    $w = [Windows.Window]::new()
    $w.Title = '文件夹 - ' + [IO.Path]::GetFileName($folder.TrimEnd('\'))
    $w.Width = 900
    $w.Height = 620
    $w.WindowStartupLocation = 'CenterOwner'
    if ($script:Window.IsVisible) { $w.Owner = $script:Window }
    $w.Background = [Windows.Media.Brushes]::White
    $script:FolderDialog = $w

    $grid = [Windows.Controls.Grid]::new()
    $grid.Margin = [Windows.Thickness]::new(16)
    for ($i = 0; $i -lt 3; $i++) { $grid.RowDefinitions.Add([Windows.Controls.RowDefinition]::new()) }

    $head = [Windows.Controls.StackPanel]::new()
    $title = [Windows.Controls.TextBlock]::new()
    $title.Text = $folder
    $title.FontSize = 13
    $title.FontWeight = 'SemiBold'
    $title.Foreground = [Windows.Media.Brushes]::DimGray
    $title.TextTrimming = 'CharacterEllipsis'
    [void]$head.Children.Add($title)
    [Windows.Controls.Grid]::SetRow($head, 0)
    [void]$grid.Children.Add($head)

    $list = [Windows.Controls.ListBox]::new()
    $list.FontSize = 13
    $list.DisplayMemberPath = 'Display'
    $list.Margin = [Windows.Thickness]::new(0, 10, 0, 10)
    $list.Add_MouseDoubleClick({
        $it = $script:FolderListBox.SelectedItem
        if ($null -eq $it) { return }
        $p = [string]$it.Path
        if ([IO.Directory]::Exists($p)) {
            Show-FolderContents -Path $p
            return
        }
        if ([IO.File]::Exists($p)) {
            try { Start-Process -FilePath $p -ErrorAction Stop | Out-Null }
            catch { Show-Warn ("无法打开文件：`n" + $_.Exception.Message) }
        }
    })
    $script:FolderListBox = $list

    try {
        $dirs = @([IO.Directory]::GetDirectories($folder) | Sort-Object)
        $files = @([IO.Directory]::GetFiles($folder) | Sort-Object)
        foreach ($d in $dirs) {
            $name = [IO.Path]::GetFileName($d)
            $list.Items.Add([pscustomobject]@{ Display = "[文件夹]  $name"; Path = $d }) | Out-Null
        }
        foreach ($f in $files) {
            $fi = [IO.FileInfo]::new($f)
            $list.Items.Add([pscustomobject]@{
                Display = ("{0}    {1}    {2:yyyy-MM-dd HH:mm}" -f $fi.Name, (Format-Bytes $fi.Length), $fi.LastWriteTime)
                Path = $f
            }) | Out-Null
        }
    }
    catch {
        $list.Items.Add([pscustomobject]@{ Display = "读取失败：" + $_.Exception.Message; Path = '' }) | Out-Null
    }
    [Windows.Controls.Grid]::SetRow($list, 1)
    [void]$grid.Children.Add($list)

    $sp = [Windows.Controls.StackPanel]::new()
    $sp.Orientation = 'Horizontal'
    $sp.HorizontalAlignment = 'Right'
    [Windows.Controls.Grid]::SetRow($sp, 2)

    $explorerBtn = [Windows.Controls.Button]::new()
    $explorerBtn.Content = '用资源管理器打开'
    $explorerBtn.Padding = [Windows.Thickness]::new(14, 7, 14, 7)
    $explorerBtn.Margin = [Windows.Thickness]::new(0, 0, 8, 0)
    $explorerBtn.Add_Click({
        try {
            Start-Process -FilePath 'explorer.exe' -ArgumentList "`"$script:FolderListPath`"" -ErrorAction Stop | Out-Null
        }
        catch {
            Show-Warn ("无法启动资源管理器：`n" + $_.Exception.Message + "`n`n可以直接在列表里双击文件打开。")
        }
    })

    $copy = [Windows.Controls.Button]::new()
    $copy.Content = '复制路径'
    $copy.Padding = [Windows.Thickness]::new(14, 7, 14, 7)
    $copy.Margin = [Windows.Thickness]::new(0, 0, 8, 0)
    $copy.Add_Click({ try { [Windows.Clipboard]::SetText($script:FolderListPath) } catch { } })

    $open = [Windows.Controls.Button]::new()
    $open.Content = '打开选中'
    $open.Padding = [Windows.Thickness]::new(14, 7, 14, 7)
    $open.Margin = [Windows.Thickness]::new(0, 0, 8, 0)
    $open.Add_Click({
        $it = $script:FolderListBox.SelectedItem
        if ($null -eq $it) { return }
        $p = [string]$it.Path
        if ([IO.Directory]::Exists($p)) { Show-FolderContents -Path $p; return }
        if ([IO.File]::Exists($p)) {
            try { Start-Process -FilePath $p -ErrorAction Stop | Out-Null }
            catch { Show-Warn ("无法打开文件：`n" + $_.Exception.Message) }
        }
    })

    $close = [Windows.Controls.Button]::new()
    $close.Content = '关闭'
    $close.Padding = [Windows.Thickness]::new(14, 7, 14, 7)
    $close.Add_Click({ $script:FolderDialog.Close() })

    [void]$sp.Children.Add($explorerBtn)
    [void]$sp.Children.Add($copy)
    [void]$sp.Children.Add($open)
    [void]$sp.Children.Add($close)
    [void]$grid.Children.Add($sp)

    $script:FolderListPath = $folder
    $w.Content = $grid
    [void]$w.ShowDialog()
}

function Open-Folder {
    param([string]$Path)
    Show-FolderContents -Path $Path
}

function Show-TextPreview {
    param(
        [string]$Title,
        [string]$Path,
        [string]$Text
    )

    $w = [Windows.Window]::new()
    $w.Title = "预览 - $Title"
    $w.Width = 980
    $w.Height = 760
    if ($script:Window.IsVisible) { $w.Owner = $script:Window }
    $w.WindowStartupLocation = 'CenterOwner'
    $w.Background = [Windows.Media.Brushes]::White
    $script:PreviewDialog = $w
    $script:PreviewPath = $Path

    $grid = [Windows.Controls.Grid]::new()
    $grid.Margin = [Windows.Thickness]::new(14)
    $grid.RowDefinitions.Add([Windows.Controls.RowDefinition]::new())
    $grid.RowDefinitions.Add([Windows.Controls.RowDefinition]::new())
    $grid.RowDefinitions.Add([Windows.Controls.RowDefinition]::new())

    $head = [Windows.Controls.TextBlock]::new()
    $head.Text = $Path
    $head.TextWrapping = 'Wrap'
    $head.Foreground = [Windows.Media.Brushes]::DimGray
    $head.Margin = [Windows.Thickness]::new(0, 0, 0, 8)
    [Windows.Controls.Grid]::SetRow($head, 0)
    [void]$grid.Children.Add($head)

    $tb = [Windows.Controls.TextBox]::new()
    $tb.Text = $Text
    $tb.IsReadOnly = $true
    $tb.AcceptsReturn = $true
    $tb.TextWrapping = 'NoWrap'
    $tb.VerticalScrollBarVisibility = 'Auto'
    $tb.HorizontalScrollBarVisibility = 'Auto'
    $tb.FontFamily = [Windows.Media.FontFamily]::new('Consolas')
    $tb.FontSize = 12
    [Windows.Controls.Grid]::SetRow($tb, 1)
    [void]$grid.Children.Add($tb)

    $sp = [Windows.Controls.StackPanel]::new()
    $sp.Orientation = 'Horizontal'
    $sp.HorizontalAlignment = 'Right'
    $sp.Margin = [Windows.Thickness]::new(0, 10, 0, 0)
    [Windows.Controls.Grid]::SetRow($sp, 2)

    $note = [Windows.Controls.TextBlock]::new()
    $note.Text = '只读预览，不会修改文件。大文件只显示开头部分。'
    $note.Foreground = [Windows.Media.Brushes]::DimGray
    $note.VerticalAlignment = 'Center'
    $note.Margin = [Windows.Thickness]::new(0, 0, 16, 0)
    [void]$sp.Children.Add($note)

    $folder = [Windows.Controls.Button]::new()
    $folder.Content = '查看文件夹'
    $folder.Padding = [Windows.Thickness]::new(12, 6, 12, 6)
    $folder.Margin = [Windows.Thickness]::new(0, 0, 8, 0)
    $folder.Add_Click({ Open-Folder -Path $script:PreviewPath })

    $close = [Windows.Controls.Button]::new()
    $close.Content = '关闭'
    $close.Padding = [Windows.Thickness]::new(12, 6, 12, 6)
    $close.Add_Click({ $script:PreviewDialog.Close() })

    [void]$sp.Children.Add($folder)
    [void]$sp.Children.Add($close)
    [void]$grid.Children.Add($sp)

    $w.Content = $grid
    [void]$w.ShowDialog()
}

function Preview-Selected {
    $item = Get-SelectedRow
    if ($null -eq $item) { return }
    $path = [string]$item['Path']
    if (-not [IO.File]::Exists($path)) {
        Show-Warn "该文件已不在原位置，无法预览：`n$path"
        return
    }
    $ext = [IO.Path]::GetExtension($path).ToLowerInvariant()
    $textExts = @('.txt', '.md', '.cpp', '.c', '.h', '.hpp', '.py', '.js', '.ts', '.json', '.xml',
                 '.html', '.htm', '.css', '.log', '.ini', '.cfg', '.yaml', '.yml', '.csv', '.java',
                 '.cs', '.go', '.rs', '.sql', '.bat', '.cmd', '.ps1', '.sh', '.conf', '.properties')
    if ($textExts -notcontains $ext) {
        Show-Info "该格式不支持内置文本预览：$ext`n`n请使用「预览选中」或「查看文件夹」。"
        return
    }
    try {
        $fi = [IO.FileInfo]::new($path)
        $lines = @(Get-Content -LiteralPath $path -TotalCount 5000 -ErrorAction Stop)
        $text = $lines -join [Environment]::NewLine
        if ($lines.Count -ge 5000) { $text += [Environment]::NewLine + '...（仅显示前 5000 行）' }
        Show-TextPreview -Title ([IO.Path]::GetFileName($path)) -Path $path -Text $text
    }
    catch {
        Show-Error "无法预览：`n$($_.Exception.Message)"
    }
}

function Get-SelectedRow {
$item = (Ui 'FilesGrid').SelectedItem
    if ($null -eq $item -or -not ($item -is [Data.DataRowView])) {
        Show-Info '请先在清单中选中一行。'
        return $null
    }
    return $item
}

function Clear-KeepInGroup {
    param([string]$GroupId)
    foreach ($r in $script:Table.Rows) {
        if ([string]$r['GroupId'] -eq $GroupId) {
            $r['Keep'] = $false
            $r['KeepMark'] = ''
        }
    }
}

function Keep-Selected {
    $item = Get-SelectedRow
    if ($null -eq $item) { return }
    if (-not [bool]$item['IsDuplicate']) {
        Show-Info '「保留此副本」只适用于完全重复组。'
        return
    }
    $gid = [string]$item['GroupId']
    Clear-KeepInGroup $gid
    $item['Keep'] = $true
    $item['KeepMark'] = '保留'
    $item['IsSelected'] = $false
    $item['AllCopiesIntent'] = $false
    Update-SelectionSummary
}

function Select-RestInGroup {
    $item = Get-SelectedRow
    if ($null -eq $item) { return }
    if (-not [bool]$item['IsDuplicate']) {
        Show-Info '「选择同组其余副本」只适用于完全重复组。'
        return
    }
    $gid = [string]$item['GroupId']
    $selectedId = [string]$item['Id']
    $n = 0
    foreach ($r in $script:Table.Rows) {
        if ([string]$r['GroupId'] -ne $gid) { continue }
        if ([string]$r['Id'] -eq $selectedId) { continue }
        if ([bool]$r['Keep']) { continue }
        if (-not [bool]$r['CanSelect']) { continue }
        $r['IsSelected'] = $true
        $n++
    }
    Update-SelectionSummary
    if ($n -eq 0) { Show-Info '该组没有其他可勾选的副本。' }
}

function Preselect-Duplicates {
    param([switch]$Silent)
    if ($null -eq $script:Table) { return }
    Commit-GridEdits

    $groups = @($script:Table.Rows | Where-Object { [bool]$_['IsDuplicate'] } | Group-Object -Property GroupId)
    $selected = 0
    $kept = 0
    $script:SuppressSelectionSummary = $true
    try {
        foreach ($g in $groups) {
            $allGroupRows = @($g.Group)
            $rows = @($allGroupRows | Where-Object { [bool]$_['Exists'] -and [bool]$_['CanSelect'] })
            if ($rows.Count -lt 2) { continue }
            if ([int]$rows[0]['GroupDistinct'] -le 1) { continue }

            $keep = $rows |
                Sort-Object -Property @{ Expression = { [DateTime]$_['Modified'] } }, @{ Expression = { ([string]$_['Path']).Length } } |
                Select-Object -First 1

            foreach ($r in $allGroupRows) {
                $r['Keep'] = $false
                $r['KeepMark'] = ''
                $r['IsSelected'] = $false
                $r['AllCopiesIntent'] = $false
            }
            $keep['Keep'] = $true
            $keep['KeepMark'] = '保留'
            $kept++
            foreach ($r in $rows) {
                if ($r -eq $keep) { continue }
                if (-not [bool]$r['CanSelect']) { continue }
                $r['IsSelected'] = $true
                $selected++
            }
        }
    }
    finally {
        $script:SuppressSelectionSummary = $false
    }

    Update-SelectionSummary
    if (-not $Silent) {
        Show-Info "已预选重复副本：$selected 个待处理，$kept 个重复组各保留 1 份。`n`n不会执行删除；请核对后生成最终操作清单。"
    }
}

function Test-AllCopiesIntentComplete {
    param([string]$GroupId)
    foreach ($r in $script:Table.Rows) {
        if ([string]$r['GroupId'] -eq $GroupId) {
            if (-not [bool]$r['IsSelected']) { return $false }
        }
    }
    return $true
}

function Show-GroupConfirm {
    param([string]$GroupId)

    $rows = @($script:Table.Rows | Where-Object { [string]$_['GroupId'] -eq $GroupId })
    if ($rows.Count -lt 2) {
        Show-Info '该组没有足够的副本。'
        return $false
    }

    $text = "该重复组共有 $($rows.Count) 个路径。`r`n`r`n"
    $text += "你选择了「处理同组全部副本」，这会处理所有路径，不保留任何副本。`r`n"
    $text += "请逐项核对：`r`n`r`n"
    foreach ($r in $rows) {
        $text += "• $($r['Path'])`r`n"
    }
    $text += "`r`n如果其中任何一份需要保留，请取消，改用「保留此副本」。"

    $w = [Windows.Window]::new()
    $w.Title = '确认处理整组副本'
    $w.Width = 940
    $w.Height = 620
    if ($script:Window.IsVisible) { $w.Owner = $script:Window }
    $w.WindowStartupLocation = 'CenterOwner'
    $w.Background = [Windows.Media.Brushes]::White
    $script:GroupDialog = $w

    $grid = [Windows.Controls.Grid]::new()
    $grid.Margin = [Windows.Thickness]::new(16)
    $grid.RowDefinitions.Add([Windows.Controls.RowDefinition]::new())
    $grid.RowDefinitions.Add([Windows.Controls.RowDefinition]::new())
    $grid.RowDefinitions.Add([Windows.Controls.RowDefinition]::new())

    $tb = [Windows.Controls.TextBox]::new()
    $tb.Text = $text
    $tb.IsReadOnly = $true
    $tb.AcceptsReturn = $true
    $tb.TextWrapping = 'Wrap'
    $tb.VerticalScrollBarVisibility = 'Auto'
    $tb.FontFamily = [Windows.Media.FontFamily]::new('Consolas')
    $tb.FontSize = 12
    [Windows.Controls.Grid]::SetRow($tb, 0)
    [void]$grid.Children.Add($tb)

    $cb = [Windows.Controls.CheckBox]::new()
    $cb.Content = '我确认以上全部路径都需要处理，且不保留该组任何副本'
    $cb.Margin = [Windows.Thickness]::new(0, 12, 0, 12)
    [Windows.Controls.Grid]::SetRow($cb, 1)
    [void]$grid.Children.Add($cb)

    $sp = [Windows.Controls.StackPanel]::new()
    $sp.Orientation = 'Horizontal'
    $sp.HorizontalAlignment = 'Right'
    [Windows.Controls.Grid]::SetRow($sp, 2)

    $cancel = [Windows.Controls.Button]::new()
    $cancel.Content = '取消'
    $cancel.Padding = [Windows.Thickness]::new(14, 7, 14, 7)
    $cancel.Margin = [Windows.Thickness]::new(0, 0, 8, 0)
    $cancel.Add_Click({ $script:GroupDialog.DialogResult = $false })

    $ok = [Windows.Controls.Button]::new()
    $ok.Content = '确认加入待处理'
    $ok.Padding = [Windows.Thickness]::new(14, 7, 14, 7)
    $ok.IsEnabled = $false
    $ok.Add_Click({ $script:GroupDialog.DialogResult = $true })
    $script:GroupDialogOK = $ok
    $cb.Add_Checked({ $script:GroupDialogOK.IsEnabled = $true })
    $cb.Add_Unchecked({ $script:GroupDialogOK.IsEnabled = $false })

    [void]$sp.Children.Add($cancel)
    [void]$sp.Children.Add($ok)
    [void]$grid.Children.Add($sp)

    $w.Content = $grid
    $result = $w.ShowDialog()
    return [bool]$result
}

function Confirm-BulkSelection {
    param(
        [object[]]$Rows,
        [string]$Title,
        [string]$Action
    )
    $rows = @($Rows | Where-Object { $null -ne $_ -and [bool]$_['CanSelect'] -and [bool]$_['Exists'] })
    if ($rows.Count -eq 0) {
        Show-Info '没有可勾选的文件。'
        return $false
    }
    if ($script:AutoConfirmBulk) {
        Write-Output "AUTO_CONFIRM=$($rows.Count)"
        return $true
    }

    $total = [long]0
    $high = 0
    $same = 0
    $dup = 0
    $temp = 0
    $folders = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($r in $rows) {
        $total += [long]$r['Size']
        if ([string]$r['Risk'] -in @('高', '不建议')) { $high++ }
        if ([bool]$r['IsSameNameDiff']) { $same++ }
        if ([bool]$r['IsDuplicate']) { $dup++ }
        if ([bool]$r['IsTemp']) { $temp++ }
        $parent = [IO.Path]::GetDirectoryName([string]$r['Path'])
        if (-not [string]::IsNullOrWhiteSpace($parent)) { [void]$folders.Add($parent) }
    }

    $msg = "$Action`n`n将勾选 $($rows.Count) 个文件，合计 $(Format-Bytes $total)，涉及 $($folders.Count) 个目录。`n`n" +
           "完全重复：$dup`n同名不同内容：$same`n高风险/不建议：$high`n临时文件：$temp`n`n" +
           '这只是勾选，不会删除；生成最终清单后仍需再次确认。'
    if ($high -gt 0 -or $same -gt 0) {
        $msg += "`n`n注意：同名不同内容不是重复文件，可能不是你要删的版本。建议先按类型/文件夹筛选，或逐个确认。"
    }

    $ans = [Windows.MessageBox]::Show($script:Window, $msg, $Title,
        [Windows.MessageBoxButton]::YesNo, [Windows.MessageBoxImage]::Warning)
    return ($ans -eq [Windows.MessageBoxResult]::Yes)
}

function Set-BulkSelection {
    param(
        [object[]]$Rows,
        [bool]$Selected
    )
    $count = 0
    $script:SuppressSelectionSummary = $true
    try {
        foreach ($r in $Rows) {
            if ($null -eq $r) { continue }
            if ($Selected) {
                if (-not [bool]$r['CanSelect'] -or -not [bool]$r['Exists']) { continue }
                $r['Keep'] = $false
                $r['KeepMark'] = ''
                $r['AllCopiesIntent'] = $false
                $r['IsSelected'] = $true
            }
            else {
                $r['IsSelected'] = $false
                $r['AllCopiesIntent'] = $false
            }
            $count++
        }
    }
    finally {
        $script:SuppressSelectionSummary = $false
    }
    Update-SelectionSummary
    return $count
}

function Select-Visible {
    if ($null -eq $script:View) { return }
    $rows = @($script:View | Where-Object { [bool]$_['CanSelect'] -and [bool]$_['Exists'] })
    if ($rows.Count -eq 0) {
        Show-Info '当前筛选结果中没有可勾选的文件。'
        return
    }
    if (-not (Confirm-BulkSelection -Rows $rows -Title '全选当前显示' -Action '将勾选当前筛选结果中的所有可勾选文件。')) { return }
    [void](Set-BulkSelection -Rows $rows -Selected $true)
}

function Clear-Visible {
    if ($null -eq $script:View) { return }
    $rows = @($script:View | Where-Object { [bool]$_['IsSelected'] })
    if ($rows.Count -eq 0) {
        Show-Info '当前筛选结果中没有已勾选的文件。'
        return
    }
    [void](Set-BulkSelection -Rows $rows -Selected $false)
}

function Select-SameFolder {
    $item = Get-SelectedRow
    if ($null -eq $item) { return }
    $parent = [IO.Path]::GetDirectoryName([string]$item['Path'])
    if ([string]::IsNullOrWhiteSpace($parent)) { return }
    $rows = @($script:Table.Rows | Where-Object {
        -not [bool]$_['IsError'] -and [bool]$_['CanSelect'] -and [bool]$_['Exists'] -and
        ([IO.Path]::GetDirectoryName([string]$_['Path']) -eq $parent)
    })
    if ($rows.Count -eq 0) {
        Show-Info '该文件夹中没有可勾选的候选。'
        return
    }
    if (-not (Confirm-BulkSelection -Rows $rows -Title '选中同文件夹' -Action "将勾选该文件夹中的全部候选：`n$parent")) { return }
    [void](Set-BulkSelection -Rows $rows -Selected $true)
}

function Select-SameFolderType {
    $item = Get-SelectedRow
    if ($null -eq $item) { return }
    $parent = [IO.Path]::GetDirectoryName([string]$item['Path'])
    $typeText = [string]$item['TypeText']
    if ([string]::IsNullOrWhiteSpace($parent) -or [string]::IsNullOrWhiteSpace($typeText)) { return }
    $rows = @($script:Table.Rows | Where-Object {
        -not [bool]$_['IsError'] -and [bool]$_['CanSelect'] -and [bool]$_['Exists'] -and
        ([IO.Path]::GetDirectoryName([string]$_['Path']) -eq $parent) -and
        ([string]$_['TypeText'] -eq $typeText)
    })
    if ($rows.Count -eq 0) {
        Show-Info '该文件夹中没有同类型的可勾选候选。'
        return
    }
    if (-not (Confirm-BulkSelection -Rows $rows -Title '同文件夹同类型' -Action "将勾选该文件夹中类型为「$typeText」的全部候选：`n$parent")) { return }
    [void](Set-BulkSelection -Rows $rows -Selected $true)
}

function Clear-SameFolder {
    $item = Get-SelectedRow
    if ($null -eq $item) { return }
    $parent = [IO.Path]::GetDirectoryName([string]$item['Path'])
    if ([string]::IsNullOrWhiteSpace($parent)) { return }
    $rows = @($script:Table.Rows | Where-Object {
        [bool]$_['IsSelected'] -and
        ([IO.Path]::GetDirectoryName([string]$_['Path']) -eq $parent)
    })
    if ($rows.Count -eq 0) {
        Show-Info '该文件夹中没有已勾选的文件。'
        return
    }
    if (-not (Confirm-BulkSelection -Rows $rows -Title '取消选中同文件夹' -Action "将从勾选结果中取消该文件夹的全部文件：`n$parent")) { return }
    [void](Set-BulkSelection -Rows $rows -Selected $false)
}

function Clear-SameFolderType {
    $item = Get-SelectedRow
    if ($null -eq $item) { return }
    $parent = [IO.Path]::GetDirectoryName([string]$item['Path'])
    $typeText = [string]$item['TypeText']
    if ([string]::IsNullOrWhiteSpace($parent) -or [string]::IsNullOrWhiteSpace($typeText)) { return }
    $rows = @($script:Table.Rows | Where-Object {
        [bool]$_['IsSelected'] -and
        ([IO.Path]::GetDirectoryName([string]$_['Path']) -eq $parent) -and
        ([string]$_['TypeText'] -eq $typeText)
    })
    if ($rows.Count -eq 0) {
        Show-Info '该文件夹中没有已勾选的同类型文件。'
        return
    }
    if (-not (Confirm-BulkSelection -Rows $rows -Title '取消同文件夹同类型' -Action "将从勾选结果中取消该文件夹中类型为「$typeText」的文件：`n$parent")) { return }
    [void](Set-BulkSelection -Rows $rows -Selected $false)
}

function Select-AllImages {
    $rows = @($script:Table.Rows | Where-Object {
        -not [bool]$_['IsError'] -and [bool]$_['CanSelect'] -and [bool]$_['Exists'] -and
        ([string]$_['TypeText'] -eq '图片')
    })
    if ($rows.Count -eq 0) {
        Show-Info '清单中没有可勾选的图片候选。'
        return
    }
    if (-not (Confirm-BulkSelection -Rows $rows -Title '选中所有图片' -Action '将勾选清单中所有图片候选（不限于当前文件夹）。')) { return }
    [void](Set-BulkSelection -Rows $rows -Selected $true)
}

function Select-AllGroup {
    $item = Get-SelectedRow
    if ($null -eq $item) { return }
    if (-not [bool]$item['IsDuplicate']) {
        Show-Info '「处理同组全部副本」只适用于完全重复组。'
        return
    }
    $gid = [string]$item['GroupId']
    if (-not (Show-GroupConfirm -GroupId $gid)) { return }

    $script:SuppressSelectionSummary = $true
    try {
        foreach ($r in $script:Table.Rows) {
            if ([string]$r['GroupId'] -eq $gid) {
                $r['Keep'] = $false
                $r['KeepMark'] = ''
                $r['IsSelected'] = $true
                $r['AllCopiesIntent'] = $true
            }
        }
    }
    finally { $script:SuppressSelectionSummary = $false }
    Update-SelectionSummary
}

function Update-StatsPanel {
    if ($null -eq $script:Table) { return }
    $panel = Ui 'StatsPanel'
    if ($null -eq $panel) { return }
    $panel.Children.Clear()

    $dup = 0; $large = 0; $temp = 0; $err = 0; $selSize = [long]0; $selCount = 0
    foreach ($r in $script:Table.Rows) {
        if ([bool]$r['IsDuplicate']) { $dup++ }
        if ([bool]$r['IsLarge']) { $large++ }
        if ([bool]$r['IsTemp']) { $temp++ }
        if ([bool]$r['IsError']) { $err++ }
        if ([bool]$r['IsSelected']) { $selCount++; $selSize += [long]$r['Size'] }
    }

    $items = [ordered]@{
        '候选'       = $script:Table.Rows.Count
        '完全重复'   = $dup
        '大文件'     = $large
        '临时文件'   = $temp
        '跳过/错误' = $err
        '已勾选'     = "$selCount 个"
        '预计释放'   = (Format-Bytes $selSize)
    }
    foreach ($k in $items.Keys) {
        $b = [Windows.Controls.Border]::new()
        $b.Background = [Windows.Media.Brushes]::WhiteSmoke
        $b.CornerRadius = [Windows.CornerRadius]::new(8)
        $b.Padding = [Windows.Thickness]::new(14, 10, 14, 10)
        $b.Margin = [Windows.Thickness]::new(0, 0, 10, 10)
        $b.MinWidth = 120
        $sp = [Windows.Controls.StackPanel]::new()
        $label = [Windows.Controls.TextBlock]::new()
        $label.Text = $k
        $label.FontSize = 11
        $label.Foreground = [Windows.Media.Brushes]::DimGray
        $value = [Windows.Controls.TextBlock]::new()
        $value.Text = [string]$items[$k]
        $value.FontSize = 16
        $value.FontWeight = 'SemiBold'
        $value.Foreground = [Windows.Media.Brushes]::Black
        $value.Margin = [Windows.Thickness]::new(0, 2, 0, 0)
        [void]$sp.Children.Add($label)
        [void]$sp.Children.Add($value)
        $b.Child = $sp
        [void]$panel.Children.Add($b)
    }
}

function Update-TabCounts {
    if ($null -eq $script:Table) { return }
    $tabs = Ui 'CategoryTabs'
    if ($null -eq $tabs) { return }
    $dup = 0; $large = 0; $temp = 0; $same = 0; $ver = 0; $err = 0
    foreach ($r in $script:Table.Rows) {
        if ([bool]$r['IsDuplicate']) { $dup++ }
        if ([bool]$r['IsLarge']) { $large++ }
        if ([bool]$r['IsTemp']) { $temp++ }
        if ([bool]$r['IsSameNameDiff']) { $same++ }
        if ([bool]$r['IsVersion']) { $ver++ }
        if ([bool]$r['IsError']) { $err++ }
    }
    foreach ($item in $tabs.Items) {
        $tag = [string]$item.Tag
        switch ($tag) {
            'CANDS'    { $item.Header = "全部候选 ($($script:Table.Rows.Count))" }
            'DUP'      { $item.Header = "完全重复 ($dup)" }
            'LARGE'    { $item.Header = "大文件 ($large)" }
            'TEMP'     { $item.Header = "临时文件 ($temp)" }
            'SAMENAME' { $item.Header = "同名不同内容 ($same)" }
            'VERSION'  { $item.Header = "历史版本 ($ver)" }
            'ERROR'    { $item.Header = "跳过/错误 ($err)" }
        }
    }
}

function Load-Result {
    param([string]$ScanDir)

    $script:CurrentScanDir = $ScanDir
    $resultPath = Join-Path $ScanDir 'result.json'
    if (-not [IO.File]::Exists($resultPath)) {
        Show-Error "找不到扫描结果：`n$resultPath"
        return
    }

    $result = Get-Content -LiteralPath $resultPath -Raw -Encoding UTF8 | ConvertFrom-Json
    if ([bool]$result.canceled -and @($result.files).Count -eq 0) {
        Show-Info '扫描已取消，没有生成完整候选清单。可以重新开始扫描。'
        Set-Page 'Range'
        return
    }

    $table = [Data.DataTable]::new('Candidates')
    $cols = @(
        @('Id', [string]), @('GroupId', [string]), @('Name', [string]), @('Path', [string]),
        @('Ext', [string]), @('TypeText', [string]), @('Purpose', [string]),
        @('Size', [long]), @('SizeText', [string]),
        @('Modified', [DateTime]), @('ModifiedText', [string]),
        @('Accessed', [DateTime]), @('AccessedText', [string]),
        @('Reason', [string]), @('Risk', [string]), @('RiskRank', [int]), @('Categories', [string]),
        @('IsDuplicate', [bool]), @('IsLarge', [bool]), @('IsTemp', [bool]),
        @('IsSameNameDiff', [bool]), @('IsVersion', [bool]), @('IsError', [bool]),
        @('IsHardLink', [bool]), @('CanSelect', [bool]), @('IsSelected', [bool]),
        @('Keep', [bool]), @('KeepMark', [string]), @('AllCopiesIntent', [bool]),
        @('Exists', [bool]), @('StatusText', [string]),
        @('GroupSize', [long]), @('GroupPaths', [int]), @('GroupDistinct', [int]),
        @('Hash', [string]), @('Identity', [string]), @('Error', [string])
    )
    foreach ($c in $cols) { [void]$table.Columns.Add([string]$c[0], [type]$c[1]) }

    $table.BeginLoadData()
    foreach ($f in @($result.files)) {
        $mod = [DateTime]::MinValue
        $acc = [DateTime]::MinValue
        try { $mod = [DateTime]::Parse([string]$f.modifiedUtc, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::RoundtripKind).ToUniversalTime() } catch { }
        try { $acc = [DateTime]::Parse([string]$f.accessedUtc, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::RoundtripKind).ToUniversalTime() } catch { }
        $modText = if ($mod.Year -ge 2000) { $mod.ToLocalTime().ToString('yyyy-MM-dd HH:mm') } else { '不可用' }
        $accText = if ($acc.Year -ge 2000) { $acc.ToLocalTime().ToString('yyyy-MM-dd HH:mm') } else { '不可用' }

        $r = $table.NewRow()
        $r['Id'] = [string]$f.id
        $r['GroupId'] = [string]$f.groupId
        $r['Name'] = [string]$f.name
        $r['Path'] = [string]$f.path
        $r['Ext'] = [string]$f.ext
        $r['TypeText'] = [string]$f.typeText
        $r['Purpose'] = [string]$f.purpose
        $r['Size'] = [long]$f.size
        $r['SizeText'] = Format-Bytes ([long]$f.size)
        $r['Modified'] = $mod
        $r['ModifiedText'] = $modText
        $r['Accessed'] = $acc
        $r['AccessedText'] = $accText
        $r['Reason'] = [string]$f.reason
        $r['Risk'] = [string]$f.risk
        $r['RiskRank'] = switch ([string]$f.risk) {
            '不建议' { 4 }
            '高' { 3 }
            '中' { 2 }
            default { 1 }
        }
        $r['Categories'] = [string]$f.categories
        $r['IsDuplicate'] = [bool]$f.isDuplicate
        $r['IsLarge'] = [bool]$f.isLarge
        $r['IsTemp'] = [bool]$f.isTemp
        $r['IsSameNameDiff'] = [bool]$f.isSameNameDiff
        $r['IsVersion'] = [bool]$f.isVersion
        $r['IsError'] = [bool]$f.isError
        $r['IsHardLink'] = [bool]$f.isHardLink
        $r['CanSelect'] = [bool]$f.canSelect
        $r['IsSelected'] = $false
        $r['Keep'] = $false
        $r['KeepMark'] = ''
        $r['AllCopiesIntent'] = $false
        $r['Exists'] = $true
        $r['StatusText'] = '正常'
        $r['GroupSize'] = [long]$f.groupSize
        $r['GroupPaths'] = [int]$f.groupPaths
        $r['GroupDistinct'] = [int]$f.groupDistinct
        $r['Hash'] = [string]$f.hash
        $r['Identity'] = [string]$f.identity
        $r['Error'] = [string]$f.error
        [void]$table.Rows.Add($r)
    }

    # Add scan-level errors that are not tied to a file row.
    foreach ($e in @($result.errors)) {
        if ($null -eq $e) { continue }
        if ([string]::IsNullOrWhiteSpace([string]$e.fileId)) {
            $r = $table.NewRow()
            $r['Id'] = 'ERR-' + [Guid]::NewGuid().ToString('N').Substring(0, 8)
            $r['GroupId'] = ''
            $r['Name'] = if ([string]::IsNullOrWhiteSpace([string]$e.path)) { '（未知路径）' } else { [IO.Path]::GetFileName([string]$e.path) }
            $r['Path'] = [string]$e.path
            $r['Ext'] = ''
            $r['TypeText'] = '无法确认'
            $r['Purpose'] = '无法确认'
            $r['Size'] = [long]0
            $r['SizeText'] = '0 B'
            $r['Modified'] = [DateTime]::MinValue
            $r['ModifiedText'] = '不可用'
            $r['Accessed'] = [DateTime]::MinValue
            $r['AccessedText'] = '不可用'
            $r['Reason'] = '扫描/读取失败：' + [string]$e.error
            $r['Risk'] = '不建议'
            $r['RiskRank'] = 4
            $r['Categories'] = '跳过/错误'
            $r['IsDuplicate'] = $false
            $r['IsLarge'] = $false
            $r['IsTemp'] = $false
            $r['IsSameNameDiff'] = $false
            $r['IsVersion'] = $false
            $r['IsError'] = $true
            $r['IsHardLink'] = $false
            $r['CanSelect'] = $false
            $r['IsSelected'] = $false
            $r['Keep'] = $false
            $r['KeepMark'] = ''
            $r['AllCopiesIntent'] = $false
            $r['Exists'] = [IO.File]::Exists([string]$e.path)
            $r['StatusText'] = if ($r['Exists']) { '正常' } else { '路径不可用' }
            $r['GroupSize'] = [long]0
            $r['GroupPaths'] = 0
            $r['GroupDistinct'] = 0
            $r['Hash'] = ''
            $r['Identity'] = ''
            $r['Error'] = [string]$e.error
            [void]$table.Rows.Add($r)
        }
    }
    $table.EndLoadData()

    $script:Table = $table
    $script:View = $table.DefaultView
    $table.Add_ColumnChanged([Data.DataColumnChangeEventHandler]{
        param($sender, $e)
        if ($e.Column.ColumnName -eq 'IsSelected') { Update-SelectionSummary }
    })
(Ui 'FilesGrid').ItemsSource = $script:View

    $dirs = @($table.Rows | Where-Object { -not [bool]$_['IsError'] } | ForEach-Object { [IO.Path]::GetDirectoryName([string]$_['Path']) } | Sort-Object -Unique)
$dirCombo = (Ui 'DirectoryFilter')
    $dirCombo.Items.Clear()
    [void]$dirCombo.Items.Add('全部目录')
    foreach ($d in $dirs) { [void]$dirCombo.Items.Add($d) }
    $dirCombo.SelectedIndex = 0

    $types = @($table.Rows | Where-Object { -not [bool]$_['IsError'] } | ForEach-Object { [string]$_['TypeText'] } | Sort-Object -Unique)
$typeCombo = (Ui 'TypeFilter')
    $typeCombo.Items.Clear()
    [void]$typeCombo.Items.Add('全部类型')
    foreach ($t in $types) { [void]$typeCombo.Items.Add($t) }
    $typeCombo.SelectedIndex = 0

$riskCombo = (Ui 'RiskFilter')
    $riskCombo.Items.Clear()
    foreach ($x in @('全部风险', '低', '中', '高', '不建议')) { [void]$riskCombo.Items.Add($x) }
    $riskCombo.SelectedIndex = 0

    $sizeCombo = (Ui 'SizeFilter')
    $sizeCombo.Items.Clear()
    foreach ($x in @('全部大小', '≥500 MiB', '100–500 MiB', '10–100 MiB', '<10 MiB')) { [void]$sizeCombo.Items.Add($x) }
    $sizeCombo.SelectedIndex = 0

    $timeCombo = (Ui 'TimeFilter')
    $timeCombo.Items.Clear()
    foreach ($x in @('全部时间', '修改≥1年', '修改≥180天', '修改≥90天', '最近30天修改')) { [void]$timeCombo.Items.Add($x) }
    $timeCombo.SelectedIndex = 0

$sortCombo = (Ui 'SortFilter')
    $sortCombo.Items.Clear()
    foreach ($x in @('默认排序', '名称', '大小从大到小', '修改时间从旧到新', '访问时间从旧到新')) { [void]$sortCombo.Items.Add($x) }
    $sortCombo.SelectedIndex = 0

    $stats = $result.stats
    $partial = if ([bool]$result.partial) { '（部分结果）' } else { '' }
(Ui 'ResultsTitleText').Text = "候选清单 $partial"
(Ui 'ResultsSummaryText').Text = "扫描完成：$($stats.filesScanned) 个文件 / $($stats.dirsScanned) 个目录；候选 $($stats.candidateFiles) 个；重复组 $($stats.duplicateGroups) 组；大文件 $($stats.largeFiles) 个；临时文件 $($stats.tempFiles) 个；错误/跳过 $($stats.errors) 个；用时 $($stats.elapsedSeconds)s。"
(Ui 'CategoryTabs').SelectedIndex = 0
    Set-Page 'Results'
    Apply-Filter
    Update-StatsPanel
    Update-TabCounts
    Update-SelectionSummary
    if ($script:Table.Rows.Count -le 5000) {
        Update-FileStatus
    }
    else {
        Ui('FooterText').Text = "已加载 $($script:Table.Rows.Count) 个候选；大清单已暂停自动状态检查，点「刷新状态」可后台检查。"
    }
    Update-RowDetails
}

function Load-LatestResult {
    $latest = Get-ChildItem -LiteralPath $script:DataDir -Directory -ErrorAction SilentlyContinue |
        Where-Object { [IO.File]::Exists((Join-Path $_.FullName 'result.json')) } |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1
    if ($null -eq $latest) {
        Show-Info '还没有找到已完成的扫描结果。'
        return
    }
    Load-Result -ScanDir $latest.FullName
}

function Start-Scan {
if (-not [bool]((Ui 'ConfirmScopeCheckBox').IsChecked)) {
        Show-Warn '请先勾选「我确认以上扫描范围和排除项」。'
        return
    }

    $roots = @()
    foreach ($cb in $script:RootCheckBoxes) {
        if ([bool]$cb.IsChecked) { $roots += [string]$cb.Tag }
    }
    if ($roots.Count -eq 0) {
        Show-Warn '至少需要勾选一个扫描目录。'
        return
    }

    $large = [long]0
if (-not [long]::TryParse(([string]((Ui 'LargeSizeBox').Text)).Trim(), [ref]$large) -or $large -lt 1048576) {
        Show-Warn '大文件阈值必须是至少 1048576 的字节数。'
        return
    }
    $modified = 0
if (-not [int]::TryParse(([string]((Ui 'ModifiedDaysBox').Text)).Trim(), [ref]$modified) -or $modified -lt 1) {
        Show-Warn '未修改天数必须是大于 0 的整数。'
        return
    }
    $accessed = 0
if (-not [int]::TryParse(([string]((Ui 'AccessedDaysBox').Text)).Trim(), [ref]$accessed) -or $accessed -lt 1) {
        Show-Warn '未访问天数必须是大于 0 的整数。'
        return
    }
    $tempDays = 0
if (-not [int]::TryParse(([string]((Ui 'TempDaysBox').Text)).Trim(), [ref]$tempDays) -or $tempDays -lt 1) {
        Show-Warn '临时文件天数必须是大于 0 的整数。'
        return
    }

    $scanId = 'SCAN-' + (Get-Date -Format 'yyyyMMdd-HHmmss')
    $outDir = Join-Path $script:DataDir $scanId
    [void][IO.Directory]::CreateDirectory($outDir)
    $statusPath = Join-Path $outDir 'status.json'
    $cancelPath = Join-Path $outDir 'cancel.flag'
    $resultPath = Join-Path $outDir 'result.json'
    $configPath = Join-Path $outDir 'config.json'

    $cacheEnabled = [bool]((Ui 'CacheEnabledCheckBox').IsChecked)
    $cachePath = Join-Path $script:DataDir 'hash-cache.json'

    $excludePaths = @(
        (Join-Path $env:USERPROFILE 'Documents\deepseek-harness')
    )

    $config = [ordered]@{
        scanId       = $scanId
        roots        = $roots
        excludePaths = $excludePaths
        thresholds   = [ordered]@{
            largeBytes   = $large
            modifiedDays = $modified
            accessedDays = $accessed
            tempDays     = $tempDays
        }
        outputDir    = $outDir
        statusPath   = $statusPath
        cancelPath   = $cancelPath
        resultPath   = $resultPath
        maxErrors    = 2000
        cacheEnabled = $cacheEnabled
        cachePath    = $cachePath
    }
    $config | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $configPath -Encoding UTF8

    try { if ([IO.File]::Exists($cancelPath)) { [IO.File]::Delete($cancelPath) } } catch { }

    $argLine = '-NoProfile -ExecutionPolicy Bypass -File "' + $script:ScannerPath + '" -ConfigPath "' + $configPath + '"'
    $stderrPath = Join-Path $outDir 'scanner-stderr.txt'
    try {
        [IO.File]::WriteAllText((Join-Path $outDir 'launch.txt'),
            ("File: $($script:PwshPath)`r`nArgs: $argLine`r`n"), [Text.UTF8Encoding]::new($false))
    }
    catch { }

    try {
        $script:CurrentScanDir = $outDir
        $script:ScannerProcess = Start-Process -FilePath $script:PwshPath -ArgumentList $argLine `
            -WorkingDirectory $script:BaseDir -WindowStyle Hidden -RedirectStandardError $stderrPath `
            -PassThru -ErrorAction Stop
        Reset-ScanLabels
        Set-Page 'Scan'
(Ui 'RangeStatusText').Text = "扫描批次：$scanId"
        $script:ScanTimer.Start()
    }
    catch {
        Show-Error "无法启动扫描器：`n$($_.Exception.Message)`n`n启动记录：$outDir\launch.txt"
    }
}

function Request-Cancel {
    if ([string]::IsNullOrWhiteSpace([string]$script:CurrentScanDir)) { return }
    $cancelPath = Join-Path $script:CurrentScanDir 'cancel.flag'
    try {
        [IO.File]::WriteAllText($cancelPath, 'cancel', [Text.UTF8Encoding]::new($false))
(Ui 'CancelScanButton').IsEnabled = $false
(Ui 'CancelScanButton').Content = '正在取消…'
(Ui 'ScanPhaseText').Text = '正在安全停止扫描…'
    }
    catch {
        Show-Error "无法写入取消标记：`n$($_.Exception.Message)"
    }
}

function Update-ScanStatus {
    if ([string]::IsNullOrWhiteSpace([string]$script:CurrentScanDir)) { return }
    $statusPath = Join-Path $script:CurrentScanDir 'status.json'
    $phase = ''
    if ([IO.File]::Exists($statusPath)) {
        try {
            $s = Get-Content -LiteralPath $statusPath -Raw -Encoding UTF8 | ConvertFrom-Json
            $phase = [string]$s.phase
(Ui 'ScanFilesText').Text = [string]$s.filesScanned
(Ui 'ScanDirsText').Text = [string]$s.dirsScanned
(Ui 'ScanHashedText').Text = [string]$s.hashedFiles
(Ui 'ScanErrorsText').Text = [string]$s.errors
(Ui 'ScanElapsedText').Text = "$($s.elapsedSeconds)s"
(Ui 'ScanPhaseText').Text = [string]$s.message
(Ui 'ScanCurrentPathText').Text = [string]$s.currentPath
        }
        catch { }
    }

    $exited = $false
    try { if ($null -ne $script:ScannerProcess) { $exited = $script:ScannerProcess.HasExited } } catch { $exited = $true }

    if ($phase -in @('done', 'canceled', 'error') -or $exited) {
        $script:ScanTimer.Stop()
        $resultPath = Join-Path $script:CurrentScanDir 'result.json'
        if ([IO.File]::Exists($resultPath)) {
            Load-Result -ScanDir $script:CurrentScanDir
        }
        elseif ($exited) {
            $errFile = Join-Path $script:CurrentScanDir 'scanner-stderr.txt'
            $extra = ''
            if ([IO.File]::Exists($errFile)) {
                try { $extra = "`n`n扫描器错误：`n" + (Get-Content -LiteralPath $errFile -Raw -Encoding UTF8) } catch { }
            }
            Show-Error "扫描进程已结束，但没有生成结果文件。`n请查看：$($script:CurrentScanDir)$extra"
            Set-Page 'Range'
        }
    }
}

function Show-PermanentConfirm {
    param(
        [int]$Count,
        [long]$Bytes
    )

    $w = [Windows.Window]::new()
    $w.Title = '永久删除确认'
    $w.Width = 660
    $w.Height = 420
    $w.ResizeMode = 'NoResize'
    $w.WindowStartupLocation = 'CenterOwner'
    if ($script:Window.IsVisible) { $w.Owner = $script:Window }
    $w.Background = [Windows.Media.Brushes]::White
    $script:PermanentDialog = $w
    $script:SecureErase = $false

    $grid = [Windows.Controls.Grid]::new()
    $grid.Margin = [Windows.Thickness]::new(18)
    for ($i = 0; $i -lt 5; $i++) { $grid.RowDefinitions.Add([Windows.Controls.RowDefinition]::new()) }

    $title = [Windows.Controls.TextBlock]::new()
    $title.Text = '永久删除（不会进入回收站）'
    $title.FontSize = 20
    $title.FontWeight = 'Bold'
    $title.Foreground = [Windows.Media.Brushes]::DarkRed
    [Windows.Controls.Grid]::SetRow($title, 0)
    [void]$grid.Children.Add($title)

    $msg = [Windows.Controls.TextBlock]::new()
    $msg.Text = "将永久删除 $Count 个文件，合计 $(Format-Bytes $Bytes)。`n`n" +
               "这些文件不会进入回收站，删除后无法通过本工具或回收站恢复。`n" +
               "如果勾选了「同时删除快捷方式」，匹配的快捷方式也会被永久删除。`n`n" +
               "请输入「永久删除」四个字以确认："
    $msg.TextWrapping = 'Wrap'
    $msg.LineHeight = 22
    $msg.Margin = [Windows.Thickness]::new(0, 10, 0, 10)
    [Windows.Controls.Grid]::SetRow($msg, 1)
    [void]$grid.Children.Add($msg)

    $secureCb = [Windows.Controls.CheckBox]::new()
    $secureCb.Content = '覆盖写入后删除（更安全，但大文件很慢；SSD 不能保证 100% 不可恢复）'
    $secureCb.IsChecked = $false
    $secureCb.Margin = [Windows.Thickness]::new(0, 6, 0, 6)
    $secureCb.Add_Checked({ $script:SecureErase = $true })
    $secureCb.Add_Unchecked({ $script:SecureErase = $false })
    [Windows.Controls.Grid]::SetRow($secureCb, 2)
    [void]$grid.Children.Add($secureCb)

    $tb = [Windows.Controls.TextBox]::new()
    $tb.Height = 30
    $tb.Padding = [Windows.Thickness]::new(8, 4, 8, 4)
    [Windows.Controls.Grid]::SetRow($tb, 3)
    [void]$grid.Children.Add($tb)
    $script:PermanentDialogText = $tb

    $sp = [Windows.Controls.StackPanel]::new()
    $sp.Orientation = 'Horizontal'
    $sp.HorizontalAlignment = 'Right'
    $sp.Margin = [Windows.Thickness]::new(0, 14, 0, 0)
    [Windows.Controls.Grid]::SetRow($sp, 4)

    $cancel = [Windows.Controls.Button]::new()
    $cancel.Content = '取消'
    $cancel.Padding = [Windows.Thickness]::new(14, 7, 14, 7)
    $cancel.Margin = [Windows.Thickness]::new(0, 0, 8, 0)
    $cancel.Add_Click({ $script:PermanentDialog.DialogResult = $false })

    $ok = [Windows.Controls.Button]::new()
    $ok.Content = '确认永久删除'
    $ok.Padding = [Windows.Thickness]::new(14, 7, 14, 7)
    $ok.IsEnabled = $false
    $ok.Background = [Windows.Media.Brushes]::DarkRed
    $ok.Foreground = [Windows.Media.Brushes]::White
    $ok.Add_Click({ $script:PermanentDialog.DialogResult = $true })
    $script:PermanentDialogOK = $ok

    $tb.Add_TextChanged({
        $script:PermanentDialogOK.IsEnabled = ([string]$script:PermanentDialogText.Text -eq '永久删除')
    })

    [void]$sp.Children.Add($cancel)
    [void]$sp.Children.Add($ok)
    [void]$grid.Children.Add($sp)
    $w.Content = $grid
    $result = $w.ShowDialog()
    return [bool]$result
}

function Show-Manifest {
    if ($null -eq $script:Table) { return }
    Commit-GridEdits
    $selected = @($script:Table.Rows | Where-Object { [bool]$_['IsSelected'] })
    if ($selected.Count -eq 0) {
        Show-Info '当前没有勾选任何文件。'
        return
    }

    $allGroupIds = @(
        $selected |
        Where-Object { [bool]$_['AllCopiesIntent'] -and -not [string]::IsNullOrWhiteSpace([string]$_['GroupId']) } |
        ForEach-Object { [string]$_['GroupId'] } |
        Sort-Object -Unique
    )

    foreach ($gid in $allGroupIds) {
        if (-not (Test-AllCopiesIntentComplete -GroupId $gid)) {
            Show-Warn "重复组 $gid 选择了「处理全部副本」，但并非该组所有路径都被勾选。`n请重新勾选整组，或改为仅处理选中的副本。"
            return
        }
    }

    $allGroupSet = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($gid in $allGroupIds) { [void]$allGroupSet.Add($gid) }
    $normal = @($selected | Where-Object { [string]::IsNullOrWhiteSpace([string]$_['GroupId']) -or -not $allGroupSet.Contains([string]$_['GroupId']) })
    $allRows = @($selected | Where-Object { $allGroupSet.Contains([string]$_['GroupId']) })

    $lines = [System.Collections.Generic.List[string]]::new()
    $lines.Add('本地文件清理助手 — 最终操作清单')
    $lines.Add('生成时间：' + (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))
    $lines.Add('扫描批次：' + (Split-Path -Leaf $script:CurrentScanDir))
    $lines.Add('')
    $lines.Add("待处理：$($selected.Count) 个文件")
    $totalBytes = [long]0
    foreach ($r in $selected) { $totalBytes += [long]$r['Size'] }
    $lines.Add("合计大小：$(Format-Bytes $totalBytes)")
    $keptCount = @($script:Table.Rows | Where-Object { [bool]$_['Keep'] }).Count
    if ($keptCount -gt 0) {
        $lines.Add("另外保留重复副本：$keptCount 个（不在本清单中，不会被处理）")
    }
    $lines.Add('默认目标：Windows 回收站；单文件 ≥ 2 GiB 或回收站不可用时使用本地隔离目录。')
    $lines.Add('也可以选择「永久删除（不可恢复）」：不进入回收站，需要输入确认文字。')
    if ($selected.Count -gt 5000) {
        $lines.Add('提示：大批量逐文件放入回收站会较慢；同一磁盘上「移到隔离目录」通常快很多，也便于整批恢复。')
    }
    $lines.Add('注意：本页面不会立即执行。你还需要勾选下方确认并再次点击执行。')
    $lines.Add('')

    $lines.Add('============================================================')
    $lines.Add('一、仅删除选中的副本')
    $lines.Add('============================================================')
    if ($normal.Count -eq 0) {
        $lines.Add('（无）')
    }
    else {
        foreach ($dirGroup in ($normal | Group-Object { [IO.Path]::GetDirectoryName([string]$_['Path']) } | Sort-Object Name)) {
            $lines.Add('')
            $lines.Add("[$($dirGroup.Name)]")
            foreach ($r in ($dirGroup.Group | Sort-Object { [string]$_['Name'] })) {
                $g = if ([string]::IsNullOrWhiteSpace([string]$r['GroupId'])) { '' } else { " | 组 $($r['GroupId'])" }
                $lines.Add("  • $($r['Name']) | $(Format-Bytes ([long]$r['Size']))$g")
                $lines.Add("    $($r['Path'])")
                $lines.Add("    理由：$($r['Reason'])")
            }
        }
    }

    $lines.Add('')
    $lines.Add('============================================================')
    $lines.Add('二、删除全部内容相同的副本')
    $lines.Add('============================================================')
    if ($allGroupIds.Count -eq 0) {
        $lines.Add('（无）')
    }
    else {
        foreach ($gid in $allGroupIds) {
            $rows = @($script:Table.Rows | Where-Object { [string]$_['GroupId'] -eq $gid })
            $lines.Add('')
            $lines.Add("组 $gid — 全组 $($rows.Count) 个路径，全部处理，不保留副本：")
            foreach ($r in $rows) {
                $lines.Add("  • $($r['Path']) | $(Format-Bytes ([long]$r['Size']))")
            }
        }
    }

    $lines.Add('')
    $lines.Add('============================================================')
    $lines.Add('执行前校验：每个文件会再次检查是否存在和大小；严格模式才会重新计算 SHA-256。')
    $lines.Add('失败处理：记录原因并暂停，不提升权限、不强制删除、不自动改用更激进的方式。')
    $lines.Add('日志：本地批次目录，记录路径、大小、哈希、组、结果；不记录文件正文。')

    $manifestFullPath = ''
    $text = $lines -join [Environment]::NewLine
    if ($lines.Count -gt 3000) {
        $manifestFullPath = Join-Path $script:LogDir ('MANIFEST-' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '.txt')
        [IO.File]::WriteAllText($manifestFullPath, $text, [Text.UTF8Encoding]::new($false))
        $preview = ($lines | Select-Object -First 3000) -join [Environment]::NewLine
        $text = $preview + [Environment]::NewLine + [Environment]::NewLine +
            "...（清单共 $($lines.Count) 行，对话框只显示前 3000 行）" + [Environment]::NewLine +
            "完整清单：$manifestFullPath"
    }
    $script:ManifestFullPath = $manifestFullPath
    $manifestTruncated = (-not [string]::IsNullOrWhiteSpace($manifestFullPath))

    $w = [Windows.Window]::new()
    $w.Title = '最终操作清单（尚未执行）'
    $w.Width = 1120
    $w.Height = 780
    if ($script:Window.IsVisible) { $w.Owner = $script:Window }
    $w.WindowStartupLocation = 'CenterOwner'
    $w.Background = [Windows.Media.Brushes]::White
    $script:ManifestDialog = $w
    $script:ManifestRows = $selected
    $script:ManifestAllGroupIds = $allGroupIds
    $script:ManifestTotalBytes = $totalBytes

    $grid = [Windows.Controls.Grid]::new()
    $grid.Margin = [Windows.Thickness]::new(16)
    for ($i = 0; $i -lt 4; $i++) { $grid.RowDefinitions.Add([Windows.Controls.RowDefinition]::new()) }

    $title = [Windows.Controls.TextBlock]::new()
    $title.Text = '最终操作清单（尚未执行）'
    $title.FontSize = 20
    $title.FontWeight = 'Bold'
    $title.Margin = [Windows.Thickness]::new(0, 0, 0, 10)
    [Windows.Controls.Grid]::SetRow($title, 0)
    [void]$grid.Children.Add($title)

    $tb = [Windows.Controls.TextBox]::new()
    $tb.Text = $text
    $tb.IsReadOnly = $true
    $tb.AcceptsReturn = $true
    $tb.TextWrapping = 'NoWrap'
    $tb.VerticalScrollBarVisibility = 'Auto'
    $tb.HorizontalScrollBarVisibility = 'Auto'
    $tb.FontFamily = [Windows.Media.FontFamily]::new('Consolas')
    $tb.FontSize = 12
    [Windows.Controls.Grid]::SetRow($tb, 1)
    [void]$grid.Children.Add($tb)

    $checks = [Windows.Controls.StackPanel]::new()
    [Windows.Controls.Grid]::SetRow($checks, 2)

    $strictCb = [Windows.Controls.CheckBox]::new()
    $strictCb.Content = '执行前重新校验 SHA-256（更安全，但大文件/大量文件会明显变慢；默认关闭）'
    $strictCb.IsChecked = $false
    $strictCb.Margin = [Windows.Thickness]::new(0, 4, 0, 4)
    $strictCb.Add_Checked({ $script:StrictHashCheck = $true })
    $strictCb.Add_Unchecked({ $script:StrictHashCheck = $false })
    [void]$checks.Children.Add($strictCb)
    $script:StrictHashCheck = $false

    $shortcutCb = [Windows.Controls.CheckBox]::new()
    $shortcutCb.Content = '永久删除时，同时删除指向这些文件的快捷方式（仅精确匹配目标路径）'
    $shortcutCb.IsChecked = $true
    $shortcutCb.Margin = [Windows.Thickness]::new(0, 4, 0, 4)
    $shortcutCb.Add_Checked({ $script:DeleteShortcuts = $true })
    $shortcutCb.Add_Unchecked({ $script:DeleteShortcuts = $false })
    [void]$checks.Children.Add($shortcutCb)
    $script:DeleteShortcuts = $true

    $cb = [Windows.Controls.CheckBox]::new()
    if ($manifestTruncated) {
        $cb.Content = '我已打开并核对完整清单文件，理解「全部副本」不会保留任何一份'
    }
    else {
        $cb.Content = '我已逐项核对以上路径和操作类型，并理解「全部副本」不会保留任何一份'
    }
    $cb.Margin = [Windows.Thickness]::new(0, 4, 0, 8)
    [void]$checks.Children.Add($cb)
    [void]$grid.Children.Add($checks)

    $sp = [Windows.Controls.StackPanel]::new()
    $sp.Orientation = 'Horizontal'
    $sp.HorizontalAlignment = 'Right'
    [Windows.Controls.Grid]::SetRow($sp, 3)

    if ($manifestTruncated) {
        $openFull = [Windows.Controls.Button]::new()
        $openFull.Content = '打开完整清单'
        $openFull.Padding = [Windows.Thickness]::new(14, 7, 14, 7)
        $openFull.Margin = [Windows.Thickness]::new(0, 0, 8, 0)
        $openFull.Add_Click({ Open-Path -Path $script:ManifestFullPath })
        [void]$sp.Children.Add($openFull)
    }

    $back = [Windows.Controls.Button]::new()
    $back.Content = '返回修改'
    $back.Padding = [Windows.Thickness]::new(14, 7, 14, 7)
    $back.Margin = [Windows.Thickness]::new(0, 0, 8, 0)
    $back.Add_Click({ $script:ManifestDialog.Close() })

    $recycle = [Windows.Controls.Button]::new()
    $recycle.Content = '移到回收站'
    $recycle.Padding = [Windows.Thickness]::new(14, 7, 14, 7)
    $recycle.Margin = [Windows.Thickness]::new(0, 0, 8, 0)
    $recycle.IsEnabled = $false

    $quarantine = [Windows.Controls.Button]::new()
    $quarantine.Content = '移到隔离目录'
    $quarantine.Padding = [Windows.Thickness]::new(14, 7, 14, 7)
    $quarantine.Margin = [Windows.Thickness]::new(0, 0, 8, 0)
    $quarantine.IsEnabled = $false

    $permanent = [Windows.Controls.Button]::new()
    $permanent.Content = '永久删除（不可恢复）'
    $permanent.Padding = [Windows.Thickness]::new(14, 7, 14, 7)
    $permanent.IsEnabled = $false
    $permanent.Background = [Windows.Media.Brushes]::DarkRed
    $permanent.Foreground = [Windows.Media.Brushes]::White

    $script:ManifestRecycleButton = $recycle
    $script:ManifestQuarantineButton = $quarantine
    $script:ManifestPermanentButton = $permanent

    $cb.Add_Checked({
        $script:ManifestRecycleButton.IsEnabled = $true
        $script:ManifestQuarantineButton.IsEnabled = $true
        $script:ManifestPermanentButton.IsEnabled = $true
    })
    $cb.Add_Unchecked({
        $script:ManifestRecycleButton.IsEnabled = $false
        $script:ManifestQuarantineButton.IsEnabled = $false
        $script:ManifestPermanentButton.IsEnabled = $false
    })

    $recycle.Add_Click({
        $msg = "将处理 $($script:ManifestRows.Count) 个文件，合计 $(Format-Bytes $script:ManifestTotalBytes)。`n`n默认移到 Windows 回收站；单文件 ≥ 2 GiB 或回收站不可用时使用隔离目录。`n`n确认执行？"
        $ans = [Windows.MessageBox]::Show($script:ManifestDialog, $msg, '二次确认：执行清理',
            [Windows.MessageBoxButton]::YesNo, [Windows.MessageBoxImage]::Warning)
        if ($ans -ne [Windows.MessageBoxResult]::Yes) { return }
        $script:ManifestRecycleButton.IsEnabled = $false
        $script:ManifestQuarantineButton.IsEnabled = $false
        $rowsToRun = $script:ManifestRows
        $groupsToRun = $script:ManifestAllGroupIds
        $script:ManifestDialog.Close()
        Invoke-Cleanup -Mode 'Recycle' -Rows $rowsToRun -AllGroupIds $groupsToRun
    })

    $quarantine.Add_Click({
        $msg = "将处理 $($script:ManifestRows.Count) 个文件，合计 $(Format-Bytes $script:ManifestTotalBytes)。`n`n全部移动到本地隔离目录。`n`n确认执行？"
        $ans = [Windows.MessageBox]::Show($script:ManifestDialog, $msg, '二次确认：执行清理',
            [Windows.MessageBoxButton]::YesNo, [Windows.MessageBoxImage]::Warning)
        if ($ans -ne [Windows.MessageBoxResult]::Yes) { return }
        $script:ManifestRecycleButton.IsEnabled = $false
        $script:ManifestQuarantineButton.IsEnabled = $false
        $rowsToRun = $script:ManifestRows
        $groupsToRun = $script:ManifestAllGroupIds
        $script:ManifestDialog.Close()
        Invoke-Cleanup -Mode 'Quarantine' -Rows $rowsToRun -AllGroupIds $groupsToRun
    })

    $permanent.Add_Click({
        if (-not (Show-PermanentConfirm -Count $script:ManifestRows.Count -Bytes $script:ManifestTotalBytes)) { return }
        $script:ManifestRecycleButton.IsEnabled = $false
        $script:ManifestQuarantineButton.IsEnabled = $false
        $script:ManifestPermanentButton.IsEnabled = $false
        $rowsToRun = $script:ManifestRows
        $groupsToRun = $script:ManifestAllGroupIds
        $script:ManifestDialog.Close()
        Invoke-Cleanup -Mode 'Permanent' -Rows $rowsToRun -AllGroupIds $groupsToRun
    })

    [void]$sp.Children.Add($back)
    [void]$sp.Children.Add($recycle)
    [void]$sp.Children.Add($quarantine)
    [void]$sp.Children.Add($permanent)
    [void]$grid.Children.Add($sp)
    $w.Content = $grid
    [void]$w.ShowDialog()
}

function ConvertTo-HexLower {
    param([byte[]]$Bytes)
    $sb = [Text.StringBuilder]::new($Bytes.Length * 2)
    foreach ($b in $Bytes) { [void]$sb.Append($b.ToString('x2')) }
    return $sb.ToString()
}

function Get-Sha256HexSimple {
    param([string]$Path)
    $sha = [Security.Cryptography.SHA256]::Create()
    $fs = $null
    try {
        $fs = [IO.File]::OpenRead($Path)
        $hash = $sha.ComputeHash($fs)
        return (ConvertTo-HexLower -Bytes $hash)
    }
    finally {
        if ($fs) { $fs.Dispose() }
        $sha.Dispose()
    }
}

function Send-ToRecycleBin {
    param([string]$Path)
    [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteFile(
        $Path,
        [Microsoft.VisualBasic.FileIO.UIOption]::OnlyErrorDialogs,
        [Microsoft.VisualBasic.FileIO.RecycleOption]::SendToRecycleBin,
        [Microsoft.VisualBasic.FileIO.UICancelOption]::DoNothing)
    if ([IO.File]::Exists($Path)) { throw '文件仍在原位置，回收站操作未完成' }
}

function Sanitize-FileName {
    param([string]$Name)
    $bad = [IO.Path]::GetInvalidFileNameChars()
    $sb = [Text.StringBuilder]::new()
    foreach ($ch in $Name.ToCharArray()) {
        if ($bad -contains $ch) { [void]$sb.Append('_') } else { [void]$sb.Append($ch) }
    }
    $s = $sb.ToString()
    if ($s.Length -gt 120) { $s = $s.Substring(0, 120) }
    if ([string]::IsNullOrWhiteSpace($s)) { $s = 'file' }
    return $s
}

function Move-ToQuarantine {
    param([string]$Source, [string]$BatchDir, [int]$Index)
    $name = [IO.Path]::GetFileName($Source)
    $safe = Sanitize-FileName $name
    $dst = Join-Path $BatchDir ('{0:D4}_{1}' -f $Index, $safe)
    try {
        [IO.File]::Move($Source, $dst)
        return $dst
    }
    catch {
        # Cross-volume fallback: copy, verify SHA-256, only then remove source.
        [IO.File]::Copy($Source, $dst, $false)
        $h1 = Get-Sha256HexSimple -Path $Source
        $h2 = Get-Sha256HexSimple -Path $dst
        if ($h1 -ne $h2) {
            try { [IO.File]::Delete($dst) } catch { }
            throw '复制到隔离目录后哈希校验失败，源文件未删除'
        }
        [IO.File]::Delete($Source)
        return $dst
    }
}

function Invoke-UiPump {
    try {
        $script:UiPumpFrame = [Windows.Threading.DispatcherFrame]::new()
        [Windows.Threading.Dispatcher]::CurrentDispatcher.BeginInvoke(
            [Windows.Threading.DispatcherPriority]::Background,
            [Action]{ $script:UiPumpFrame.Continue = $false }) | Out-Null
        [Windows.Threading.Dispatcher]::PushFrame($script:UiPumpFrame)
    }
    catch { }
}

function Show-CleanupProgress {
    param([int]$Total)
    if (-not $script:Window.IsVisible) { return }
    $script:CleanupCancel = $false
    $script:CleanupDone = $false

    $w = [Windows.Window]::new()
    $w.Title = '正在执行清理'
    $w.Width = 560
    $w.Height = 190
    $w.ResizeMode = 'NoResize'
    $w.WindowStartupLocation = 'CenterOwner'
    $w.Owner = $script:Window
    $w.Background = [Windows.Media.Brushes]::White
    $script:CleanupProgressWindow = $w

    $grid = [Windows.Controls.Grid]::new()
    $grid.Margin = [Windows.Thickness]::new(18)
    for ($i = 0; $i -lt 4; $i++) { $grid.RowDefinitions.Add([Windows.Controls.RowDefinition]::new()) }

    $title = [Windows.Controls.TextBlock]::new()
    $title.Text = "正在处理 $Total 个文件"
    $title.FontSize = 16
    $title.FontWeight = 'Bold'
    [Windows.Controls.Grid]::SetRow($title, 0)
    [void]$grid.Children.Add($title)

    $pb = [Windows.Controls.ProgressBar]::new()
    $pb.Minimum = 0
    $pb.Maximum = [Math]::Max(1, $Total)
    $pb.Height = 22
    $pb.Margin = [Windows.Thickness]::new(0, 12, 0, 8)
    [Windows.Controls.Grid]::SetRow($pb, 1)
    [void]$grid.Children.Add($pb)
    $script:CleanupProgressBar = $pb

    $txt = [Windows.Controls.TextBlock]::new()
    $txt.Text = '准备中…'
    $txt.Foreground = [Windows.Media.Brushes]::DimGray
    $txt.TextWrapping = 'Wrap'
    [Windows.Controls.Grid]::SetRow($txt, 2)
    [void]$grid.Children.Add($txt)
    $script:CleanupProgressText = $txt

    $sp = [Windows.Controls.StackPanel]::new()
    $sp.Orientation = 'Horizontal'
    $sp.HorizontalAlignment = 'Right'
    [Windows.Controls.Grid]::SetRow($sp, 3)
    $cancel = [Windows.Controls.Button]::new()
    $cancel.Content = '停止'
    $cancel.Padding = [Windows.Thickness]::new(14, 6, 14, 6)
    $cancel.Margin = [Windows.Thickness]::new(0, 10, 0, 0)
    $script:CleanupCancelButton = $cancel
    $cancel.Add_Click({
        $script:CleanupCancel = $true
        $script:CleanupCancelButton.IsEnabled = $false
        $script:CleanupCancelButton.Content = '正在停止…'
    })
    [void]$sp.Children.Add($cancel)
    [void]$grid.Children.Add($sp)

    $w.Content = $grid
    $w.Add_Closing({
        param($sender, $e)
        if (-not $script:CleanupDone) {
            $script:CleanupCancel = $true
            $e.Cancel = $true
        }
    })
    $script:Window.IsHitTestVisible = $false
    $w.Show()
    Invoke-UiPump
}

function Update-CleanupProgress {
    param([int]$Index, [int]$Total, [int]$Success, [int]$Failed, [string]$Current)
    if ($null -eq $script:CleanupProgressBar) { return }
    $script:CleanupProgressBar.Value = [Math]::Min($Index, $script:CleanupProgressBar.Maximum)
    $script:CleanupProgressText.Text = "$Index / $Total | 成功 $Success | 失败 $Failed`n$Current"
    Invoke-UiPump
}

function Close-CleanupProgress {
    $script:CleanupDone = $true
    if ($null -ne $script:CleanupProgressWindow) {
        try { $script:CleanupProgressWindow.Close() } catch { }
        $script:CleanupProgressWindow = $null
    }
    $script:CleanupProgressBar = $null
    $script:CleanupProgressText = $null
    if ($script:Window.IsVisible) { $script:Window.IsHitTestVisible = $true }
}

function Get-RegValueCount {
    param([string]$SubKey)
    try {
        $k = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey($SubKey)
        if ($null -eq $k) { return 0 }
        $c = $k.ValueCount
        $k.Close()
        return $c
    }
    catch { return 0 }
}

function Get-OfficeMruKeys {
    $out = [System.Collections.Generic.List[string]]::new()
    foreach ($ver in @('12.0','14.0','15.0','16.0')) {
        foreach ($app in @('Word','Excel','PowerPoint')) {
            foreach ($leaf in @('File MRU','User MRU')) {
                $out.Add("Software\Microsoft\Office\$ver\$app\$leaf")
            }
        }
    }
    return $out
}

function Get-TracePreview {
    $items = [System.Collections.Generic.List[object]]::new()

    foreach ($k in (Get-OfficeMruKeys)) {
        $c = Get-RegValueCount -SubKey $k
        if ($c -gt 0) { $items.Add([pscustomobject]@{ Category = 'Office 最近使用'; Kind = 'Registry'; Target = $k; Count = $c }) }
    }

    $explorerKeys = @(
        'Software\Microsoft\Windows\CurrentVersion\Explorer\RecentDocs',
        'Software\Microsoft\Windows\CurrentVersion\Explorer\RunMRU',
        'Software\Microsoft\Windows\CurrentVersion\Explorer\ComDlg32\OpenSavePidlMRU',
        'Software\Microsoft\Windows\CurrentVersion\Explorer\ComDlg32\LastVisitedPidlMRU',
        'Software\Microsoft\Windows\CurrentVersion\Explorer\WordWheelQuery',
        'Software\Classes\Local Settings\Software\Microsoft\Windows\Shell\MuiCache'
    )
    foreach ($k in $explorerKeys) {
        $c = Get-RegValueCount -SubKey $k
        if ($c -gt 0) { $items.Add([pscustomobject]@{ Category = '资源管理器历史'; Kind = 'Registry'; Target = $k; Count = $c }) }
    }

    $recentDir = Join-Path $env:APPDATA 'Microsoft\Windows\Recent'
    if ([IO.Directory]::Exists($recentDir)) {
        $lnk = @(Get-ChildItem -LiteralPath $recentDir -Filter '*.lnk' -File -ErrorAction SilentlyContinue).Count
        if ($lnk -gt 0) { $items.Add([pscustomobject]@{ Category = 'Windows 最近使用'; Kind = 'File'; Target = $recentDir; Count = $lnk }) }
        $jump = @(Get-ChildItem -LiteralPath $recentDir -File -ErrorAction SilentlyContinue | Where-Object { $_.Name -like '*.automaticDestinations-ms' -or $_.Name -like '*.customDestinations-ms' }).Count
        if ($jump -gt 0) { $items.Add([pscustomobject]@{ Category = 'Jump List'; Kind = 'File'; Target = $recentDir; Count = $jump }) }
    }

    return ,$items
}

function Backup-TraceRegistry {
    param([string]$BackupDir)
    $exported = 0
    $keys = @(
        'HKCU\Software\Microsoft\Office\12.0\Word',
        'HKCU\Software\Microsoft\Office\12.0\Excel',
        'HKCU\Software\Microsoft\Office\12.0\PowerPoint',
        'HKCU\Software\Microsoft\Office\14.0\Word',
        'HKCU\Software\Microsoft\Office\14.0\Excel',
        'HKCU\Software\Microsoft\Office\14.0\PowerPoint',
        'HKCU\Software\Microsoft\Office\15.0\Word',
        'HKCU\Software\Microsoft\Office\15.0\Excel',
        'HKCU\Software\Microsoft\Office\15.0\PowerPoint',
        'HKCU\Software\Microsoft\Office\16.0\Word',
        'HKCU\Software\Microsoft\Office\16.0\Excel',
        'HKCU\Software\Microsoft\Office\16.0\PowerPoint',
        'HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\RecentDocs',
        'HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\RunMRU',
        'HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\ComDlg32',
        'HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\WordWheelQuery',
        'HKCU\Software\Classes\Local Settings\Software\Microsoft\Windows\Shell\MuiCache'
    )
    foreach ($k in $keys) {
        $file = Join-Path $BackupDir ((($k -replace '[\\:]','_') -replace ' ', '') + '.reg')
        $tmp = "$file.tmp"
        try {
            $p = Start-Process -FilePath 'reg.exe' -ArgumentList @('export', "`"$k`"", "`"$tmp`"", '/y') -Wait -PassThru -NoNewWindow -RedirectStandardError (Join-Path $BackupDir 'reg.err')
            if ([IO.File]::Exists($tmp) -and (Get-Item $tmp).Length -gt 0) {
                [IO.File]::Move($tmp, $file)
                $exported++
            }
            elseif ([IO.File]::Exists($tmp)) { Remove-Item -LiteralPath $tmp -Force }
        }
        catch { }
    }
    return $exported
}

function Clear-RegistryValues {
    param([string[]]$SubKeys)
    $cleared = 0
    foreach ($sk in $SubKeys) {
        try {
            $k = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey($sk, $true)
            if ($null -eq $k) { continue }
            foreach ($vn in @($k.GetValueNames())) {
                $k.DeleteValue($vn, $false)
                $cleared++
            }
            $k.Close()
        }
        catch { }
    }
    return $cleared
}

function Invoke-TraceCleanup {
    param(
        [switch]$OfficeRecent,
        [switch]$WindowsRecent,
        [switch]$JumpLists,
        [switch]$ExplorerMru
    )

    $backupDir = Join-Path $script:LogDir ("TraceBackup-" + (Get-Date -Format 'yyyyMMdd-HHmmss'))
    [void][IO.Directory]::CreateDirectory($backupDir)
    $backupCount = Backup-TraceRegistry -BackupDir $backupDir

    $regKeys = [System.Collections.Generic.List[string]]::new()
    if ($OfficeRecent) { foreach ($k in (Get-OfficeMruKeys)) { $regKeys.Add($k) } }
    if ($ExplorerMru) {
        $regKeys.Add('Software\Microsoft\Windows\CurrentVersion\Explorer\RecentDocs')
        $regKeys.Add('Software\Microsoft\Windows\CurrentVersion\Explorer\RunMRU')
        $regKeys.Add('Software\Microsoft\Windows\CurrentVersion\Explorer\ComDlg32\OpenSavePidlMRU')
        $regKeys.Add('Software\Microsoft\Windows\CurrentVersion\Explorer\ComDlg32\LastVisitedPidlMRU')
        $regKeys.Add('Software\Microsoft\Windows\CurrentVersion\Explorer\WordWheelQuery')
        $regKeys.Add('Software\Classes\Local Settings\Software\Microsoft\Windows\Shell\MuiCache')
    }
    $clearedValues = Clear-RegistryValues -SubKeys $regKeys.ToArray()

    $filesDeleted = 0
    $recentDir = Join-Path $env:APPDATA 'Microsoft\Windows\Recent'
    if ([IO.Directory]::Exists($recentDir)) {
        if ($WindowsRecent) {
            foreach ($f in @(Get-ChildItem -LiteralPath $recentDir -Filter '*.lnk' -File -ErrorAction SilentlyContinue)) {
                try { [IO.File]::Delete($f.FullName); $filesDeleted++ } catch { }
            }
        }
        if ($JumpLists) {
            foreach ($f in @(Get-ChildItem -LiteralPath $recentDir -File -ErrorAction SilentlyContinue | Where-Object { $_.Name -like '*.automaticDestinations-ms' -or $_.Name -like '*.customDestinations-ms' })) {
                try { [IO.File]::Delete($f.FullName); $filesDeleted++ } catch { }
            }
        }
    }

    $manifest = [ordered]@{
        createdAt      = (Get-Date).ToString('o')
        backupDir      = $backupDir
        backupRegFiles = $backupCount
        registryValues = $clearedValues
        filesDeleted   = $filesDeleted
    }
    try { [IO.File]::WriteAllText((Join-Path $backupDir 'trace-manifest.json'), ($manifest | ConvertTo-Json -Depth 5), [Text.UTF8Encoding]::new($false)) } catch { }
    return $manifest
}

function Show-TraceCleanupDialog {
    $items = @(Get-TracePreview)
    $w = [Windows.Window]::new()
    $w.Title = '痕迹清理（先备份注册表）'
    $w.Width = 660
    $w.Height = 520
    $w.ResizeMode = 'NoResize'
    $w.WindowStartupLocation = 'CenterOwner'
    if ($script:Window.IsVisible) { $w.Owner = $script:Window }
    $w.Background = [Windows.Media.Brushes]::White

    $grid = [Windows.Controls.Grid]::new()
    $grid.Margin = [Windows.Thickness]::new(18)
    for ($i = 0; $i -lt 3; $i++) { $grid.RowDefinitions.Add([Windows.Controls.RowDefinition]::new()) }

    $title = [Windows.Controls.TextBlock]::new()
    $title.Text = '这些入口不会导致报错，只是残留历史；清理前会先备份注册表。'
    $title.TextWrapping = 'Wrap'
    $title.Foreground = [Windows.Media.Brushes]::DimGray
    $title.Margin = [Windows.Thickness]::new(0, 0, 0, 10)
    [Windows.Controls.Grid]::SetRow($title, 0)
    [void]$grid.Children.Add($title)

    $body = [Windows.Controls.StackPanel]::new()
    [Windows.Controls.Grid]::SetRow($body, 1)

    $officeCb = [Windows.Controls.CheckBox]::new()
    $officeCb.Content = 'Word / Excel / PowerPoint 最近使用（注册表 MRU）'
    $officeCb.IsChecked = $true
    [void]$body.Children.Add($officeCb)

    $recentCb = [Windows.Controls.CheckBox]::new()
    $recentCb.Content = 'Windows 最近使用（Recent 里的 .lnk）'
    $recentCb.IsChecked = $true
    [void]$body.Children.Add($recentCb)

    $jumpCb = [Windows.Controls.CheckBox]::new()
    $jumpCb.Content = 'Jump List（任务栏/开始菜单最近）'
    $jumpCb.IsChecked = $true
    [void]$body.Children.Add($jumpCb)

    $mruCb = [Windows.Controls.CheckBox]::new()
    $mruCb.Content = '资源管理器历史（RecentDocs / RunMRU / MuiCache 等）'
    $mruCb.IsChecked = $true
    [void]$body.Children.Add($mruCb)

    $summary = [Windows.Controls.TextBlock]::new()
    $summary.Text = "扫描到 $(($items | Measure-Object Count -Sum).Sum) 条历史项（Office/资源管理器/最近使用/Jump List）。"
    $summary.TextWrapping = 'Wrap'
    $summary.Margin = [Windows.Thickness]::new(0, 10, 0, 10)
    [void]$body.Children.Add($summary)

    [void]$grid.Children.Add($body)

    $sp = [Windows.Controls.StackPanel]::new()
    $sp.Orientation = 'Horizontal'
    $sp.HorizontalAlignment = 'Right'
    [Windows.Controls.Grid]::SetRow($sp, 2)
    $cancel = [Windows.Controls.Button]::new()
    $cancel.Content = '取消'
    $cancel.Padding = [Windows.Thickness]::new(14, 7, 14, 7)
    $cancel.Margin = [Windows.Thickness]::new(0, 0, 8, 0)
    $cancel.Add_Click({ $script:TraceCleanupDialog.Close() })
    $run = [Windows.Controls.Button]::new()
    $run.Content = '开始清理'
    $run.Padding = [Windows.Thickness]::new(14, 7, 14, 7)
    $run.Background = [Windows.Media.Brushes]::DarkRed
    $run.Foreground = [Windows.Media.Brushes]::White
    $script:TraceCleanupDialog = $w
    $script:TraceOfficeCb = $officeCb
    $script:TraceRecentCb = $recentCb
    $script:TraceJumpCb = $jumpCb
    $script:TraceMruCb = $mruCb
    $run.Add_Click({
        $ans = [Windows.MessageBox]::Show($script:TraceCleanupDialog, '确认清理以上选中的痕迹？注册表会先备份到日志目录。', '确认痕迹清理',
            [Windows.MessageBoxButton]::YesNo, [Windows.MessageBoxImage]::Warning)
        if ($ans -ne [Windows.MessageBoxResult]::Yes) { return }
        $script:TraceCleanupDialog.Close()
        $m = Invoke-TraceCleanup -OfficeRecent:$script:TraceOfficeCb.IsChecked -WindowsRecent:$script:TraceRecentCb.IsChecked -JumpLists:$script:TraceJumpCb.IsChecked -ExplorerMru:$script:TraceMruCb.IsChecked
        Show-Info "痕迹清理完成。`n注册表值：清理 $($m.registryValues) 个`n文件：删除 $($m.filesDeleted) 个`n备份：$($m.backupRegFiles) 个 .reg`n`n备份目录：$($m.backupDir)"
    })
    [void]$sp.Children.Add($cancel)
    [void]$sp.Children.Add($run)
    [void]$grid.Children.Add($sp)
    $w.Content = $grid
    [void]$w.ShowDialog()
}

function Show-HistoryDialog {
    $w = [Windows.Window]::new()
    $w.Title = '历史记录'
    $w.Width = 760
    $w.Height = 500
    $w.WindowStartupLocation = 'CenterOwner'
    if ($script:Window.IsVisible) { $w.Owner = $script:Window }
    $w.Background = [Windows.Media.Brushes]::White

    $grid = [Windows.Controls.Grid]::new()
    $grid.Margin = [Windows.Thickness]::new(18)
    $grid.RowDefinitions.Add([Windows.Controls.RowDefinition]::new())
    $grid.RowDefinitions.Add([Windows.Controls.RowDefinition]::new())
    $grid.RowDefinitions.Add([Windows.Controls.RowDefinition]::new())

    $title = [Windows.Controls.TextBlock]::new()
    $title.Text = '以下为本地批次记录（来自 logs 目录）'
    $title.FontSize = 15
    $title.FontWeight = 'Bold'
    $title.Margin = [Windows.Thickness]::new(0, 0, 0, 8)
    [Windows.Controls.Grid]::SetRow($title, 0)
    [void]$grid.Children.Add($title)

    $scroll = [Windows.Controls.ScrollViewer]::new()
    $scroll.VerticalScrollBarVisibility = 'Auto'
    [Windows.Controls.Grid]::SetRow($scroll, 1)
    $txt = [Windows.Controls.TextBlock]::new()
    $txt.TextWrapping = 'Wrap'
    $txt.FontFamily = [Windows.Media.FontFamily]::new('Consolas')

    $sb = [Text.StringBuilder]::new()
    $dirs = @(Get-ChildItem -LiteralPath $script:LogDir -Directory -Filter 'BATCH-*' -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending)
    if ($dirs.Count -eq 0) {
        [void]$sb.AppendLine('（暂无批次记录）')
    }
    foreach ($d in $dirs) {
        $m = $null
        $mf = Join-Path $d.FullName 'manifest.json'
        if ([IO.File]::Exists($mf)) { try { $m = Get-Content -LiteralPath $mf -Raw -Encoding UTF8 | ConvertFrom-Json } catch { } }
        $mode = if ($m) { [string]$m.mode } else { '?' }
        $ok = if ($m) { [int]$m.success } else { 0 }
        $fail = if ($m) { [int]$m.failed } else { 0 }
        $lnk = if ($m -and $null -ne $m.shortcutsRemoved) { " | 快捷方式删除 $([int]$m.shortcutsRemoved)" } else { '' }
        [void]$sb.AppendLine("$($d.Name)  模式=$mode  成功=$ok 失败=$fail$lnk")
    }
    $txt.Text = $sb.ToString()
    $scroll.Content = $txt
    [void]$grid.Children.Add($scroll)

    $sp = [Windows.Controls.StackPanel]::new()
    $sp.Orientation = 'Horizontal'
    $sp.HorizontalAlignment = 'Right'
    $sp.Margin = [Windows.Thickness]::new(0, 10, 0, 0)
    [Windows.Controls.Grid]::SetRow($sp, 2)
    $open = [Windows.Controls.Button]::new()
    $open.Content = '打开日志目录'
    $open.Padding = [Windows.Thickness]::new(14, 7, 14, 7)
    $open.Margin = [Windows.Thickness]::new(0, 0, 8, 0)
    $open.Add_Click({ Open-Path -Path $script:LogDir })
    $close = [Windows.Controls.Button]::new()
    $close.Content = '关闭'
    $close.Padding = [Windows.Thickness]::new(14, 7, 14, 7)
    $close.Add_Click({ $script:HistoryDialog.Close() })
    $script:HistoryDialog = $w
    [void]$sp.Children.Add($open)
    [void]$sp.Children.Add($close)
    [void]$grid.Children.Add($sp)
    $w.Content = $grid
    [void]$w.ShowDialog()
}

function Erase-SecureFile {
    param([string]$Path)
    try {
        $fs = [IO.FileStream]::new(
            $Path,
            [IO.FileMode]::Open,
            [IO.FileAccess]::Write,
            [IO.FileShare]::None,
            1048576,
            [IO.FileOptions]::SequentialScan)
        try {
            $len = $fs.Length
            if ($len -gt 0) {
                $rng = [Security.Cryptography.RandomNumberGenerator]::Create()
                $buffer = New-Object byte[] 1048576
                $written = [long]0
                while ($written -lt $len) {
                    if ($script:CleanupCancel) { throw 'CANCELED' }
                    $chunk = [int][Math]::Min([long]$buffer.Length, ($len - $written))
                    $rng.GetBytes($buffer)
                    $fs.Position = $written
                    $fs.Write($buffer, 0, $chunk)
                    $written += $chunk
                }
                $rng.Dispose()
                $fs.Flush()
                $fs.SetLength(0)
            }
        }
        finally { $fs.Dispose() }
        [IO.File]::Delete($Path)
    }
    catch { throw }
}

function Remove-MatchingShortcuts {
    param(
        [string[]]$DeletedPaths,
        [string]$LogPath
    )

    $targets = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($p in @($DeletedPaths)) {
        if ([string]::IsNullOrWhiteSpace($p)) { continue }
        try { [void]$targets.Add([IO.Path]::GetFullPath($p).TrimEnd('\')) } catch { }
    }
    $result = [pscustomobject]@{ Removed = 0; Scanned = 0; Errors = [System.Collections.Generic.List[string]]::new() }
    if ($targets.Count -eq 0) { return $result }

    $roots = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $candidates = @(
        [Environment]::GetFolderPath('Desktop'),
        [Environment]::GetFolderPath('StartMenu'),
        [Environment]::GetFolderPath('CommonStartMenu'),
        (Join-Path $env:APPDATA 'Microsoft\Internet Explorer\Quick Launch'),
        (Join-Path $env:APPDATA 'Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar'),
        (Join-Path $env:APPDATA 'Microsoft\Windows\Start Menu\Programs'),
        (Join-Path $env:ProgramData 'Microsoft\Windows\Start Menu\Programs')
    )
    foreach ($r in $candidates) {
        if (-not [string]::IsNullOrWhiteSpace($r) -and [IO.Directory]::Exists($r)) { [void]$roots.Add($r) }
    }

    $ws = $null
    try { $ws = New-Object -ComObject WScript.Shell } catch { }
    if ($null -eq $ws) {
        $result.Errors.Add('无法创建 WScript.Shell，快捷方式未清理')
        return $result
    }

    $logLines = [System.Collections.Generic.List[string]]::new()
    foreach ($root in $roots) {
        $links = @(Get-ChildItem -LiteralPath $root -Filter '*.lnk' -Recurse -File -ErrorAction SilentlyContinue)
        foreach ($lnk in $links) {
            $result.Scanned++
            try {
                $raw = [string]$ws.CreateShortcut($lnk.FullName).TargetPath
                if ([string]::IsNullOrWhiteSpace($raw)) { continue }
                $target = [IO.Path]::GetFullPath([Environment]::ExpandEnvironmentVariables($raw)).TrimEnd('\')
                if ($targets.Contains($target)) {
                    [IO.File]::Delete($lnk.FullName)
                    $result.Removed++
                    $logLines.Add("removed`t$($lnk.FullName)`t$target")
                }
            }
            catch {
                $result.Errors.Add("$($lnk.FullName)：$($_.Exception.Message)")
                $logLines.Add("error`t$($lnk.FullName)`t$($_.Exception.Message)")
            }
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($LogPath)) {
        try { [IO.File]::WriteAllLines($LogPath, $logLines, [Text.UTF8Encoding]::new($false)) } catch { }
    }
    return $result
}

function Invoke-Cleanup {
    param(
        [string]$Mode,
        [object[]]$Rows,
        [string[]]$AllGroupIds,
        [switch]$Quiet
    )

    $batchId = 'BATCH-' + (Get-Date -Format 'yyyyMMdd-HHmmss')
    $batchDir = Join-Path $script:LogDir $batchId
    $quarantineBatch = Join-Path $script:QuarantineDir $batchId
    [void][IO.Directory]::CreateDirectory($batchDir)
    [void][IO.Directory]::CreateDirectory($quarantineBatch)

    $opsPath = Join-Path $batchDir 'operations.jsonl'
    $manifestPath = Join-Path $batchDir 'manifest.json'
    $operations = [System.Collections.Generic.List[object]]::new()
    $deletedPaths = [System.Collections.Generic.List[string]]::new()
    $success = 0
    $failed = 0
    $stopped = $false
    $canceled = $false
    $index = 0

    $allGroupSet = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($gid in @($AllGroupIds)) { if (-not [string]::IsNullOrWhiteSpace($gid)) { [void]$allGroupSet.Add($gid) } }

    $pathIndex = [System.Collections.Generic.Dictionary[string, object]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($row in $script:Table.Rows) { $pathIndex[[string]$row['Path']] = $row }

    $opsWriter = [IO.StreamWriter]::new($opsPath, $false, [Text.UTF8Encoding]::new($false))
    $opsWriter.AutoFlush = $false

(Ui 'FooterText').Text = '正在执行已确认的清理操作…'
    [Windows.Input.Mouse]::OverrideCursor = [Windows.Input.Cursors]::Wait
    if (-not $Quiet) { Show-CleanupProgress -Total $Rows.Count }

    try {
        foreach ($r in $Rows) {
            if ($script:CleanupCancel) { $canceled = $true; $stopped = $true; break }
            $index++
            if ($stopped) { break }

            $path = [string]$r['Path']
            $size = [long]$r['Size']
            $hash = [string]$r['Hash']
            $gid = [string]$r['GroupId']
            $actionType = if ($allGroupSet.Contains($gid)) { '全部相同副本' } else { '仅选中的副本' }
            $op = [ordered]@{
                at           = (Get-Date).ToString('o')
                batchId      = $batchId
                originalPath = $path
                targetPath   = ''
                size         = $size
                hash         = $hash
                groupId      = $gid
                actionType   = $actionType
                storage      = ''
                result       = 'pending'
                error        = ''
            }

            try {
                if (-not [IO.File]::Exists($path)) { throw '文件已不存在' }
                $fi = [IO.FileInfo]::new($path)
                if ([long]$fi.Length -ne $size) { throw '文件大小已变化' }

                if ($script:StrictHashCheck -and -not [string]::IsNullOrWhiteSpace($hash)) {
                    $nowHash = Get-Sha256HexSimple -Path $path
                    if ($nowHash -ne $hash) { throw '文件内容哈希已变化' }
                }

                $target = ''
                if ($Mode -eq 'Permanent') {
                    if ($script:SecureErase) {
                        Erase-SecureFile -Path $path
                    }
                    else {
                        [IO.File]::Delete($path)
                    }
                    if ([IO.File]::Exists($path)) { throw '文件仍在原位置，永久删除未完成' }
                    $op.storage = if ($script:SecureErase) { 'secure-erased' } else { 'permanent' }
                    $op.result = 'success'
                    $success++
                    $deletedPaths.Add($path)
                }
                elseif ($Mode -eq 'Recycle' -and $size -lt 2GB) {
                    try {
                        Send-ToRecycleBin -Path $path
                        $op.storage = 'recyclebin'
                        $op.result = 'success'
                        $success++
                    }
                    catch {
                        $op.error = '回收站失败，尝试隔离目录：' + $_.Exception.Message
                    }
                }

                if ($op.result -ne 'success') {
                    $target = Move-ToQuarantine -Source $path -BatchDir $quarantineBatch -Index $index
                    $op.targetPath = $target
                    $op.storage = 'quarantine'
                    $op.result = 'success'
                    if ([string]::IsNullOrWhiteSpace($op.error)) {
                        $op.error = if ($Mode -eq 'Recycle' -and $size -ge 2GB) { '文件过大，使用隔离目录' } else { '' }
                    }
                    $success++
                }
            }
            catch {
                $op.result = 'failed'
                $op.error = $_.Exception.Message
                $failed++
                $stopped = $true
            }

            $operations.Add([pscustomobject]$op)
            $opsWriter.WriteLine(($op | ConvertTo-Json -Depth 6 -Compress))
            if (($index % 10) -eq 0) {
                $opsWriter.Flush()
                Ui('FooterText').Text = "正在执行清理：$index / $($Rows.Count)（成功 $success，失败 $failed）…"
                if (-not $Quiet) { Update-CleanupProgress -Index $index -Total $Rows.Count -Success $success -Failed $failed -Current $path }
            }
        }
        $opsWriter.Flush()

        $shortcutResult = [pscustomobject]@{ Removed = 0; Scanned = 0; Errors = @() }
        if ($Mode -eq 'Permanent' -and $script:DeleteShortcuts -and $deletedPaths.Count -gt 0) {
            if (-not $Quiet) {
                Update-CleanupProgress -Index $index -Total $Rows.Count -Success $success -Failed $failed -Current '正在删除匹配的快捷方式…'
            }
            $shortcutResult = Remove-MatchingShortcuts -DeletedPaths $deletedPaths.ToArray() -LogPath (Join-Path $batchDir 'shortcuts.txt')
        }

        $manifest = [ordered]@{
            batchId         = $batchId
            createdAt       = (Get-Date).ToString('o')
            mode            = $Mode
            batchDir        = $batchDir
            quarantine      = $quarantineBatch
            success         = $success
            failed          = $failed
            stopped         = $stopped
            shortcutsRemoved = $shortcutResult.Removed
            shortcutsScanned = $shortcutResult.Scanned
            shortcutErrors  = @($shortcutResult.Errors)
            operations      = $operations
        }
        $manifest | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $manifestPath -Encoding UTF8

        # Remove successfully processed rows from the UI table.
        $removeRows = [System.Collections.Generic.List[object]]::new()
        foreach ($op in $operations) {
            if ($op.result -ne 'success') { continue }
            $original = [string]$op.originalPath
            if ($pathIndex.ContainsKey($original)) { $removeRows.Add($pathIndex[$original]) }
        }
        if ($removeRows.Count -gt 0) {
            $script:Table.BeginLoadData()
            try {
                foreach ($row in $removeRows) {
                    try { $script:Table.Rows.Remove($row) } catch { }
                }
            }
            finally { $script:Table.EndLoadData() }
        }
        $script:Table.AcceptChanges()
        Ui('VisibleCountText').Text = "当前显示 $($script:View.Count) 个"
        Update-StatsPanel
        Update-TabCounts
        Update-SelectionSummary

        if ($canceled) {
            if (-not $Quiet) {
                Show-Warn "已按你的操作停止。`n已处理：$success`n未处理：$($Rows.Count - $index)`n`n日志：$opsPath"
            }
        }
        elseif ($failed -gt 0) {
            if (-not $Quiet) {
                Show-Warn "批次已暂停。`n已处理：$success`n失败：$failed`n其余文件未处理。`n`n原因：$($operations[-1].error)`n`n日志：$opsPath"
            }
        }
        else {
            if (-not $Quiet) {
                $extra = ''
                if ($Mode -eq 'Permanent' -and $script:DeleteShortcuts) {
                    $extra = "`n快捷方式：扫描 $($shortcutResult.Scanned) 个，删除 $($shortcutResult.Removed) 个，失败 $($shortcutResult.Errors.Count) 个"
                }
                Show-Info "批次完成。`n已处理：$success 个文件`n模式：$Mode$extra`n`n日志：$opsPath"
            }
        }
    }
    finally {
        if ($null -ne $opsWriter) {
            try { $opsWriter.Flush(); $opsWriter.Dispose() } catch { }
        }
        if (-not $Quiet) { Close-CleanupProgress }
        [Windows.Input.Mouse]::OverrideCursor = $null
(Ui 'FooterText').Text = '本工具只在本机运行，不上传文件内容、文件名或路径。'
    }
}

function Show-Batches {
    $batches = @(
        Get-ChildItem -LiteralPath $script:LogDir -Directory -ErrorAction SilentlyContinue |
        Where-Object { [IO.File]::Exists((Join-Path $_.FullName 'manifest.json')) } |
        Sort-Object LastWriteTime -Descending
    )
    if ($batches.Count -eq 0) {
        Show-Info '还没有操作批次。'
        return
    }

    $w = [Windows.Window]::new()
    $w.Title = '批次与恢复'
    $w.Width = 1000
    $w.Height = 680
    if ($script:Window.IsVisible) { $w.Owner = $script:Window }
    $w.WindowStartupLocation = 'CenterOwner'
    $w.Background = [Windows.Media.Brushes]::White
    $script:BatchDialog = $w

    $grid = [Windows.Controls.Grid]::new()
    $grid.Margin = [Windows.Thickness]::new(16)
    $grid.ColumnDefinitions.Add([Windows.Controls.ColumnDefinition]::new())
    $grid.ColumnDefinitions.Add([Windows.Controls.ColumnDefinition]::new())
    $grid.RowDefinitions.Add([Windows.Controls.RowDefinition]::new())
    $grid.RowDefinitions.Add([Windows.Controls.RowDefinition]::new())

    $list = [Windows.Controls.ListBox]::new()
    $list.Width = 300
    $list.Margin = [Windows.Thickness]::new(0, 0, 12, 0)
    [Windows.Controls.Grid]::SetRow($list, 0)
    [Windows.Controls.Grid]::SetRowSpan($list, 2)
    [void]$grid.Children.Add($list)

    $items = [System.Collections.Generic.List[object]]::new()
    foreach ($b in $batches) {
        $items.Add([pscustomobject]@{
            Display = $b.Name
            Path    = $b.FullName
        })
    }
    $list.ItemsSource = $items
    $list.DisplayMemberPath = 'Display'

    $details = [Windows.Controls.TextBox]::new()
    $details.IsReadOnly = $true
    $details.AcceptsReturn = $true
    $details.TextWrapping = 'Wrap'
    $details.VerticalScrollBarVisibility = 'Auto'
    $details.FontFamily = [Windows.Media.FontFamily]::new('Consolas')
    $details.FontSize = 12
    [Windows.Controls.Grid]::SetColumn($details, 1)
    [Windows.Controls.Grid]::SetRow($details, 0)
    [void]$grid.Children.Add($details)

    $sp = [Windows.Controls.StackPanel]::new()
    $sp.Orientation = 'Horizontal'
    $sp.HorizontalAlignment = 'Right'
    [Windows.Controls.Grid]::SetColumn($sp, 1)
    [Windows.Controls.Grid]::SetRow($sp, 1)

    $restore = [Windows.Controls.Button]::new()
    $restore.Content = '恢复隔离区文件'
    $restore.Padding = [Windows.Thickness]::new(12, 7, 12, 7)
    $restore.Margin = [Windows.Thickness]::new(0, 8, 8, 0)
    $restore.IsEnabled = $false

    $recycle = [Windows.Controls.Button]::new()
    $recycle.Content = '打开回收站'
    $recycle.Padding = [Windows.Thickness]::new(12, 7, 12, 7)
    $recycle.Margin = [Windows.Thickness]::new(0, 8, 8, 0)

    $openLog = [Windows.Controls.Button]::new()
    $openLog.Content = '打开日志目录'
    $openLog.Padding = [Windows.Thickness]::new(12, 7, 12, 7)
    $openLog.Margin = [Windows.Thickness]::new(0, 8, 0, 0)

    $script:BatchList = $list
    $script:BatchDetails = $details
    $script:BatchRestoreButton = $restore
    $script:BatchOpenLogButton = $openLog
    $script:BatchCurrent = $null

    [void]$sp.Children.Add($restore)
    [void]$sp.Children.Add($recycle)
    [void]$sp.Children.Add($openLog)
    [void]$grid.Children.Add($sp)

    $list.Add_SelectionChanged({
        $script:BatchCurrent = $script:BatchList.SelectedItem
        $script:BatchDetails.Text = ''
        $script:BatchRestoreButton.IsEnabled = $false
        if ($null -eq $script:BatchCurrent) { return }
        $mp = Join-Path ([string]$script:BatchCurrent.Path) 'manifest.json'
        if (-not [IO.File]::Exists($mp)) { return }
        $m = Get-Content -LiteralPath $mp -Raw -Encoding UTF8 | ConvertFrom-Json
        $lines = [System.Collections.Generic.List[string]]::new()
        $lines.Add("批次：$($m.batchId)")
        $lines.Add("时间：$([DateTime]::Parse([string]$m.createdAt).ToLocalTime().ToString('yyyy-MM-dd HH:mm:ss'))")
        $lines.Add("模式：$($m.mode)")
        $lines.Add("成功：$($m.success)；失败：$($m.failed)；暂停：$($m.stopped)")
        $lines.Add('')
        foreach ($op in @($m.operations)) {
            $lines.Add("[$($op.result)] $($op.actionType) | $($op.storage)")
            $lines.Add("  原路径：$($op.originalPath)")
            if (-not [string]::IsNullOrWhiteSpace([string]$op.targetPath)) { $lines.Add("  目标：$($op.targetPath)") }
            if (-not [string]::IsNullOrWhiteSpace([string]$op.error)) { $lines.Add("  说明：$($op.error)") }
        }
        $script:BatchDetails.Text = ($lines -join [Environment]::NewLine)
        $hasQuarantine = @($m.operations | Where-Object { $_.storage -eq 'quarantine' -and $_.result -eq 'success' }).Count -gt 0
        $script:BatchRestoreButton.IsEnabled = $hasQuarantine
    })

    $restore.Add_Click({
        if ($null -eq $script:BatchCurrent) { return }
        $mp = Join-Path ([string]$script:BatchCurrent.Path) 'manifest.json'
        $m = Get-Content -LiteralPath $mp -Raw -Encoding UTF8 | ConvertFrom-Json
        $logPath = Join-Path ([string]$script:BatchCurrent.Path) ('restore-' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '.jsonl')
        $ok = 0
        $skip = 0
        $fail = 0
        foreach ($op in @($m.operations)) {
            if ($op.storage -ne 'quarantine' -or $op.result -ne 'success') { continue }
            $src = [string]$op.targetPath
            $dst = [string]$op.originalPath
            $rec = [ordered]@{
                at = (Get-Date).ToString('o')
                source = $src
                target = $dst
                result = ''
                error = ''
            }
            try {
                if (-not [IO.File]::Exists($src)) { throw '隔离文件已不存在' }
                if ([IO.File]::Exists($dst)) { throw '原路径已有文件，未覆盖' }
                $parent = [IO.Path]::GetDirectoryName($dst)
                if (-not [IO.Directory]::Exists($parent)) { [void][IO.Directory]::CreateDirectory($parent) }
                [IO.File]::Move($src, $dst)
                $rec.result = 'success'
                $ok++
            }
            catch {
                $rec.result = 'failed'
                $rec.error = $_.Exception.Message
                $fail++
            }
            [IO.File]::AppendAllText($logPath, (($rec | ConvertTo-Json -Depth 5 -Compress) + [Environment]::NewLine), [Text.UTF8Encoding]::new($false))
        }
        Show-Info "恢复完成。`n成功：$ok`n失败：$fail`n`n恢复日志：$logPath"
    })

    $recycle.Add_Click({
        try { $shell = New-Object -ComObject Shell.Application; $shell.Namespace(10).Self.InvokeVerb('open') } catch { }
    })
    $openLog.Add_Click({
        if ($null -ne $script:BatchCurrent) { Start-Process -FilePath ([string]$script:BatchCurrent.Path) | Out-Null }
    })

    $w.Content = $grid
    [void]$w.ShowDialog()
}

# ---------------------------------------------------------------------------
# Wire up UI
# ---------------------------------------------------------------------------
Add-RootCheckBoxes
Set-ExclusionText

(Ui 'ConfirmScopeCheckBox').Add_Checked({ (Ui 'StartScanButton').IsEnabled = $true })
(Ui 'ConfirmScopeCheckBox').Add_Unchecked({ (Ui 'StartScanButton').IsEnabled = $false })
(Ui 'StartScanButton').Add_Click({ Start-Scan })
(Ui 'LoadLastButton').Add_Click({ Load-LatestResult })
(Ui 'TraceCleanupButton').Add_Click({ Show-TraceCleanupDialog })
(Ui 'HistoryButton').Add_Click({ Show-HistoryDialog })
(Ui 'CancelScanButton').Add_Click({ Request-Cancel })

(Ui 'CategoryTabs').Add_SelectionChanged({ Apply-Filter })
(Ui 'SearchBox').Add_TextChanged({ Start-FilterDebounce })
(Ui 'DirectoryFilter').Add_SelectionChanged({ Apply-Filter })
(Ui 'TypeFilter').Add_SelectionChanged({ Apply-Filter })
(Ui 'RiskFilter').Add_SelectionChanged({ Apply-Filter })
(Ui 'SizeFilter').Add_SelectionChanged({ Apply-Filter })
(Ui 'TimeFilter').Add_SelectionChanged({ Apply-Filter })
(Ui 'SortFilter').Add_SelectionChanged({ Apply-Filter })
(Ui 'SelectedOnlyCheckBox').Add_Checked({ Apply-Filter })
(Ui 'SelectedOnlyCheckBox').Add_Unchecked({ Apply-Filter })
(Ui 'ApplyFilterButton').Add_Click({ Apply-Filter })
(Ui 'FilesGrid').Add_SelectionChanged({ Update-RowDetails })
(Ui 'FilesGrid').Add_CurrentCellChanged({
    try { [void](Ui 'FilesGrid').CommitEdit([Windows.Controls.DataGridEditingUnit]::Cell, $true) } catch { }
})

(Ui 'SelectNoneButton').Add_Click({
    if ($null -eq $script:Table) { return }
    $script:SuppressSelectionSummary = $true
    try {
        foreach ($r in $script:Table.Rows) {
            $r['IsSelected'] = $false
            $r['AllCopiesIntent'] = $false
        }
    }
    finally { $script:SuppressSelectionSummary = $false }
    Update-SelectionSummary
})

(Ui 'SelectFilteredButton').Add_Click({ Select-Visible })
(Ui 'SelectVisibleButton').Add_Click({ Select-Visible })
(Ui 'ClearVisibleButton').Add_Click({ Clear-Visible })
(Ui 'SelectSameFolderButton').Add_Click({ Select-SameFolder })
(Ui 'SelectSameFolderTypeButton').Add_Click({ Select-SameFolderType })
(Ui 'ClearSameFolderButton').Add_Click({ Clear-SameFolder })
(Ui 'ClearSameFolderTypeButton').Add_Click({ Clear-SameFolderType })
(Ui 'SelectImagesButton').Add_Click({ Select-AllImages })

(Ui 'KeepSelectedButton').Add_Click({ Keep-Selected })
(Ui 'SelectRestButton').Add_Click({ Select-RestInGroup })
(Ui 'AllGroupButton').Add_Click({ Select-AllGroup })
(Ui 'PreselectDuplicatesButton').Add_Click({ Preselect-Duplicates })
(Ui 'GenerateManifestButton').Add_Click({ Show-Manifest })
(Ui 'PreviewSelectedButton').Add_Click({ Preview-Selected })
(Ui 'OpenSelectedFolderButton').Add_Click({
    $item = Get-SelectedRow
    if ($null -eq $item) { return }
    Open-Folder -Path ([string]$item['Path'])
})
(Ui 'CopyPathButton').Add_Click({
    $item = Get-SelectedRow
    if ($null -eq $item) { return }
    try {
        [Windows.Clipboard]::SetText([string]$item['Path'])
        Ui('FooterText').Text = '路径已复制到剪贴板。'
    }
    catch { Show-Error "复制路径失败：`n$($_.Exception.Message)" }
})
(Ui 'RefreshStatusButton').Add_Click({ Update-FileStatus -ShowMessage })
(Ui 'OpenRecycleBinButton').Add_Click({
    try { $shell = New-Object -ComObject Shell.Application; $shell.Namespace(10).Self.InvokeVerb('open') } catch { }
})
(Ui 'BatchesButton').Add_Click({ Show-Batches })
(Ui 'OpenDataButton').Add_Click({
    try {
        $shell = New-Object -ComObject Shell.Application
        $shell.Namespace($script:BaseDir).Self.InvokeVerb('open')
    }
    catch { }
})

(Ui 'FilesGrid').Add_LoadingRow({
    param($sender, $e)
    $row = $e.Row
    $row.Add_Loaded({
        param($rowSender, $rowEvent)
        if ($rowSender.Uid -eq 'row-buttons-bound') { return }
        $buttons = @(Get-VisualChildren -Parent $rowSender -Type ([Windows.Controls.Button]))
        if ($buttons.Count -eq 0) { return }
        $rowSender.Uid = 'row-buttons-bound'
        foreach ($b in $buttons) {
            if ($b.Uid -eq 'row-button-bound') { continue }
            $b.Uid = 'row-button-bound'
            $b.Add_Click({
                param($button, $clickEvent)
                try {
                    $rowView = $button.Tag
                    if ($null -eq $rowView -or -not ($rowView -is [Data.DataRowView])) { return }
                    $path = [string]$rowView['Path']
                    $clickEvent.Handled = $true
                    if ([string]$button.Content -eq '查看文件夹') { Open-Folder -Path $path }
                }
                catch {
                    try { Show-Error "打开操作失败：`n$($_.Exception.Message)" } catch { }
                }
            })
        }
    })
})

$script:ScanTimer = [Windows.Threading.DispatcherTimer]::new()
$script:ScanTimer.Interval = [TimeSpan]::FromMilliseconds(700)
$script:ScanTimer.Add_Tick({ Update-ScanStatus })

$script:FileStatusTimer = [Windows.Threading.DispatcherTimer]::new()
$script:FileStatusTimer.Interval = [TimeSpan]::FromMilliseconds(30000)
$script:FileStatusTimer.Add_Tick({
    if ((Ui 'ResultsPage').Visibility -ne 'Visible' -or $null -eq $script:Table) { return }
    if ($null -ne $script:StatusCheckTimer -and $script:StatusCheckTimer.IsEnabled) { return }
    if ($script:Table.Rows.Count -gt 5000) { return }
    Update-FileStatus
})
$script:FileStatusTimer.Start()

$script:Window.Add_Loaded({
    if ($script:PendingDuplicateTab -and $null -ne $script:Table) {
        $script:PendingDuplicateTab = $false
        Ui('CategoryTabs').SelectedIndex = 1
        Apply-Filter
    }
})

$script:Window.Dispatcher.Add_UnhandledException([Windows.Threading.DispatcherUnhandledExceptionEventHandler]{
    param($sender, $e)
    try {
        Show-Error "界面操作发生错误，但程序已阻止自动退出：`n$($e.Exception.Message)"
    }
    catch { }
    $e.Handled = $true
})

$script:Window.Add_Closing({
    if ($null -ne $script:ScannerProcess) {
        try {
            if (-not $script:ScannerProcess.HasExited -and -not [string]::IsNullOrWhiteSpace([string]$script:CurrentScanDir)) {
                [IO.File]::WriteAllText((Join-Path $script:CurrentScanDir 'cancel.flag'), 'cancel', [Text.UTF8Encoding]::new($false))
            }
        } catch { }
    }
})

if (($LoadLatest -or $PreselectDuplicates) -and [string]::IsNullOrWhiteSpace($TestScanDir) -and -not $SmokeTest) {
    Load-LatestResult
}
if (-not [string]::IsNullOrWhiteSpace($LoadScanDir)) {
    Load-Result -ScanDir $LoadScanDir
}
if ($PreselectDuplicates -and $null -ne $script:Table) {
    Preselect-Duplicates -Silent
    $script:PendingDuplicateTab = $true
    Ui('FooterText').Text = '已预选重复副本：每组保留 1 份，其余勾选；不会执行删除。'
}

if ($TestPermanent) {
    $scanDir = $TestScanDir
    if ([string]::IsNullOrWhiteSpace($scanDir)) {
        $scanDir = (Get-ChildItem -LiteralPath $script:DataDir -Directory |
            Where-Object { [IO.File]::Exists((Join-Path $_.FullName 'result.json')) } |
            Sort-Object LastWriteTime -Descending |
            Select-Object -First 1).FullName
    }
    Load-Result -ScanDir $scanDir
    if ($null -ne $script:StatusCheckTimer) { $script:StatusCheckTimer.Stop() }

    $row = @($script:Table.Rows | Where-Object { [string]$_['Name'] -eq 'partial.crdownload' } | Select-Object -First 1)
    if ($row.Count -eq 0) {
        Write-Output 'PERMANENT_TARGET_NOT_FOUND'
        $script:Window.Close()
        exit 1
    }
    $targetPath = [string]$row[0]['Path']
    $backupPath = Join-Path $script:BaseDir 'partial.crdownload.permanent-test.bak'
    Copy-Item -LiteralPath $targetPath -Destination $backupPath -Force
    $beforeRows = $script:Table.Rows.Count
    $script:SecureErase = $true
    Invoke-Cleanup -Mode 'Permanent' -Rows @($row[0]) -AllGroupIds @() -Quiet

    $desktopShortcut = Join-Path ([Environment]::GetFolderPath('Desktop')) 'CleanupPermTest.lnk'
    Write-Output "PERMANENT_SOURCE_EXISTS=$([IO.File]::Exists($targetPath))"
    Write-Output "PERMANENT_SHORTCUT_EXISTS=$([IO.File]::Exists($desktopShortcut))"
    Write-Output "PERMANENT_ROWS_BEFORE=$beforeRows AFTER=$($script:Table.Rows.Count)"
    if ([IO.File]::Exists($backupPath)) {
        Copy-Item -LiteralPath $backupPath -Destination $targetPath -Force
        Remove-Item -LiteralPath $backupPath -Force
    }
    $script:Window.Close()
    exit 0
}

if ($TestFolder) {
    $script:TestFolderTimer = [Windows.Threading.DispatcherTimer]::new()
    $script:TestFolderTimer.Interval = [TimeSpan]::FromMilliseconds(800)
    $script:TestFolderTimer.Add_Tick({
        $script:TestFolderTimer.Stop()
        if ($null -ne $script:FolderDialog) { $script:FolderDialog.Close() }
    })
    $script:TestFolderTimer.Start()
    Show-FolderContents -Path ([Environment]::GetFolderPath('Desktop'))
    Write-Output 'FOLDER_DIALOG_OK'
    $script:Window.Close()
    exit 0
}

if ($TestTrace) {
    $testKey = 'Software\LocalFileCleanupAssistantTest'
    $k = [Microsoft.Win32.Registry]::CurrentUser.CreateSubKey($testKey)
    $k.SetValue('a', '1'); $k.SetValue('b', '2'); $k.Close()

    $preview = @(Get-TracePreview)
    Write-Output "TRACE_PREVIEW_COUNT=$($preview.Count)"
    $cleared = Clear-RegistryValues -SubKeys @($testKey)
    Write-Output "TRACE_CLEARED_VALUES=$cleared"
    [Microsoft.Win32.Registry]::CurrentUser.DeleteSubKey($testKey, $false)
    Write-Output 'TRACE_OK'
    $script:Window.Close()
    exit 0
}

if ($TestCleanup) {
    $scanDir = $TestScanDir
    if ([string]::IsNullOrWhiteSpace($scanDir)) {
        $scanDir = (Get-ChildItem -LiteralPath $script:DataDir -Directory |
            Where-Object { [IO.File]::Exists((Join-Path $_.FullName 'result.json')) } |
            Sort-Object LastWriteTime -Descending |
            Select-Object -First 1).FullName
    }
    Load-Result -ScanDir $scanDir
    if ($null -ne $script:StatusCheckTimer) { $script:StatusCheckTimer.Stop() }

    $row = @($script:Table.Rows | Where-Object { [string]$_['Name'] -eq 'partial.crdownload' } | Select-Object -First 1)
    if ($row.Count -eq 0) {
        Write-Output 'CLEANUP_TARGET_NOT_FOUND'
        $script:Window.Close()
        exit 1
    }
    $targetPath = [string]$row[0]['Path']
    $beforeRows = $script:Table.Rows.Count
    Invoke-Cleanup -Mode 'Quarantine' -Rows @($row[0]) -AllGroupIds @() -Quiet

    $stillThere = [IO.File]::Exists($targetPath)
    $qFile = Get-ChildItem -LiteralPath $script:QuarantineDir -Recurse -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -like '*partial.crdownload' } | Select-Object -First 1
    Write-Output "CLEANUP_SOURCE_EXISTS=$stillThere"
    Write-Output "CLEANUP_ROWS_BEFORE=$beforeRows AFTER=$($script:Table.Rows.Count)"
    Write-Output "CLEANUP_QUARANTINE=$($qFile.FullName)"
    if ($null -ne $qFile -and -not $stillThere) {
        Copy-Item -LiteralPath $qFile.FullName -Destination $targetPath -Force
    }
    $script:Window.Close()
    exit 0
}

if ($TestBulk) {
    $scanDir = $TestScanDir
    if ([string]::IsNullOrWhiteSpace($scanDir)) {
        $scanDir = (Get-ChildItem -LiteralPath $script:DataDir -Directory |
            Where-Object { [IO.File]::Exists((Join-Path $_.FullName 'result.json')) } |
            Sort-Object LastWriteTime -Descending |
            Select-Object -First 1).FullName
    }
    Load-Result -ScanDir $scanDir
    $script:AutoConfirmBulk = $true

    Select-Visible
    $n1 = @($script:Table.Rows | Where-Object { [bool]$_['IsSelected'] }).Count
    Write-Output "BULK_VISIBLE=$n1"

    Clear-Visible
    $n2 = @($script:Table.Rows | Where-Object { [bool]$_['IsSelected'] }).Count
    Write-Output "BULK_CLEAR=$n2"

    $first = @($script:View | Where-Object { [bool]$_['CanSelect'] -and [bool]$_['Exists'] } | Select-Object -First 1)
    if ($first.Count -gt 0) {
        $parent = [IO.Path]::GetDirectoryName([string]$first[0]['Path'])
        $folderRows = @($script:Table.Rows | Where-Object {
            -not [bool]$_['IsError'] -and [bool]$_['CanSelect'] -and [bool]$_['Exists'] -and
            ([IO.Path]::GetDirectoryName([string]$_['Path']) -eq $parent)
        })
        [void](Set-BulkSelection -Rows $folderRows -Selected $true)
        Write-Output "BULK_FOLDER=$($folderRows.Count)"
        [void](Set-BulkSelection -Rows @($script:Table.Rows) -Selected $false)
    }
    else {
        Write-Output 'BULK_FOLDER=0'
    }

    Select-AllImages
    $n3 = @($script:Table.Rows | Where-Object { [bool]$_['IsSelected'] }).Count
    Write-Output "BULK_IMAGES=$n3"

    $script:Window.Close()
    exit 0
}

if ($TestPreview) {
    $scanDir = $TestScanDir
    if ([string]::IsNullOrWhiteSpace($scanDir)) {
        $scanDir = (Get-ChildItem -LiteralPath $script:DataDir -Directory |
            Where-Object { [IO.File]::Exists((Join-Path $_.FullName 'result.json')) } |
            Sort-Object LastWriteTime -Descending |
            Select-Object -First 1).FullName
    }
    Load-Result -ScanDir $scanDir
    $target = @($script:View | Where-Object { [string]$_['Name'] -eq 'dup.txt' } | Select-Object -First 1)
    if ($target.Count -eq 0) {
        Write-Output 'PREVIEW_TARGET_NOT_FOUND'
        $script:Window.Close()
        exit 1
    }
    Ui('FilesGrid').SelectedItem = $target[0]
    $script:TestPreviewTimer = [Windows.Threading.DispatcherTimer]::new()
    $script:TestPreviewTimer.Interval = [TimeSpan]::FromMilliseconds(900)
    $script:TestPreviewTimer.Add_Tick({
        $script:TestPreviewTimer.Stop()
        if ($null -ne $script:PreviewDialog) { $script:PreviewDialog.Close() }
    })
    $script:TestPreviewTimer.Start()
    Preview-Selected
    Write-Output "PREVIEW_OPENED=$($null -ne $script:PreviewDialog)"
    $script:Window.Close()
    exit 0
}

if ($TestDialogs) {
    $scanDir = $TestScanDir
    if ([string]::IsNullOrWhiteSpace($scanDir)) {
        $scanDir = (Get-ChildItem -LiteralPath $script:DataDir -Directory |
            Where-Object { [IO.File]::Exists((Join-Path $_.FullName 'result.json')) } |
            Sort-Object LastWriteTime -Descending |
            Select-Object -First 1).FullName
    }
    Load-Result -ScanDir $scanDir
    foreach ($r in $script:Table.Rows) {
        if ([bool]$r['IsDuplicate']) { $r['IsSelected'] = $true }
    }
    $script:TestTimer = [Windows.Threading.DispatcherTimer]::new()
    $script:TestTimer.Interval = [TimeSpan]::FromMilliseconds(900)
    $script:TestTimer.Add_Tick({ $script:TestTimer.Stop(); if ($null -ne $script:ManifestDialog) { $script:ManifestDialog.Close() } })
    $script:TestTimer.Start()
    Show-Manifest
    Write-Output "MANIFEST_DIALOG_OK=$($script:ManifestRows.Count)"
    $gid = ($script:Table.Rows | Where-Object { [bool]$_['IsDuplicate'] } | Select-Object -First 1)['GroupId']
    $script:TestTimer2 = [Windows.Threading.DispatcherTimer]::new()
    $script:TestTimer2.Interval = [TimeSpan]::FromMilliseconds(700)
    $script:TestTimer2.Add_Tick({ $script:TestTimer2.Stop(); if ($null -ne $script:GroupDialog) { $script:GroupDialog.Close() } })
    $script:TestTimer2.Start()
    [void](Show-GroupConfirm -GroupId ([string]$gid))
    Write-Output 'GROUP_DIALOG_OK'
    $script:Window.Close()
    exit 0
}

if (-not [string]::IsNullOrWhiteSpace($TestScanDir)) {
    Load-Result -ScanDir $TestScanDir
    if ($null -ne $script:StatusCheckTimer) { $script:StatusCheckTimer.Stop() }
    Update-FileStatusSync
    Write-Output "TABLE_ROWS=$($script:Table.Rows.Count)"
    Write-Output "VIEW_ROWS=$($script:View.Count)"
    Write-Output "MISSING=$(@($script:Table.Rows | Where-Object { -not [bool]$_['Exists'] }).Count)"
    Write-Output "TAB=$((Ui 'CategoryTabs').SelectedItem.Tag) SORT=$($script:View.Sort)"
    Write-Output "FIRST5=$((@($script:View | Select-Object -First 5) | ForEach-Object { "$($_['Name'])|$($_['RiskRank'])|$($_['Size'])" }) -join ';')"
    (Ui 'SearchBox').Text = 'dup'
    Apply-Filter
    Write-Output "SEARCH_VIEW=$($script:View.Count)"
    (Ui 'SearchBox').Text = ''
    (Ui 'CategoryTabs').SelectedIndex = 1
    Apply-Filter
    Write-Output "DUP_VIEW=$($script:View.Count)"
    (Ui 'SizeFilter').SelectedItem = '<10 MiB'
    Apply-Filter
    Write-Output "SIZE_VIEW=$($script:View.Count)"
    (Ui 'SizeFilter').SelectedIndex = 0
    (Ui 'CategoryTabs').SelectedIndex = 0
    (Ui 'TimeFilter').SelectedItem = '修改≥180天'
    Apply-Filter
    Write-Output "TIME_VIEW=$($script:View.Count)"
    Preselect-Duplicates -Silent
    $selCount = @($script:Table.Rows | Where-Object { [bool]$_['IsSelected'] }).Count
    $keepCount = @($script:Table.Rows | Where-Object { [bool]$_['Keep'] }).Count
    Write-Output "PRESELECT_SELECTED=$selCount KEPT=$keepCount"
    (Ui 'CategoryTabs').SelectedIndex = 1
    Apply-Filter
    (Ui 'CategoryTabs').SelectedIndex = 0
    Apply-Filter
    $afterFilter = @($script:Table.Rows | Where-Object { [bool]$_['IsSelected'] }).Count
    Write-Output "PERSIST_AFTER_FILTER=$afterFilter"
    $script:Window.Close()
    exit 0
}

if ($SmokeTest) {
    $script:Window.Close()
    Write-Output 'SMOKE_OK'
    exit 0
}

[void]$script:Window.ShowDialog()





















































