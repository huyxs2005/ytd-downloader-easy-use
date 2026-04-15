param(
    [switch] $Worker,
    [switch] $NoRun,
    [string] $JobFile,
    [string] $TempRoot,
    [string] $TargetFolder,
    [string] $WorkerLogFile,
    [ValidateSet('Create', 'Update')]
    [string] $RunKind,
    [ValidateSet('Audio', 'Video')]
    [string] $Mode,
    [string] $SourceLabel,
    [string] $WorkerLabel
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
$OutputEncoding = [System.Text.UTF8Encoding]::new($false)

$yt = Join-Path $PSScriptRoot 'yt-dlp.exe'
$cookies = Join-Path $PSScriptRoot 'cookies.txt'
$ffmpegDir = Join-Path $PSScriptRoot 'bin'
$downloadRoot = Join-Path $PSScriptRoot 'All downloaded playlist'
$musicSearchSongsParams = 'EgWKAQIIAWoSEAUQBBADEAkQChAQEBEQFRAO'
$browserUserAgent = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36'
$script:LauncherScriptPath = $PSCommandPath
if ([string]::IsNullOrWhiteSpace($script:LauncherScriptPath)) {
    $script:LauncherScriptPath = Join-Path $PSScriptRoot 'download_playlist.ps1'
}

$script:LogRecords = New-Object System.Collections.Generic.List[object]

function Get-SafeFileName {
    param([Parameter(Mandatory = $true)][string] $Name)

    $safe = $Name -replace '[\\/:*?"<>|]', '_'
    $safe = $safe -replace '\s+', ' '
    $safe = $safe.Trim().TrimEnd('.')
    if ([string]::IsNullOrWhiteSpace($safe)) {
        return 'untitled'
    }
    if ($safe.Length -gt 180) {
        return $safe.Substring(0, 180).Trim()
    }
    return $safe
}

function Get-UniqueFolderPath {
    param([Parameter(Mandatory = $true)][string] $BasePath)

    if (-not (Test-Path -LiteralPath $BasePath)) {
        return $BasePath
    }

    $parent = Split-Path -Parent $BasePath
    $leaf = Split-Path -Leaf $BasePath
    for ($i = 1; $i -lt 1000; $i++) {
        $candidate = Join-Path $parent ("{0} ({1})" -f $leaf, $i)
        if (-not (Test-Path -LiteralPath $candidate)) {
            return $candidate
        }
    }

    throw "Unable to find a unique folder name for $BasePath"
}

function Get-UniqueFilePath {
    param([Parameter(Mandatory = $true)][string] $BasePath)

    if (-not (Test-Path -LiteralPath $BasePath)) {
        return $BasePath
    }

    $dir = Split-Path -Parent $BasePath
    $leaf = Split-Path -Leaf $BasePath
    $ext = [System.IO.Path]::GetExtension($leaf)
    $stem = [System.IO.Path]::GetFileNameWithoutExtension($leaf)
    for ($i = 1; $i -lt 1000; $i++) {
        $candidate = Join-Path $dir ("{0} ({1}){2}" -f $stem, $i, $ext)
        if (-not (Test-Path -LiteralPath $candidate)) {
            return $candidate
        }
    }

    throw "Unable to find a unique file name for $BasePath"
}

function Get-VideoIdSidecarPath {
    param([Parameter(Mandatory = $true)][string] $MediaPath)

    $dir = Split-Path -Parent $MediaPath
    $leaf = [System.IO.Path]::GetFileNameWithoutExtension($MediaPath)
    return (Join-Path $dir ($leaf + '.yt-dlp.id'))
}

function Write-VideoIdSidecar {
    param(
        [Parameter(Mandatory = $true)][string] $MediaPath,
        [Parameter(Mandatory = $true)][string] $VideoId
    )

    if ([string]::IsNullOrWhiteSpace($VideoId)) {
        return
    }

    Set-Content -LiteralPath (Get-VideoIdSidecarPath -MediaPath $MediaPath) -Value $VideoId -Encoding UTF8
}

function Set-ProcessTitle {
    param([Parameter(Mandatory = $true)][string] $Title)

    try {
        $Host.UI.RawUI.WindowTitle = $Title
    } catch {
    }

    try {
        [Console]::Title = $Title
    } catch {
    }
}

function Assert-Workspace {
    if (-not (Test-Path -LiteralPath $yt)) {
        throw "yt-dlp.exe not found at $yt"
    }
    if (-not (Test-Path -LiteralPath $cookies)) {
        throw "cookies.txt not found at $cookies"
    }
    if (-not (Test-Path -LiteralPath $ffmpegDir)) {
        throw "ffmpeg folder not found at $ffmpegDir"
    }
    if (-not (Test-Path -LiteralPath $downloadRoot)) {
        New-Item -ItemType Directory -Path $downloadRoot | Out-Null
    }
}

function Invoke-YtDlp {
    param(
        [Parameter(Mandatory = $true)]
        [string[]] $Arguments,
        [switch] $CaptureOutput,
        [int] $TimeoutSeconds = 600
    )

    $common = @(
        '--cookies', $cookies,
        '--ffmpeg-location', $ffmpegDir,
        '--encoding', 'utf-8'
    )

    $full = @($common + $Arguments)
    if ($CaptureOutput) {
        $output = @(& $yt @full 2>&1)
        return [pscustomobject]@{
            ExitCode  = $LASTEXITCODE
            Output    = @($output | ForEach-Object { [string]$_ })
            TimedOut  = $false
        }
    }

    & $yt @full
    return [pscustomobject]@{
        ExitCode = $LASTEXITCODE
        Output   = @()
        TimedOut = $false
    }
}

function Get-FlatEntries {
    param([Parameter(Mandatory = $true)][string] $Url)

    $result = Invoke-YtDlp -Arguments @(
        '--quiet',
        '--no-warnings',
        '--ignore-errors',
        '--flat-playlist',
        '--dump-json',
        '--no-download',
        $Url
    ) -CaptureOutput

    $entries = New-Object System.Collections.Generic.List[object]
    foreach ($line in $result.Output) {
        if ([string]::IsNullOrWhiteSpace($line)) {
            continue
        }
        try {
            $entries.Add(($line | ConvertFrom-Json))
        } catch {
            continue
        }
    }

    return $entries.ToArray()
}

function Get-VideoDetails {
    param([Parameter(Mandatory = $true)][string] $Url)

    $result = Invoke-YtDlp -Arguments @(
        '--quiet',
        '--no-warnings',
        '--ignore-errors',
        '--skip-download',
        '--dump-single-json',
        '--no-playlist',
        $Url
    ) -CaptureOutput

    $json = ($result.Output -join "`n").Trim()
    if ([string]::IsNullOrWhiteSpace($json)) {
        return $null
    }

    try {
        return $json | ConvertFrom-Json
    } catch {
        return $null
    }
}

function Test-ShortsItem {
    param($Item)

    $url = [string]$Item.webpage_url
    if ([string]::IsNullOrWhiteSpace($url)) {
        $url = [string]$Item.url
    }
    $title = [string]$Item.title
    return ($url -match '/shorts/' -or $title -match '(?i)\bshorts\b')
}

function Test-GuardTitleOverride {
    param([string] $Title)

    return ($Title -match '(?i)\b(persona|ost|soundtrack)\b')
}

function Test-GuardAuthorOverride {
    param([string] $Value)

    return ($Value -match '(?i)\b(persona|atlus)\b')
}

function Test-UnavailableVideoPlaceholder {
    param(
        [Parameter(Mandatory = $true)] $Item
    )

    $title = [string]$Item.title
    if ([string]::IsNullOrWhiteSpace($title)) {
        return $false
    }

    return ($title -match '(?i)^\s*\[?(private video|video unavailable|deleted video|removed video|this video is unavailable)\]?\s*$')
}

function Get-MusicalVerdict {
    param($Item)

    if (Test-ShortsItem -Item $Item) {
        return 'Reject'
    }

    if (Test-UnavailableVideoPlaceholder -Item $Item) {
        return 'Reject'
    }

    $channel = [string]$Item.channel
    if ([string]::IsNullOrWhiteSpace($channel)) {
        $channel = [string]$Item.uploader
    }
    $title = [string]$Item.title
    $track = [string]$Item.track
    $artist = [string]$Item.artist
    $category = [string]$Item.category

    $categories = @()
    if ($Item.categories) {
        $categories = @($Item.categories)
    }

    if (Test-GuardTitleOverride -Title $title) {
        return 'Accept'
    }

    if (
        (Test-GuardAuthorOverride -Value $channel) -or
        (Test-GuardAuthorOverride -Value $([string]$Item.uploader)) -or
        (Test-GuardAuthorOverride -Value $artist)
    ) {
        return 'Accept'
    }

    $badPattern = '(?i)\b(walkthrough|gameplay|stream|episode|trailer|reaction|guide|part|boss|playthrough|nerf|nerfed|item|items|runes)\b'
    if ($title -match $badPattern -or $channel -match $badPattern) {
        return 'Reject'
    }

    if ($title -match '(?i)\b(cover|lyrics)\b') {
        return 'Accept'
    }

    if ($channel -match '(?i)\bTopic\b') {
        return 'Accept'
    }

    if (-not [string]::IsNullOrWhiteSpace($track) -or -not [string]::IsNullOrWhiteSpace($artist)) {
        return 'Accept'
    }

    if ($category -match '(?i)\bmusic\b') {
        return 'Accept'
    }

    foreach ($itemCategory in $categories) {
        if ([string]$itemCategory -match '(?i)\bmusic\b') {
            return 'Accept'
        }
    }

    return 'Unknown'
}

function Test-MusicalMetadata {
    param($Item)

    return ((Get-MusicalVerdict -Item $Item) -eq 'Accept')
}

function Test-DownloadedMusicMetadata {
    param([Parameter(Mandatory = $true)][string] $InfoPath)

    if (-not (Test-Path -LiteralPath $InfoPath)) {
        return $false
    }

    try {
        $info = Get-Content -LiteralPath $InfoPath -Raw | ConvertFrom-Json
    } catch {
        return $false
    }

    $channel = [string]$info.channel
    if ([string]::IsNullOrWhiteSpace($channel)) {
        $channel = [string]$info.uploader
    }
    $title = [string]$info.title
    $track = [string]$info.track
    $artist = [string]$info.artist
    $category = [string]$info.category

    $categories = @()
    if ($info.categories) {
        $categories = @($info.categories)
    }

    if (Test-UnavailableVideoPlaceholder -Item $info) {
        return $false
    }

    if (Test-GuardTitleOverride -Title $title) {
        return $true
    }

    if (
        (Test-GuardAuthorOverride -Value $channel) -or
        (Test-GuardAuthorOverride -Value $([string]$info.uploader)) -or
        (Test-GuardAuthorOverride -Value $artist)
    ) {
        return $true
    }

    $badPattern = '(?i)\b(walkthrough|gameplay|stream|episode|trailer|reaction|guide|part|boss|playthrough|nerf|nerfed|item|items|runes)\b'
    if ($title -match $badPattern -or $channel -match $badPattern) {
        return $false
    }

    if ($title -match '(?i)\b(cover|lyrics)\b') {
        return $true
    }

    if ($channel -match '(?i)\bTopic\b') {
        return $true
    }

    if (-not [string]::IsNullOrWhiteSpace($track) -or -not [string]::IsNullOrWhiteSpace($artist)) {
        return $true
    }

    if ($category -match '(?i)\bmusic\b') {
        return $true
    }

    foreach ($itemCategory in $categories) {
        if ([string]$itemCategory -match '(?i)\bmusic\b') {
            return $true
        }
    }

    return $false
}

function Get-ComparableTitleTokens {
    param([Parameter(Mandatory = $true)][string] $Title)

    $clean = $Title.ToLowerInvariant()
    $clean = $clean -replace '[^a-z0-9]+', ' '
    $tokens = @($clean.Split(' ', [System.StringSplitOptions]::RemoveEmptyEntries))
    $stopWords = @(
        'official', 'video', 'audio', 'lyrics', 'lyric', 'mv', 'topic', 'feat', 'ft',
        'music', 'slowed', 'reverb', 'version', 'ost', 'ending', 'opening', 'theme',
        'remix', 'mix', 'full', 'ver', 'performance', 'special', 'explicit', 'feat.'
    )

    $filtered = New-Object System.Collections.Generic.List[string]
    foreach ($token in $tokens) {
        if ($token.Length -lt 2) {
            continue
        }
        if ($stopWords -contains $token) {
            continue
        }
        $filtered.Add($token) | Out-Null
    }

    return $filtered.ToArray()
}

function Test-FallbackTitleMatch {
    param(
        [Parameter(Mandatory = $true)][string] $SourceTitle,
        [Parameter(Mandatory = $true)][string] $CandidateTitle
    )

    $sourceClean = ([regex]::Replace($SourceTitle.ToLowerInvariant(), '[^a-z0-9]+', ' ')).Trim()
    $candidateClean = ([regex]::Replace($CandidateTitle.ToLowerInvariant(), '[^a-z0-9]+', ' ')).Trim()
    if ([string]::IsNullOrWhiteSpace($sourceClean) -or [string]::IsNullOrWhiteSpace($candidateClean)) {
        return $false
    }

    if ($sourceClean -eq $candidateClean) {
        return $true
    }

    if ($sourceClean.Contains($candidateClean) -or $candidateClean.Contains($sourceClean)) {
        return $true
    }

    $sourceTokens = @(Get-ComparableTitleTokens -Title $SourceTitle)
    $candidateTokens = @(Get-ComparableTitleTokens -Title $CandidateTitle)
    if ($sourceTokens.Count -eq 0 -or $candidateTokens.Count -eq 0) {
        return $false
    }

    $shared = 0
    foreach ($token in $sourceTokens) {
        if ($candidateTokens -contains $token) {
            $shared++
        }
    }

    if ($sourceTokens.Count -le 1) {
        return ($shared -ge 1)
    }

    return (($shared / [double]$sourceTokens.Count) -ge 0.6)
}

function Get-NormalizedIdentity {
    param([string] $Value)

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $null
    }

    $normalized = $Value.ToLowerInvariant()
    $normalized = [regex]::Replace($normalized, '[^a-z0-9]+', ' ')
    $normalized = [regex]::Replace($normalized, '\s+', ' ').Trim()
    if ([string]::IsNullOrWhiteSpace($normalized)) {
        return $null
    }

    return $normalized
}

function Get-ComparableKeywordTokens {
    param([string] $Value)

    $normalized = Get-NormalizedIdentity -Value $Value
    if ([string]::IsNullOrWhiteSpace($normalized)) {
        return @()
    }

    $tokens = @($normalized.Split(' ', [System.StringSplitOptions]::RemoveEmptyEntries))
    $filtered = New-Object System.Collections.Generic.List[string]
    foreach ($token in $tokens) {
        if ($token.Length -lt 3) {
            continue
        }
        $filtered.Add($token) | Out-Null
    }

    return $filtered.ToArray()
}

function Test-FallbackAuthorMatch {
    param(
        [Parameter(Mandatory = $true)] $SourceItem,
        [Parameter(Mandatory = $true)] $CandidateInfo
    )

    $sourceValues = @(
        [string]$SourceItem.channel,
        [string]$SourceItem.uploader,
        [string]$SourceItem.artist
    )
    $candidateValues = @(
        [string]$CandidateInfo.channel,
        [string]$CandidateInfo.uploader,
        [string]$CandidateInfo.artist
    )

    $sourceSet = New-Object System.Collections.Generic.HashSet[string]
    foreach ($value in $sourceValues) {
        $normalized = Get-NormalizedIdentity -Value $value
        if (-not [string]::IsNullOrWhiteSpace($normalized)) {
            $sourceSet.Add($normalized) | Out-Null
        }
    }

    $candidateSet = New-Object System.Collections.Generic.HashSet[string]
    foreach ($value in $candidateValues) {
        $normalized = Get-NormalizedIdentity -Value $value
        if (-not [string]::IsNullOrWhiteSpace($normalized)) {
            $candidateSet.Add($normalized) | Out-Null
        }
    }

    if ($sourceSet.Count -eq 0 -or $candidateSet.Count -eq 0) {
        return $false
    }

    foreach ($normalized in $sourceSet) {
        if ($candidateSet.Contains($normalized)) {
            return $true
        }
    }

    return $false
}

function Test-FallbackKeywordMatch {
    param(
        [Parameter(Mandatory = $true)] $SourceItem,
        [Parameter(Mandatory = $true)] $CandidateInfo,
        [Parameter(Mandatory = $true)][string] $SourceTitle,
        [Parameter(Mandatory = $true)][string] $CandidateTitle
    )

    $titleTokens = New-Object System.Collections.Generic.HashSet[string]
    foreach ($token in @(Get-ComparableKeywordTokens -Value $SourceTitle)) {
        if (-not [string]::IsNullOrWhiteSpace($token)) {
            $titleTokens.Add($token) | Out-Null
        }
    }

    $candidateTitleTokens = New-Object System.Collections.Generic.HashSet[string]
    foreach ($token in @(Get-ComparableKeywordTokens -Value $CandidateTitle)) {
        if (-not [string]::IsNullOrWhiteSpace($token)) {
            $candidateTitleTokens.Add($token) | Out-Null
        }
    }

    $candidateIdentityTokens = New-Object System.Collections.Generic.HashSet[string]
    foreach ($value in @([string]$CandidateInfo.channel, [string]$CandidateInfo.uploader, [string]$CandidateInfo.artist)) {
        foreach ($token in @(Get-ComparableKeywordTokens -Value $value)) {
            if (-not [string]::IsNullOrWhiteSpace($token)) {
                $candidateIdentityTokens.Add($token) | Out-Null
            }
        }
    }

    $matchedTokens = New-Object System.Collections.Generic.HashSet[string]
    foreach ($token in $titleTokens) {
        if ($candidateTitleTokens.Contains($token) -or $candidateIdentityTokens.Contains($token)) {
            $matchedTokens.Add($token) | Out-Null
        }
    }

    if ($titleTokens.Count -eq 0) {
        return $false
    }

    if ($titleTokens.Count -eq 1) {
        return ($matchedTokens.Count -ge 1)
    }

    return ($matchedTokens.Count -ge 2)
}

function Add-LogRecord {
    param(
        [Parameter(Mandatory = $true)]
        [string] $Section,
        [Parameter(Mandatory = $true)]
        [string] $Message
    )

    $script:LogRecords.Add([pscustomobject]@{
        Section = $Section
        Message = $Message
    }) | Out-Null
}

function Get-PlaylistDownloadBuckets {
    param(
        [Parameter(Mandatory = $true)] $Items,
        [Parameter(Mandatory = $true)][ValidateSet('Audio', 'Video')] [string] $Mode
    )

    $downloadItems = @()
    $reviewItems = @()

    foreach ($item in $Items) {
        if ($Mode -eq 'Audio' -and (Get-MusicalVerdict -Item $item) -eq 'Reject') {
            $reviewItems += $item
            continue
        }

        $downloadItems += $item
    }

    return [pscustomobject]@{
        DownloadItems = @($downloadItems)
        ReviewItems   = @($reviewItems)
    }
}

function Select-SkippedPlaylistItems {
    param(
        [Parameter(Mandatory = $true)] [object[]] $Items,
        [string] $SourceLabel
    )

    if ($Items.Count -eq 0) {
        return @()
    }

    try {
        Add-Type -AssemblyName System.Windows.Forms, System.Drawing -ErrorAction Stop
    } catch {
        Write-Host ''
        Write-Host 'Skipped items were found, but the checkbox window is unavailable.'
        Write-Host 'Enter comma-separated item numbers to keep, or press Enter to skip them all.'
        for ($i = 0; $i -lt $Items.Count; $i++) {
            $item = $Items[$i]
            $prefix = ''
            if ($item.PSObject.Properties.Match('playlist_index').Count -gt 0 -and -not [string]::IsNullOrWhiteSpace([string]$item.playlist_index)) {
                $prefix = ('{0:0000} - ' -f [int]$item.playlist_index)
            }
            Write-Host ("{0}. {1}{2}" -f ($i + 1), $prefix, $item.title)
        }

        $choice = (Read-Host 'Keep which items').Trim()
        if ([string]::IsNullOrWhiteSpace($choice)) {
            return @()
        }

        $selected = New-Object System.Collections.Generic.List[object]
        foreach ($part in ($choice -split ',')) {
            $index = 0
            if ([int]::TryParse($part.Trim(), [ref]$index) -and $index -ge 1 -and $index -le $Items.Count) {
                $selected.Add($Items[$index - 1]) | Out-Null
            }
        }
        return @($selected)
    }

    $form = New-Object System.Windows.Forms.Form
    $form.Text = if ([string]::IsNullOrWhiteSpace($SourceLabel)) { 'Keep skipped items' } else { 'Keep skipped items - ' + $SourceLabel }
    $form.StartPosition = 'CenterScreen'
    $form.Width = 860
    $form.Height = 560
    $form.MinimizeBox = $false
    $form.MaximizeBox = $false
    $form.TopMost = $true

    $layout = New-Object System.Windows.Forms.TableLayoutPanel
    $layout.Dock = 'Fill'
    $layout.ColumnCount = 1
    $layout.RowCount = 3
    $layout.Padding = New-Object System.Windows.Forms.Padding(12)
    $layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null
    $layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
    $layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null
    $form.Controls.Add($layout)

    $label = New-Object System.Windows.Forms.Label
    $label.AutoSize = $true
    $label.Text = 'Check the items you want to keep. Selected items will be downloaded and added back in playlist order.'
    $label.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 8)
    $layout.Controls.Add($label, 0, 0)

    $checkedList = New-Object System.Windows.Forms.CheckedListBox
    $checkedList.Dock = 'Fill'
    $checkedList.CheckOnClick = $true
    $checkedList.HorizontalScrollbar = $true
    $checkedList.IntegralHeight = $false
    $checkedList.FormattingEnabled = $true
    $checkedList.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 8)

    $displayItems = New-Object System.Collections.Generic.List[string]
    foreach ($item in $Items) {
        $prefix = ''
        if ($item.PSObject.Properties.Match('playlist_index').Count -gt 0 -and -not [string]::IsNullOrWhiteSpace([string]$item.playlist_index)) {
            $prefix = ('{0:0000} - ' -f [int]$item.playlist_index)
        }
        $displayItems.Add($prefix + [string]$item.title + '  [not music-like]') | Out-Null
    }

    for ($i = 0; $i -lt $displayItems.Count; $i++) {
        [void]$checkedList.Items.Add($displayItems[$i], $false)
    }
    $layout.Controls.Add($checkedList, 0, 1)

    $buttonRow = New-Object System.Windows.Forms.FlowLayoutPanel
    $buttonRow.Dock = 'Fill'
    $buttonRow.FlowDirection = 'RightToLeft'
    $buttonRow.WrapContents = $false
    $buttonRow.AutoSize = $true
    $layout.Controls.Add($buttonRow, 0, 2)

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = 'Keep Selected'
    $okButton.AutoSize = $true
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $buttonRow.Controls.Add($okButton)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = 'Skip All'
    $cancelButton.AutoSize = $true
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $buttonRow.Controls.Add($cancelButton)

    $selectNoneButton = New-Object System.Windows.Forms.Button
    $selectNoneButton.Text = 'Select None'
    $selectNoneButton.AutoSize = $true
    $buttonRow.Controls.Add($selectNoneButton)

    $selectAllButton = New-Object System.Windows.Forms.Button
    $selectAllButton.Text = 'Select All'
    $selectAllButton.AutoSize = $true
    $buttonRow.Controls.Add($selectAllButton)

    $keepAllButton = New-Object System.Windows.Forms.Button
    $keepAllButton.Text = 'Keep All Skipped Items'
    $keepAllButton.AutoSize = $true
    $buttonRow.Controls.Add($keepAllButton)

    $selectAllButton.Add_Click({
        for ($i = 0; $i -lt $checkedList.Items.Count; $i++) {
            $checkedList.SetItemChecked($i, $true)
        }
    })

    $keepAllButton.Add_Click({
        for ($i = 0; $i -lt $checkedList.Items.Count; $i++) {
            $checkedList.SetItemChecked($i, $true)
        }
    })

    $selectNoneButton.Add_Click({
        for ($i = 0; $i -lt $checkedList.Items.Count; $i++) {
            $checkedList.SetItemChecked($i, $false)
        }
    })

    $form.AcceptButton = $okButton
    $form.CancelButton = $cancelButton

    $result = $form.ShowDialog()
    if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
        return @()
    }

    $selected = New-Object System.Collections.Generic.List[object]
    for ($i = 0; $i -lt $checkedList.CheckedIndices.Count; $i++) {
        $index = [int]$checkedList.CheckedIndices[$i]
        if ($index -ge 0 -and $index -lt $Items.Count) {
            $selected.Add($Items[$index]) | Out-Null
        }
    }

    return @($selected)
}

function Get-StagedMediaMap {
    param([Parameter(Mandatory = $true)][string] $StageRoot)

    $map = @{}
    foreach ($stageFile in Get-ChildItem -LiteralPath $StageRoot -Recurse -File -ErrorAction SilentlyContinue) {
        if ($stageFile.Extension -eq '.json' -or $stageFile.Extension -eq '.id') {
            continue
        }
        if (-not $map.ContainsKey($stageFile.BaseName)) {
            $map[$stageFile.BaseName] = $stageFile.FullName
        }
    }

    return $map
}

function Get-ItemDisplayLabel {
    param([Parameter(Mandatory = $true)] $Item)

    $parts = New-Object System.Collections.Generic.List[string]
    if ($Item.PSObject.Properties.Match('playlist_index').Count -gt 0) {
        try {
            $playlistIndex = [int]$Item.playlist_index
            if ($playlistIndex -gt 0) {
                $parts.Add(('{0:0000}' -f $playlistIndex)) | Out-Null
            }
        } catch {
        }
    }

    $title = [string]$Item.title
    if ([string]::IsNullOrWhiteSpace($title)) {
        $title = 'untitled'
    }

    if ($parts.Count -gt 0) {
        return (($parts -join ' - ') + ' - ' + $title)
    }

    return $title
}

function Select-UpdateTargetFolder {
    param([Parameter(Mandatory = $true)][string] $DownloadRoot)

    Write-Host '1. Pick from downloaded folders'
    Write-Host '2. Pick a custom folder'

    $choice = ''
    while ($choice -notin @('1', '2')) {
        $choice = (Read-Host 'Select target folder source').Trim()
    }

    if ($choice -eq '1') {
        $folders = Get-ChildItem -LiteralPath $DownloadRoot -Directory -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -notlike '__*' } |
            Sort-Object Name

        if ($folders.Count -eq 0) {
            throw 'No playlist folders were found under All downloaded playlist.'
        }

        for ($i = 0; $i -lt $folders.Count; $i++) {
            Write-Host ("{0}. {1}" -f ($i + 1), $folders[$i].Name)
        }

        $selection = 0
        while ($selection -lt 1 -or $selection -gt $folders.Count) {
            [void][int]::TryParse((Read-Host 'Pick target folder'), [ref]$selection)
        }

        return [pscustomobject]@{
            Path  = $folders[$selection - 1].FullName
            Label = $folders[$selection - 1].Name
        }
    }

    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
    } catch {
        Write-Host 'Windows folder picker is unavailable. Enter a folder path instead.'
        $manualPath = (Read-Host 'Enter folder path').Trim()
        if ([string]::IsNullOrWhiteSpace($manualPath) -or -not (Test-Path -LiteralPath $manualPath)) {
            return $null
        }

        return [pscustomobject]@{
            Path  = (Resolve-Path -LiteralPath $manualPath).Path
            Label = Split-Path -Leaf $manualPath
        }
    }

    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = 'Select the folder to scan and update'
    $dialog.ShowNewFolderButton = $false
    $dialog.SelectedPath = $DownloadRoot

    try {
        $result = $dialog.ShowDialog()
    } catch {
        Write-Host ('Folder picker failed: ' + $_.Exception.Message)
        return $null
    }

    if ($result -ne [System.Windows.Forms.DialogResult]::OK -or [string]::IsNullOrWhiteSpace($dialog.SelectedPath)) {
        return $null
    }

    return [pscustomobject]@{
        Path  = $dialog.SelectedPath
        Label = Split-Path -Leaf $dialog.SelectedPath
    }
}

function Merge-LogFiles {
    param(
        [string[]] $LogFiles,
        [string] $FinalLogPath
    )

    $all = New-Object System.Collections.Generic.List[object]
    foreach ($path in $LogFiles) {
        if ([string]::IsNullOrWhiteSpace($path) -or -not (Test-Path -LiteralPath $path)) {
            continue
        }
        try {
            $records = Get-Content -LiteralPath $path -Raw | ConvertFrom-Json
            if ($records -is [System.Array]) {
                foreach ($record in $records) { $all.Add($record) | Out-Null }
            } elseif ($records) {
                $all.Add($records) | Out-Null
            }
        } catch {
            continue
        }
    }

    foreach ($record in $script:LogRecords) {
        $all.Add($record) | Out-Null
    }

    $lines = New-Object System.Collections.Generic.List[string]
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $lines.Add("Run: $timestamp") | Out-Null
    $lines.Add('') | Out-Null

    if ($all.Count -gt 0) {
        foreach ($section in @('Fallback Success', 'Skipped By Filter', 'Kept After Review', 'Failed After All Attempts')) {
            $sectionRecords = @($all | Where-Object { $_.Section -eq $section })
            if ($sectionRecords.Count -eq 0) {
                continue
            }
            $lines.Add($section) | Out-Null
            foreach ($record in $sectionRecords) {
                $lines.Add(('- ' + $record.Message)) | Out-Null
            }
            $lines.Add('') | Out-Null
        }
    }

    $finalDir = Split-Path -Parent $FinalLogPath
    if (-not (Test-Path -LiteralPath $finalDir)) {
        New-Item -ItemType Directory -Path $finalDir | Out-Null
    }

    Set-Content -LiteralPath $FinalLogPath -Value $lines -Encoding UTF8
}

function Get-SearchUrlsFromMusic {
    param([Parameter(Mandatory = $true)][string] $Query)

    $encoded = [uri]::EscapeDataString($Query)
    $url = "https://music.youtube.com/search?q=$encoded&params=$musicSearchSongsParams&hl=en&gl=US"
    $headers = @{
        'User-Agent'      = $browserUserAgent
        'Accept-Language' = 'en-US,en;q=0.9'
    }

    try {
        $resp = Invoke-WebRequest -UseBasicParsing -Uri $url -Headers $headers -TimeoutSec 25
    } catch {
        return @()
    }

    $urls = New-Object System.Collections.Generic.List[string]
    $regexMatches = [regex]::Matches($resp.Content, 'videoId\\x22:\\x22(?<id>[A-Za-z0-9_-]{11})')
    foreach ($match in $regexMatches) {
        $id = $match.Groups['id'].Value
        if (-not [string]::IsNullOrWhiteSpace($id)) {
            $url = "https://music.youtube.com/watch?v=$id"
            if (-not $urls.Contains($url)) {
                $urls.Add($url) | Out-Null
            }
        }
    }

    return $urls.ToArray()
}

function Get-SearchUrlsFromYouTube {
    param(
        [Parameter(Mandatory = $true)][string] $Query,
        [string] $TitleHint,
        [string] $UploaderHint
    )

    $result = Invoke-YtDlp -Arguments @(
        '--quiet',
        '--no-warnings',
        '--flat-playlist',
        '--skip-download',
        '--print', "%(id)s`t%(title)s`t%(uploader)s",
        ('ytsearch25:' + $Query)
    ) -CaptureOutput

    if ($result.ExitCode -ne 0) {
        return @()
    }

    $queryTitle = ''
    if ($TitleHint) {
        $queryTitle = $TitleHint.Trim().ToLowerInvariant()
    }
    $queryUploader = ''
    if ($UploaderHint) {
        $queryUploader = $UploaderHint.Trim().ToLowerInvariant()
    }

    $exactMatches = New-Object System.Collections.Generic.List[string]
    $nearMatches = New-Object System.Collections.Generic.List[string]
    $fallbackMatches = New-Object System.Collections.Generic.List[string]
    foreach ($line in $result.Output) {
        if ([string]::IsNullOrWhiteSpace($line)) {
            continue
        }
        $parts = $line -split "`t", 3
        if ($parts.Count -lt 1) {
            continue
        }

        $id = $parts[0].Trim()
        if ([string]::IsNullOrWhiteSpace($id)) {
            continue
        }

        $title = if ($parts.Count -ge 2) { $parts[1] } else { '' }
        $uploader = if ($parts.Count -ge 3) { $parts[2] } else { '' }
        $titleLower = $title.ToLowerInvariant()
        $uploaderLower = $uploader.ToLowerInvariant()
        $candidate = "https://www.youtube.com/watch?v=$id"

        if (-not [string]::IsNullOrWhiteSpace($queryTitle) -and $titleLower -notlike "*$queryTitle*") {
            if (-not $fallbackMatches.Contains($candidate)) {
                $fallbackMatches.Add($candidate) | Out-Null
            }
            continue
        }

        if (-not [string]::IsNullOrWhiteSpace($queryUploader) -and $uploaderLower -notlike "*$queryUploader*") {
            if (-not $nearMatches.Contains($candidate)) {
                $nearMatches.Add($candidate) | Out-Null
            }
            continue
        }

        if (-not $exactMatches.Contains($candidate)) {
            $exactMatches.Add($candidate) | Out-Null
        }
    }

    $ordered = New-Object System.Collections.Generic.List[string]
    foreach ($group in @($exactMatches, $nearMatches, $fallbackMatches)) {
        foreach ($candidate in $group) {
            if (-not $ordered.Contains($candidate)) {
                $ordered.Add($candidate) | Out-Null
            }
        }
    }

    return $ordered.ToArray()
}

function Get-ExistingCatalog {
    param([Parameter(Mandatory = $true)][string] $Folder)

    $catalog = New-Object System.Collections.Generic.List[object]
    $mediaFiles = Get-ChildItem -LiteralPath $Folder -File -ErrorAction SilentlyContinue |
        Where-Object {
            $_.Extension -ne '.json' -and $_.Extension -ne '.id'
        }

    foreach ($media in $mediaFiles) {
        $infoPath = [System.IO.Path]::ChangeExtension($media.FullName, '.info.json')
        $info = $null
        if (Test-Path -LiteralPath $infoPath) {
            try {
                $info = Get-Content -LiteralPath $infoPath -Raw | ConvertFrom-Json
            } catch {
                $info = $null
            }
        }

        $sidecarId = $null
        $idPath = Get-VideoIdSidecarPath -MediaPath $media.FullName
        if (Test-Path -LiteralPath $idPath) {
            try {
                $sidecarId = (Get-Content -LiteralPath $idPath -Raw).Trim()
            } catch {
                $sidecarId = $null
            }
        }

        $id = $null
        if ($info -and -not [string]::IsNullOrWhiteSpace([string]$info.id)) {
            $id = [string]$info.id
        } elseif (-not [string]::IsNullOrWhiteSpace($sidecarId)) {
            $id = $sidecarId
        }

        if ([string]::IsNullOrWhiteSpace($id)) {
            continue
        }

        $index = $null
        if ($media.BaseName -match '^(\d{4})\s*-\s*') {
            $index = [int]$matches[1]
        } elseif ($info -and $info.playlist_index) {
            try {
                $index = [int]$info.playlist_index
            } catch {
                $index = $null
            }
        }

        $title = if ($info -and -not [string]::IsNullOrWhiteSpace([string]$info.title)) {
            [string]$info.title
        } else {
            $media.BaseName
        }

        $catalog.Add([pscustomobject]@{
            Id         = $id
            Title      = $title
            Index      = $index
            MediaPath  = $media.FullName
            InfoPath   = if (Test-Path -LiteralPath $infoPath) { $infoPath } else { $null }
            IdPath     = if (Test-Path -LiteralPath $idPath) { $idPath } else { $null }
            BaseName   = $media.BaseName
            Extension  = $media.Extension
            IsNumbered = ($media.BaseName -match '^\d{4}\s*-\s*')
        }) | Out-Null
    }

    return $catalog.ToArray()
}

function Get-ExistingNumberedFiles {
    param([Parameter(Mandatory = $true)][string] $Folder)

    Get-ChildItem -LiteralPath $Folder -File -ErrorAction SilentlyContinue |
        Where-Object { $_.BaseName -match '^\d{4}\s*-\s*' -and $_.Extension -ne '.json' } |
        Sort-Object { [int]($_.BaseName.Substring(0, 4)) }, BaseName
}

function Select-DownloadAttemptSequence {
    param(
        [Parameter(Mandatory = $true)] $Item,
        [Parameter(Mandatory = $true)][ValidateSet('Audio', 'Video')] [string] $Mode,
        [switch] $BypassGuard
    )

    $watchUrl = [string]$Item.webpage_url
    if ([string]::IsNullOrWhiteSpace($watchUrl)) {
        $watchUrl = [string]$Item.url
    }
    if ([string]::IsNullOrWhiteSpace($watchUrl)) {
        $watchUrl = 'https://www.youtube.com/watch?v=' + $Item.id
    }
    $musicWatchUrl = 'https://music.youtube.com/watch?v=' + $Item.id
    $title = [string]$Item.title
    $uploader = [string]$Item.channel
    if ([string]::IsNullOrWhiteSpace($uploader)) {
        $uploader = [string]$Item.uploader
    }
    $track = [string]$Item.track
    $artist = [string]$Item.artist
    $attempts = New-Object System.Collections.Generic.List[object]

    if ($Mode -eq 'Audio') {
        $musicQueries = New-Object System.Collections.Generic.List[string]
        foreach ($query in @(
            @($title, $uploader) -join ' ',
            @($title, $artist) -join ' ',
            @($track, $artist) -join ' ',
            @($title, 'Topic') -join ' ',
            @($title, $uploader, 'Topic') -join ' ',
            $title
        )) {
            if (-not [string]::IsNullOrWhiteSpace($query) -and -not $musicQueries.Contains($query)) {
                $musicQueries.Add($query) | Out-Null
            }
        }

        $youtubeQueries = New-Object System.Collections.Generic.List[string]
        foreach ($query in @(
            @($title, $uploader) -join ' ',
            @($title, $artist) -join ' ',
            @($track, $artist) -join ' ',
            @($title, 'Topic') -join ' ',
            @($title, $uploader, 'Topic') -join ' ',
            $title
        )) {
            if (-not [string]::IsNullOrWhiteSpace($query) -and -not $youtubeQueries.Contains($query)) {
                $youtubeQueries.Add($query) | Out-Null
            }
        }

        if ($BypassGuard) {
            $attempts.Add([pscustomobject]@{ Kind = 'direct-youtube'; Url = $watchUrl; Format = 'audio'; Label = 'direct YouTube URL' }) | Out-Null

            foreach ($query in $youtubeQueries) {
                $attempts.Add([pscustomobject]@{ Kind = 'youtube-search'; Query = $query; Format = 'audio'; Label = "YouTube search: $query" }) | Out-Null
            }

            $attempts.Add([pscustomobject]@{ Kind = 'direct-music'; Url = $musicWatchUrl; Format = 'audio'; Label = 'direct YouTube Music URL' }) | Out-Null

            foreach ($query in $musicQueries) {
                $attempts.Add([pscustomobject]@{ Kind = 'music-search'; Query = $query; Format = 'audio'; Label = "YouTube Music search: $query" }) | Out-Null
            }
        } else {
            $attempts.Add([pscustomobject]@{ Kind = 'direct-music'; Url = $musicWatchUrl; Format = 'audio'; Label = 'direct YouTube Music URL' }) | Out-Null

            foreach ($query in $musicQueries) {
                $attempts.Add([pscustomobject]@{ Kind = 'music-search'; Query = $query; Format = 'audio'; Label = "YouTube Music search: $query" }) | Out-Null
            }

            $attempts.Add([pscustomobject]@{ Kind = 'direct-youtube'; Url = $watchUrl; Format = 'audio'; Label = 'direct YouTube URL' }) | Out-Null

            foreach ($query in $youtubeQueries) {
                $attempts.Add([pscustomobject]@{ Kind = 'youtube-search'; Query = $query; Format = 'audio'; Label = "YouTube search: $query" }) | Out-Null
            }
        }
    } else {
        $attempts.Add([pscustomobject]@{ Kind = 'direct-youtube'; Url = $watchUrl; Format = 'video'; Label = 'direct YouTube URL' }) | Out-Null

        $youtubeQueries = New-Object System.Collections.Generic.List[string]
        foreach ($query in @(
            @($title, $uploader) -join ' ',
            @($title, $artist) -join ' ',
            @($track, $artist) -join ' ',
            @($title, 'Topic') -join ' ',
            @($title, $uploader, 'Topic') -join ' ',
            $title
        )) {
            if (-not [string]::IsNullOrWhiteSpace($query) -and -not $youtubeQueries.Contains($query)) {
                $youtubeQueries.Add($query) | Out-Null
            }
        }

        foreach ($query in $youtubeQueries) {
            $attempts.Add([pscustomobject]@{ Kind = 'youtube-search'; Query = $query; Format = 'video'; Label = "YouTube search: $query" }) | Out-Null
        }

        $attempts.Add([pscustomobject]@{ Kind = 'direct-youtube'; Url = $watchUrl; Format = 'audio'; Label = 'audio fallback from YouTube URL' }) | Out-Null

        foreach ($query in $youtubeQueries) {
            $attempts.Add([pscustomobject]@{ Kind = 'youtube-search'; Query = $query; Format = 'audio'; Label = "audio fallback from YouTube search: $query" }) | Out-Null
        }
    }

    return $attempts.ToArray()
}

function Invoke-DownloadAttempt {
    param(
        [Parameter(Mandatory = $true)] $Item,
        [Parameter(Mandatory = $true)][ValidateSet('Audio', 'Video')] [string] $Mode,
        [Parameter(Mandatory = $true)][string] $StageFolder,
        [switch] $BypassGuard
    )

    if (Test-UnavailableVideoPlaceholder -Item $Item) {
        return [pscustomobject]@{
            Success      = $false
            FilteredOut  = $true
            UsedFallback = $false
            AttemptLabel = 'private or unavailable video'
            MediaPath    = $null
            InfoPath     = $null
            IdPath       = $null
            ItemId       = $null
            Url          = $null
            AttemptIndex = -1
        }
    }

    $attempts = Select-DownloadAttemptSequence -Item $Item -Mode $Mode -BypassGuard:$BypassGuard
    $itemId = [string]$Item.id
    $title = [string]$Item.title
    if ([string]::IsNullOrWhiteSpace($title)) {
        $title = 'untitled'
    }
    $allowPersonaOverride = Test-GuardTitleOverride -Title $title

    for ($i = 0; $i -lt $attempts.Count; $i++) {
        $attempt = $attempts[$i]
        $resolvedUrls = @()
        switch ($attempt.Kind) {
            'direct-music' { $resolvedUrls = @($attempt.Url) }
            'direct-youtube' { $resolvedUrls = @($attempt.Url) }
            'music-search' { $resolvedUrls = @(Get-SearchUrlsFromMusic -Query $attempt.Query) }
            'youtube-search' {
                $hintUploader = [string]$Item.channel
                if ([string]::IsNullOrWhiteSpace($hintUploader)) {
                    $hintUploader = [string]$Item.uploader
                }
                $resolvedUrls = @(Get-SearchUrlsFromYouTube -Query $attempt.Query -TitleHint $title -UploaderHint $hintUploader)
            }
        }

        if ($resolvedUrls.Count -eq 0) {
            continue
        }

        foreach ($resolvedUrl in $resolvedUrls) {
            if ([string]::IsNullOrWhiteSpace($resolvedUrl)) {
                continue
            }

            $attemptStage = Join-Path $StageFolder ([Guid]::NewGuid().ToString('N'))
            New-Item -ItemType Directory -Path $attemptStage | Out-Null

            $formatArgs = @()
            if ($attempt.Format -eq 'audio') {
                $formatArgs = @(
                    '-f', 'bestaudio[acodec=opus]/bestaudio/best',
                    '-x',
                    '--audio-format', 'opus',
                    '--embed-thumbnail',
                    '--embed-metadata',
                    '--sponsorblock-remove', 'music_offtopic'
                )
            } else {
                $formatArgs = @(
                    '-f', 'bestvideo*+bestaudio/best'
                )
            }

            $outTemplate = Join-Path $attemptStage '%(id)s.%(ext)s'
            $downloadArgs = @(
                '--newline',
                '--progress',
                '--console-title',
                '--progress-template', 'download:[download] %(progress._percent_str)s of %(progress._downloaded_bytes_str)s / %(progress._total_bytes_str)s at %(progress._speed_str)s ETA %(progress._eta_str)s',
                '--progress-template', 'postprocess:[postprocess] %(progress.postprocessor)s %(progress._percent_str)s',
                '--write-info-json',
                '--force-overwrites',
                '--no-mtime',
                '--no-playlist',
                '-o', $outTemplate
            )
            $downloadArgs += $formatArgs
            $downloadArgs += $resolvedUrl
            try {
                Write-Host ("Attempt {0}/{1}: {2}" -f ($i + 1), $attempts.Count, $attempt.Label)
                $result = Invoke-YtDlp -Arguments $downloadArgs
            } catch {
                if (Test-Path -LiteralPath $attemptStage) {
                    Remove-Item -LiteralPath $attemptStage -Recurse -Force -ErrorAction SilentlyContinue
                }
                continue
            }

            if ($result.ExitCode -ne 0) {
                if (Test-Path -LiteralPath $attemptStage) {
                    Remove-Item -LiteralPath $attemptStage -Recurse -Force -ErrorAction SilentlyContinue
                }
                continue
            }

            $media = Get-ChildItem -LiteralPath $attemptStage -File -ErrorAction SilentlyContinue |
                Where-Object { $_.Extension -ne '.json' -and $_.Extension -ne '.id' } |
                Select-Object -First 1

            if (-not $media) {
                if (Test-Path -LiteralPath $attemptStage) {
                    Remove-Item -LiteralPath $attemptStage -Recurse -Force -ErrorAction SilentlyContinue
                }
                continue
            }

            $infoPath = [System.IO.Path]::ChangeExtension($media.FullName, '.info.json')
            $candidateInfo = $null
            if ($attempt.Format -eq 'audio' -and -not $BypassGuard) {
                $isMusicLike = Test-DownloadedMusicMetadata -InfoPath $infoPath
                if (-not $isMusicLike) {
                    if (Test-Path -LiteralPath $attemptStage) {
                        Remove-Item -LiteralPath $attemptStage -Recurse -Force -ErrorAction SilentlyContinue
                    }

                    if ($attempt.Kind -eq 'direct-music') {
                        if ($allowPersonaOverride) {
                            continue
                        }

                        return [pscustomobject]@{
                            Success      = $false
                            FilteredOut  = $true
                            UsedFallback = $false
                            AttemptLabel = $attempt.Label
                            MediaPath    = $null
                            InfoPath     = $null
                            IdPath       = $null
                            ItemId       = $null
                            Url          = $resolvedUrl
                            AttemptIndex = $i
                        }
                    }

                    continue
                }

                if ($i -gt 0) {
                    try {
                        $candidateInfo = Get-Content -LiteralPath $infoPath -Raw | ConvertFrom-Json
                    } catch {
                        $candidateInfo = $null
                    }

                    $candidateTitle = if ($candidateInfo) { [string]$candidateInfo.title } else { $null }
                    $acceptCandidate = $false
                    if ($candidateInfo) {
                        $titleMatch = Test-FallbackTitleMatch -SourceTitle $title -CandidateTitle $candidateTitle
                        $authorMatch = Test-FallbackAuthorMatch -SourceItem $Item -CandidateInfo $candidateInfo
                        $keywordMatch = Test-FallbackKeywordMatch -SourceItem $Item -CandidateInfo $candidateInfo -SourceTitle $title -CandidateTitle $candidateTitle
                        $acceptCandidate = (($titleMatch -and $authorMatch) -or $keywordMatch)
                    }

                    if (-not $acceptCandidate) {
                        if (Test-Path -LiteralPath $attemptStage) {
                            Remove-Item -LiteralPath $attemptStage -Recurse -Force -ErrorAction SilentlyContinue
                        }
                        continue
                    }
                }
            }

            $actualId = $itemId
            if (Test-Path -LiteralPath $infoPath) {
                try {
                    $info = Get-Content -LiteralPath $infoPath -Raw | ConvertFrom-Json
                    if ($info -and -not [string]::IsNullOrWhiteSpace([string]$info.id)) {
                        $actualId = [string]$info.id
                    }
                } catch {
                    $actualId = $itemId
                }
            } else {
                $actualId = $media.BaseName
            }

            $finalMedia = Join-Path $attemptStage ($itemId + $media.Extension)
            if ($media.FullName -ne $finalMedia) {
                Move-Item -LiteralPath $media.FullName -Destination $finalMedia
            }

            $finalInfo = $null
            if (Test-Path -LiteralPath $infoPath) {
                $finalInfo = Join-Path $attemptStage ($itemId + '.info.json')
                if ($infoPath -ne $finalInfo) {
                    Move-Item -LiteralPath $infoPath -Destination $finalInfo
                }
            }

            $sidecarPath = Join-Path $attemptStage ($itemId + '.yt-dlp.id')
            Set-Content -LiteralPath $sidecarPath -Value $actualId -Encoding UTF8

            return [pscustomobject]@{
                Success      = $true
                FilteredOut  = $false
                UsedFallback = ($i -gt 0)
                AttemptLabel = $attempt.Label
                MediaPath    = $finalMedia
                InfoPath     = $finalInfo
                IdPath       = $sidecarPath
                ItemId       = $actualId
                Url          = $resolvedUrl
                AttemptIndex = $i
            }
        }
    }

    return [pscustomobject]@{
        Success      = $false
        FilteredOut  = $false
        UsedFallback = $false
        AttemptLabel = $null
        MediaPath    = $null
        InfoPath     = $null
        IdPath       = $null
        ItemId       = $null
        Url          = $null
        AttemptIndex = -1
    }
}

function Move-FilePairToStage {
    param(
        [Parameter(Mandatory = $true)][string] $SourceMedia,
        [string] $SourceInfo,
        [Parameter(Mandatory = $true)][string] $StageFolder,
        [Parameter(Mandatory = $true)][string] $TempBase
    )

    $mediaExt = [System.IO.Path]::GetExtension($SourceMedia)
    $stagedMedia = Join-Path $StageFolder ($TempBase + $mediaExt)
    Move-Item -LiteralPath $SourceMedia -Destination $stagedMedia

    $stagedInfo = $null
    if (-not [string]::IsNullOrWhiteSpace($SourceInfo) -and (Test-Path -LiteralPath $SourceInfo)) {
        $stagedInfo = Join-Path $StageFolder ($TempBase + '.info.json')
        Move-Item -LiteralPath $SourceInfo -Destination $stagedInfo
    }

    $sourceIdSidecar = Get-VideoIdSidecarPath -MediaPath $SourceMedia
    $stagedIdSidecar = $null
    if (Test-Path -LiteralPath $sourceIdSidecar) {
        $stagedIdSidecar = Join-Path $StageFolder ($TempBase + '.yt-dlp.id')
        Move-Item -LiteralPath $sourceIdSidecar -Destination $stagedIdSidecar
    }

    return [pscustomobject]@{
        MediaPath = $stagedMedia
        InfoPath  = $stagedInfo
        IdPath    = $stagedIdSidecar
    }
}

function Move-StagedPairToFinal {
    param(
        [Parameter(Mandatory = $true)][string] $SourceMedia,
        [string] $SourceInfo,
        [string] $SourceId,
        [Parameter(Mandatory = $true)][string] $FinalMediaBase,
        [Parameter(Mandatory = $true)][string] $TargetFolder,
        [switch] $AllowUniqueSuffix,
        [string] $ItemId
    )

    $mediaExt = [System.IO.Path]::GetExtension($SourceMedia)
    $finalMedia = Join-Path $TargetFolder ($FinalMediaBase + $mediaExt)
    if ($AllowUniqueSuffix) {
        $finalMedia = Get-UniqueFilePath -BasePath $finalMedia
    }
    Move-Item -LiteralPath $SourceMedia -Destination $finalMedia

    if (-not [string]::IsNullOrWhiteSpace($SourceInfo) -and (Test-Path -LiteralPath $SourceInfo)) {
        $finalInfo = [System.IO.Path]::ChangeExtension($finalMedia, '.info.json')
        Move-Item -LiteralPath $SourceInfo -Destination $finalInfo
    }

    if (-not [string]::IsNullOrWhiteSpace($SourceId) -and (Test-Path -LiteralPath $SourceId)) {
        $finalId = [System.IO.Path]::ChangeExtension($finalMedia, '.yt-dlp.id')
        Move-Item -LiteralPath $SourceId -Destination $finalId
    } elseif (-not [string]::IsNullOrWhiteSpace($ItemId)) {
        Write-VideoIdSidecar -MediaPath $finalMedia -VideoId $ItemId
    }

    return $finalMedia
}

function Start-WorkerProcess {
    param(
        [Parameter(Mandatory = $true)][string] $JobFile,
        [Parameter(Mandatory = $true)][string] $TempRoot,
        [Parameter(Mandatory = $true)][string] $TargetFolder,
        [Parameter(Mandatory = $true)][string] $WorkerLogFile,
        [Parameter(Mandatory = $true)][ValidateSet('Create', 'Update')] [string] $RunKind,
        [Parameter(Mandatory = $true)][ValidateSet('Audio', 'Video')] [string] $Mode,
        [Parameter(Mandatory = $true)][string] $SourceLabel,
        [string] $WorkerLabel
    )

    $argumentList = @(
        '-NoProfile'
        '-ExecutionPolicy Bypass'
        ('-File "' + $script:LauncherScriptPath + '"')
        '-Worker'
        ('-JobFile "' + $JobFile + '"')
        ('-TempRoot "' + $TempRoot + '"')
        ('-TargetFolder "' + $TargetFolder + '"')
        ('-WorkerLogFile "' + $WorkerLogFile + '"')
        ('-RunKind ' + $RunKind)
        ('-Mode ' + $Mode)
        ('-SourceLabel "' + $SourceLabel + '"')
        ('-WorkerLabel "' + $WorkerLabel + '"')
    ) -join ' '

    $workerCommand = @(
        'start'
        '""'
        '/wait'
        'powershell.exe'
        $argumentList
    ) -join ' '

    return Start-Process -FilePath 'cmd.exe' -ArgumentList ('/c ' + $workerCommand) -WorkingDirectory $PSScriptRoot -PassThru -WindowStyle Normal
}

function Remove-StaleNumberedFiles {
    param([Parameter(Mandatory = $true)][string] $Folder)

    $staleFiles = Get-ChildItem -LiteralPath $Folder -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match '^\d{4}\s*-\s*' }

    foreach ($file in $staleFiles) {
        Remove-Item -LiteralPath $file.FullName -Force -ErrorAction SilentlyContinue
    }
}

function Initialize-StageRoot {
    param([Parameter(Mandatory = $true)][string] $StageRoot)

    if (Test-Path -LiteralPath $StageRoot) {
        Remove-Item -LiteralPath $StageRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
    New-Item -ItemType Directory -Path $StageRoot | Out-Null
}

function Wait-WorkerProcess {
    param([Parameter(Mandatory = $true)] $Process)

    if ($null -eq $Process) {
        return
    }

    try {
        $null = $Process.WaitForExit()
    } catch {
        return
    }
}

function Invoke-WorkerMode {
    Assert-Workspace

    $windowTitle = 'yt-dlp worker'
    if (-not [string]::IsNullOrWhiteSpace($WorkerLabel)) {
        $windowTitle = $WorkerLabel
    }
    $details = New-Object System.Collections.Generic.List[string]
    if (-not [string]::IsNullOrWhiteSpace($RunKind)) { $details.Add($RunKind) | Out-Null }
    if (-not [string]::IsNullOrWhiteSpace($Mode)) { $details.Add($Mode) | Out-Null }
    if (-not [string]::IsNullOrWhiteSpace($SourceLabel)) { $details.Add($SourceLabel) | Out-Null }
    if ($details.Count -gt 0) {
        $windowTitle = $windowTitle + ' | ' + ($details -join ' | ')
    }
    Set-ProcessTitle -Title $windowTitle
    Write-Host ("[{0}] {1}" -f $RunKind, $SourceLabel)

    if (-not (Test-Path -LiteralPath $JobFile)) {
        throw "Worker job file not found: $JobFile"
    }
    if ([string]::IsNullOrWhiteSpace($TempRoot)) {
        throw 'Worker temp root was not provided.'
    }
    if ([string]::IsNullOrWhiteSpace($TargetFolder)) {
        throw 'Worker target folder was not provided.'
    }

    $job = Get-Content -LiteralPath $JobFile -Raw | ConvertFrom-Json
    if ($null -eq $job) {
        Set-Content -LiteralPath $WorkerLogFile -Value '[]' -Encoding UTF8
        return
    }

    if (-not (Test-Path -LiteralPath $TempRoot)) {
        New-Item -ItemType Directory -Path $TempRoot | Out-Null
    }

    $workerStage = Join-Path $TempRoot ('worker-' + [Guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $workerStage | Out-Null

    $workerRecords = New-Object System.Collections.Generic.List[object]
    foreach ($item in @($job)) {
        try {
            $displayLabel = Get-ItemDisplayLabel -Item $item
            Set-ProcessTitle -Title ("{0} | {1}" -f $windowTitle, $displayLabel)
            Write-Host ("[{0}] Downloading: {1}" -f $RunKind, $displayLabel)
            $download = Invoke-DownloadAttempt -Item $item -Mode $Mode -StageFolder $workerStage
            if ($download.Success) {
                if ($download.UsedFallback) {
                    Write-Host ("[{0}] Finished via fallback: {1}" -f $RunKind, $displayLabel)
                    $workerRecords.Add([pscustomobject]@{
                        Section = 'Fallback Success'
                        Message = ("{0} -> {1}" -f $item.title, $download.AttemptLabel)
                    }) | Out-Null
                } else {
                    Write-Host ("[{0}] Finished: {1}" -f $RunKind, $displayLabel)
                }
                continue
            }

            if ($download.FilteredOut) {
                Write-Host ("[{0}] Skipped by filter: {1}" -f $RunKind, $displayLabel)
                $workerRecords.Add([pscustomobject]@{
                    Section = 'Skipped By Filter'
                    Message = ($item.title + ' - direct YouTube Music download was not music-like')
                }) | Out-Null
                continue
            }

            Write-Host ("[{0}] Failed: {1}" -f $RunKind, $displayLabel)
            $workerRecords.Add([pscustomobject]@{
                Section = 'Failed After All Attempts'
                Message = ($item.title + ' (id: ' + $item.id + ')')
            }) | Out-Null
        } catch {
            $displayLabel = Get-ItemDisplayLabel -Item $item
            Write-Host ("[{0}] Failed: {1}" -f $RunKind, $displayLabel)
            $workerRecords.Add([pscustomobject]@{
                Section = 'Failed After All Attempts'
                Message = ($item.title + ' (id: ' + $item.id + ') - ' + $_.Exception.Message)
            }) | Out-Null
        }
    }

    if ($workerRecords.Count -gt 0) {
        Set-Content -LiteralPath $WorkerLogFile -Value ($workerRecords | ConvertTo-Json -Depth 6) -Encoding UTF8
    } else {
        Set-Content -LiteralPath $WorkerLogFile -Value '[]' -Encoding UTF8
    }
}

function Invoke-CreateSingle {
    param(
        [Parameter(Mandatory = $true)] $Item,
        [Parameter(Mandatory = $true)][string] $TargetFolder,
        [Parameter(Mandatory = $true)][ValidateSet('Audio', 'Video')] [string] $Mode
    )

    if ($Mode -eq 'Audio' -and (Get-MusicalVerdict -Item $Item) -eq 'Reject') {
        Add-LogRecord -Section 'Skipped By Filter' -Message ($Item.title + ' - not music-like')
        return
    }

    $stageRoot = Join-Path $TargetFolder '__yt-dlp_stage'
    Initialize-StageRoot -StageRoot $stageRoot
    $stage = Join-Path $stageRoot ([Guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $stage | Out-Null

    try {
        $download = Invoke-DownloadAttempt -Item $Item -Mode $Mode -StageFolder $stage
        if (-not $download.Success) {
            if ($download.FilteredOut) {
                Add-LogRecord -Section 'Skipped By Filter' -Message ($Item.title + ' - direct YouTube Music download was not music-like')
                return
            }
            Add-LogRecord -Section 'Failed After All Attempts' -Message ($Item.title + ' (id: ' + $Item.id + ')')
            return
        }

        $safeTitle = Get-SafeFileName -Name $Item.title
        $finalPath = Move-StagedPairToFinal -SourceMedia $download.MediaPath -SourceInfo $download.InfoPath -SourceId $download.IdPath -FinalMediaBase $safeTitle -TargetFolder $TargetFolder -AllowUniqueSuffix -ItemId $download.ItemId
        if ($download.UsedFallback) {
            Add-LogRecord -Section 'Fallback Success' -Message ($Item.title + ' -> ' + (Split-Path -Leaf $finalPath))
        }
    } finally {
        if (Test-Path -LiteralPath $stage) {
            Remove-Item -LiteralPath $stage -Recurse -Force -ErrorAction SilentlyContinue
        }
        if (Test-Path -LiteralPath $stageRoot) {
            Remove-Item -LiteralPath $stageRoot -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

function Invoke-CreatePlaylist {
    param(
        [Parameter(Mandatory = $true)] $Items,
        [Parameter(Mandatory = $true)][string] $PlaylistFolder,
        [Parameter(Mandatory = $true)][ValidateSet('Audio', 'Video')] [string] $Mode,
        [Parameter(Mandatory = $true)][string] $LogPath
    )

    $stageRoot = Join-Path $PlaylistFolder '__yt-dlp_stage'
    Initialize-StageRoot -StageRoot $stageRoot
    $workerRoot = Join-Path $stageRoot 'workers'
    New-Item -ItemType Directory -Path $workerRoot -Force | Out-Null
    $labelForWorkers = $SourceLabel
    if ([string]::IsNullOrWhiteSpace($labelForWorkers)) {
        $labelForWorkers = 'playlist'
    }

    Set-ProcessTitle -Title ("yt-dlp | Create | {0}" -f $labelForWorkers)
    Write-Host ("Create playlist: {0}" -f $labelForWorkers)

    $buckets = Get-PlaylistDownloadBuckets -Items $Items -Mode $Mode
    $downloadItems = @($buckets.DownloadItems)
    $reviewItems = @($buckets.ReviewItems)

    foreach ($item in $reviewItems) {
        Add-LogRecord -Section 'Skipped By Filter' -Message ($item.title + ' - not music-like')
    }

    $workerInfos = New-Object System.Collections.Generic.List[object]
    if ($downloadItems.Count -gt 0) {
        $sliceSize = [Math]::Ceiling($downloadItems.Count / 4.0)
        if ($sliceSize -lt 1) {
            $sliceSize = 1
        }

        for ($workerIndex = 0; $workerIndex -lt 4; $workerIndex++) {
            $start = $workerIndex * $sliceSize
            if ($start -ge $downloadItems.Count) {
                continue
            }
            $end = [Math]::Min($start + $sliceSize - 1, $downloadItems.Count - 1)
            $slice = @($downloadItems[$start..$end])
            if ($slice.Count -eq 0) {
                continue
            }

            $jobFile = Join-Path $workerRoot ("worker-{0}.json" -f ($workerIndex + 1))
            $workerLog = Join-Path $workerRoot ("worker-{0}.log.json" -f ($workerIndex + 1))
            $slice | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $jobFile -Encoding UTF8

            $workerLabel = 'Worker {0}/4' -f ($workerIndex + 1)
            $process = Start-WorkerProcess -JobFile $jobFile -TempRoot $workerRoot -TargetFolder $PlaylistFolder -WorkerLogFile $workerLog -RunKind Create -Mode $Mode -SourceLabel $labelForWorkers -WorkerLabel $workerLabel
            $workerInfos.Add([pscustomobject]@{
                Process = $process
                LogFile = $workerLog
            }) | Out-Null
        }

        foreach ($info in $workerInfos) {
            Wait-WorkerProcess -Process $info.Process
        }
    }

    $selectedReviewIds = New-Object System.Collections.Generic.HashSet[string]
    if ($reviewItems.Count -gt 0) {
        Set-ProcessTitle -Title ("yt-dlp | Review skipped items | {0}" -f $labelForWorkers)
        $selectedReviewItems = @(Select-SkippedPlaylistItems -Items $reviewItems -SourceLabel $labelForWorkers)
        foreach ($item in $selectedReviewItems) {
            if (-not [string]::IsNullOrWhiteSpace([string]$item.id)) {
                $selectedReviewIds.Add([string]$item.id) | Out-Null
                Add-LogRecord -Section 'Kept After Review' -Message ($item.title + ' (id: ' + $item.id + ')')
            }
        }

        if ($selectedReviewItems.Count -gt 0) {
            Set-ProcessTitle -Title ("yt-dlp | Downloading reviewed items | {0}" -f $labelForWorkers)
            Write-Host 'Downloading reviewed items...'
            foreach ($item in $selectedReviewItems) {
                $reviewStage = Join-Path $workerRoot ('review-' + [Guid]::NewGuid().ToString('N'))
                New-Item -ItemType Directory -Path $reviewStage | Out-Null
                $download = Invoke-DownloadAttempt -Item $item -Mode $Mode -StageFolder $reviewStage -BypassGuard
                if (-not $download.Success) {
                    Add-LogRecord -Section 'Failed After All Attempts' -Message ($item.title + ' (id: ' + $item.id + ')')
                } elseif ($download.UsedFallback) {
                    Add-LogRecord -Section 'Kept After Review' -Message ($item.title + ' -> ' + $download.AttemptLabel)
                }
            }
        }
    }

    Set-ProcessTitle -Title ("yt-dlp | Finalizing playlist | {0}" -f $labelForWorkers)
    $downloadedById = Get-StagedMediaMap -StageRoot $workerRoot
    $finalIndex = 1
    foreach ($item in $Items) {
        if ($Mode -eq 'Audio' -and (Get-MusicalVerdict -Item $item) -eq 'Reject' -and -not $selectedReviewIds.Contains([string]$item.id)) {
            continue
        }
        if (-not $downloadedById.ContainsKey([string]$item.id)) {
            continue
        }

        $staged = Get-Item -LiteralPath $downloadedById[$item.id]
        $stagedInfo = [System.IO.Path]::ChangeExtension($staged.FullName, '.info.json')
        $finalBase = ('{0:0000} - {1}' -f $finalIndex, (Get-SafeFileName -Name $item.title))
        if (-not (Test-Path -LiteralPath $stagedInfo)) {
            $stagedInfo = $null
        }
        $moved = Move-FilePairToStage -SourceMedia $staged.FullName -SourceInfo $stagedInfo -StageFolder $stageRoot -TempBase ([Guid]::NewGuid().ToString('N'))
        Move-StagedPairToFinal -SourceMedia $moved.MediaPath -SourceInfo $moved.InfoPath -SourceId $moved.IdPath -FinalMediaBase $finalBase -TargetFolder $PlaylistFolder -AllowUniqueSuffix -ItemId $item.id | Out-Null
        $finalIndex++
    }

    $logFiles = @($workerInfos | ForEach-Object { $_.LogFile })
    Merge-LogFiles -LogFiles $logFiles -FinalLogPath $LogPath

    if (Test-Path -LiteralPath $stageRoot) {
        Remove-Item -LiteralPath $stageRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}

function Invoke-UpdateSingle {
    param(
        [Parameter(Mandatory = $true)] $Item,
        [Parameter(Mandatory = $true)][string] $TargetFolder,
        [Parameter(Mandatory = $true)][ValidateSet('Audio', 'Video')] [string] $Mode
    )

    if ($Mode -eq 'Audio' -and (Get-MusicalVerdict -Item $Item) -eq 'Reject') {
        Add-LogRecord -Section 'Skipped By Filter' -Message ($Item.title + ' - not music-like')
        return
    }

    $existing = Get-ExistingCatalog -Folder $TargetFolder
    if ($existing | Where-Object { $_.Id -eq $Item.id }) {
        return
    }

    $stageRoot = Join-Path $TargetFolder '__yt-dlp_stage'
    Initialize-StageRoot -StageRoot $stageRoot
    $stage = Join-Path $stageRoot ([Guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $stage | Out-Null

    try {
        $download = Invoke-DownloadAttempt -Item $Item -Mode $Mode -StageFolder $stage
        if (-not $download.Success) {
            if ($download.FilteredOut) {
                Add-LogRecord -Section 'Skipped By Filter' -Message ($Item.title + ' - direct YouTube Music download was not music-like')
                return
            }
            Add-LogRecord -Section 'Failed After All Attempts' -Message ($Item.title + ' (id: ' + $Item.id + ')')
            return
        }

        $existingCatalog = Get-ExistingCatalog -Folder $TargetFolder
        $numbered = @($existingCatalog | Where-Object { $_.IsNumbered } | Sort-Object Index, BaseName)

        $finalItems = New-Object System.Collections.Generic.List[object]
        $finalItems.Add([pscustomobject]@{
            Item = $Item
            SourceMedia = $download.MediaPath
            SourceInfo = $download.InfoPath
            SourceId = $download.IdPath
            ItemId = $Item.id
            FinalBase = ('0001 - ' + (Get-SafeFileName -Name $Item.title))
        }) | Out-Null

        $index = 2
        foreach ($entry in $numbered) {
            $media = $entry.MediaPath
            $info = $entry.InfoPath
            $baseName = $entry.BaseName
            $base = ('{0:0000} - {1}' -f $index, (Get-SafeFileName -Name $baseName.Substring(7)))
            $entryId = $entry.Id
            $finalItems.Add([pscustomobject]@{
                Item = $entry
                SourceMedia = $media
                SourceInfo = $info
                SourceId = $entry.IdPath
                ItemId = $entryId
                FinalBase = $base
            }) | Out-Null
            $index++
        }

        foreach ($entry in $finalItems) {
            $staged = Move-FilePairToStage -SourceMedia $entry.SourceMedia -SourceInfo $entry.SourceInfo -StageFolder $stage -TempBase ([Guid]::NewGuid().ToString('N'))
            Move-StagedPairToFinal -SourceMedia $staged.MediaPath -SourceInfo $staged.InfoPath -SourceId $staged.IdPath -FinalMediaBase $entry.FinalBase -TargetFolder $TargetFolder -AllowUniqueSuffix -ItemId $entry.ItemId | Out-Null
        }

        if ($download.UsedFallback) {
            Add-LogRecord -Section 'Fallback Success' -Message ($Item.title + ' -> ' + $download.AttemptLabel)
        }
    } finally {
        if (Test-Path -LiteralPath $stage) {
            Remove-Item -LiteralPath $stage -Recurse -Force -ErrorAction SilentlyContinue
        }
        if (Test-Path -LiteralPath $stageRoot) {
            Remove-Item -LiteralPath $stageRoot -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

function Invoke-UpdatePlaylist {
    param(
        [Parameter(Mandatory = $true)] $Items,
        [Parameter(Mandatory = $true)][string] $TargetFolder,
        [Parameter(Mandatory = $true)][ValidateSet('Audio', 'Video')] [string] $Mode,
        [Parameter(Mandatory = $true)][string] $LogPath
    )

    $existing = Get-ExistingCatalog -Folder $TargetFolder
    $existingById = @{}
    foreach ($entry in $existing) {
        if (-not [string]::IsNullOrWhiteSpace($entry.Id) -and -not $existingById.ContainsKey($entry.Id)) {
            $existingById[$entry.Id] = $entry
        }
    }

    $missing = New-Object System.Collections.Generic.List[object]
    $reviewItems = New-Object System.Collections.Generic.List[object]

    foreach ($item in $Items) {
        if ($Mode -eq 'Audio' -and (Get-MusicalVerdict -Item $item) -eq 'Reject') {
            Add-LogRecord -Section 'Skipped By Filter' -Message ($item.title + ' - not music-like')
            $reviewItems.Add($item) | Out-Null
            continue
        }

        if (-not $existingById.ContainsKey($item.id)) {
            $missing.Add($item) | Out-Null
        }
    }

    $stageRoot = Join-Path $TargetFolder '__yt-dlp_stage'
    Initialize-StageRoot -StageRoot $stageRoot
    $workerRoot = Join-Path $stageRoot 'workers'
    New-Item -ItemType Directory -Path $workerRoot -Force | Out-Null
    $labelForWorkers = $SourceLabel
    if ([string]::IsNullOrWhiteSpace($labelForWorkers)) {
        $labelForWorkers = 'playlist'
    }
    Set-ProcessTitle -Title ("yt-dlp | Update | {0}" -f $labelForWorkers)
    Write-Host ("Update playlist: {0}" -f $labelForWorkers)

    $workerInfos = New-Object System.Collections.Generic.List[object]
    if ($missing.Count -gt 0) {
        $sliceSize = [Math]::Ceiling($missing.Count / 4.0)
        if ($sliceSize -lt 1) {
            $sliceSize = 1
        }

        for ($workerIndex = 0; $workerIndex -lt 4; $workerIndex++) {
            $start = $workerIndex * $sliceSize
            if ($start -ge $missing.Count) {
                continue
            }
            $end = [Math]::Min($start + $sliceSize - 1, $missing.Count - 1)
            $slice = @($missing[$start..$end])
            if ($slice.Count -eq 0) {
                continue
            }

            $jobFile = Join-Path $workerRoot ("worker-{0}.json" -f ($workerIndex + 1))
            $workerLog = Join-Path $workerRoot ("worker-{0}.log.json" -f ($workerIndex + 1))
            $slice | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $jobFile -Encoding UTF8

            $workerLabel = 'Worker {0}/4' -f ($workerIndex + 1)
            $process = Start-WorkerProcess -JobFile $jobFile -TempRoot $workerRoot -TargetFolder $TargetFolder -WorkerLogFile $workerLog -RunKind Update -Mode $Mode -SourceLabel $labelForWorkers -WorkerLabel $workerLabel
            $workerInfos.Add([pscustomobject]@{
                Process = $process
                LogFile = $workerLog
            }) | Out-Null
        }

        foreach ($info in $workerInfos) {
            Wait-WorkerProcess -Process $info.Process
        }
    }

    $selectedReviewIds = New-Object System.Collections.Generic.HashSet[string]
    if ($reviewItems.Count -gt 0) {
        Set-ProcessTitle -Title ("yt-dlp | Review skipped items | {0}" -f $labelForWorkers)
        $selectedReviewItems = @(Select-SkippedPlaylistItems -Items $reviewItems -SourceLabel $labelForWorkers)
        foreach ($item in $selectedReviewItems) {
            if (-not [string]::IsNullOrWhiteSpace([string]$item.id)) {
                $selectedReviewIds.Add([string]$item.id) | Out-Null
                Add-LogRecord -Section 'Kept After Review' -Message ($item.title + ' (id: ' + $item.id + ')')
            }
        }

        if ($selectedReviewItems.Count -gt 0) {
            Set-ProcessTitle -Title ("yt-dlp | Downloading reviewed items | {0}" -f $labelForWorkers)
            Write-Host 'Downloading reviewed items...'
            foreach ($item in $selectedReviewItems) {
                if ($existingById.ContainsKey($item.id)) {
                    continue
                }

                $reviewStage = Join-Path $workerRoot ('review-' + [Guid]::NewGuid().ToString('N'))
                New-Item -ItemType Directory -Path $reviewStage | Out-Null
                $download = Invoke-DownloadAttempt -Item $item -Mode $Mode -StageFolder $reviewStage -BypassGuard
                if (-not $download.Success) {
                    Add-LogRecord -Section 'Failed After All Attempts' -Message ($item.title + ' (id: ' + $item.id + ')')
                } elseif ($download.UsedFallback) {
                    Add-LogRecord -Section 'Kept After Review' -Message ($item.title + ' -> ' + $download.AttemptLabel)
                }
            }
        }
    }

    Set-ProcessTitle -Title ("yt-dlp | Finalizing playlist | {0}" -f $labelForWorkers)
    $downloadedById = Get-StagedMediaMap -StageRoot $workerRoot

    $finalItems = New-Object System.Collections.Generic.List[object]
    foreach ($item in $Items) {
        $isRejected = ($Mode -eq 'Audio' -and (Get-MusicalVerdict -Item $item) -eq 'Reject')
        if ($isRejected -and -not $selectedReviewIds.Contains([string]$item.id)) {
            continue
        }

        $sourceMedia = $null
        $sourceInfo = $null
        $sourceId = $null

        if ($existingById.ContainsKey($item.id)) {
            $sourceMedia = $existingById[$item.id].MediaPath
            $sourceInfo = $existingById[$item.id].InfoPath
            $sourceId = $existingById[$item.id].IdPath
        } elseif ($downloadedById.ContainsKey($item.id)) {
            $sourceMedia = $downloadedById[$item.id]
            $sourceInfo = [System.IO.Path]::ChangeExtension($sourceMedia, '.info.json')
            if (-not (Test-Path -LiteralPath $sourceInfo)) {
                $sourceInfo = $null
            }
            $sourceId = [System.IO.Path]::ChangeExtension($sourceMedia, '.yt-dlp.id')
            if (-not (Test-Path -LiteralPath $sourceId)) {
                $sourceId = $null
            }
        } else {
            Add-LogRecord -Section 'Failed After All Attempts' -Message ($item.title + ' (id: ' + $item.id + ')')
            continue
        }

        $finalItems.Add([pscustomobject]@{
            Item = $item
            SourceMedia = $sourceMedia
            SourceInfo = $sourceInfo
            SourceId = $sourceId
            ItemId = $item.id
            FinalBase = $null
            StagedMedia = $null
            StagedInfo  = $null
            StagedId    = $null
        }) | Out-Null
    }

    foreach ($slot in $finalItems) {
        $staged = Move-FilePairToStage -SourceMedia $slot.SourceMedia -SourceInfo $slot.SourceInfo -StageFolder $stageRoot -TempBase ([Guid]::NewGuid().ToString('N'))
        $slot.StagedMedia = $staged.MediaPath
        $slot.StagedInfo = $staged.InfoPath
        $slot.StagedId = $staged.IdPath
    }

    Remove-StaleNumberedFiles -Folder $TargetFolder

    $finalIndex = 1
    foreach ($slot in $finalItems) {
        $slot.FinalBase = ('{0:0000} - {1}' -f $finalIndex, (Get-SafeFileName -Name $slot.Item.title))
        Move-StagedPairToFinal -SourceMedia $slot.StagedMedia -SourceInfo $slot.StagedInfo -SourceId $slot.StagedId -FinalMediaBase $slot.FinalBase -TargetFolder $TargetFolder -AllowUniqueSuffix -ItemId $slot.ItemId | Out-Null
        $finalIndex++
    }

    $logFiles = @($workerInfos | ForEach-Object { $_.LogFile })
    Merge-LogFiles -LogFiles $logFiles -FinalLogPath $LogPath

    if (Test-Path -LiteralPath $stageRoot) {
        Remove-Item -LiteralPath $stageRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}

function Get-RunLogPath {
    param([Parameter(Mandatory = $true)][string] $Label)

    $safeLabel = Get-SafeFileName -Name $Label
    $dateStamp = Get-Date -Format 'dd-MM-yyyy'
    return (Join-Path $downloadRoot ("{0} {1}.log" -f $safeLabel, $dateStamp))
}

function Invoke-InteractiveMode {
    Assert-Workspace
    Set-ProcessTitle -Title 'yt-dlp | waiting for input'
    Write-Host 'Waiting for input...'

    Write-Host '1. Create new'
    Write-Host '2. Update existing'
    $choice = ''
    while ($choice -notin @('1', '2')) {
        $choice = (Read-Host 'Select action').Trim()
    }

    $url = ''
    while ([string]::IsNullOrWhiteSpace($url)) {
        $url = (Read-Host 'Enter URL').Trim()
    }

    Write-Host '1. Audio only'
    Write-Host '2. Video only'
    $modeChoice = ''
    while ($modeChoice -notin @('1', '2')) {
        $modeChoice = (Read-Host 'Select mode').Trim()
    }
    $selectedMode = if ($modeChoice -eq '1') { 'Audio' } else { 'Video' }

    $flatEntries = Get-FlatEntries -Url $url
    $sourceIsPlaylist = ($flatEntries.Count -gt 1)
    $sourceItem = $null
    $sourceItems = @()

    if ($sourceIsPlaylist) {
        $sourceItems = @($flatEntries | Sort-Object { [int]$_.playlist_index }, id)
        $sourceLabel = [string]$sourceItems[0].playlist_title
        if ([string]::IsNullOrWhiteSpace($sourceLabel)) {
            $sourceLabel = [string]$sourceItems[0].playlist
        }
        if ([string]::IsNullOrWhiteSpace($sourceLabel)) {
            $sourceLabel = [string]$sourceItems[0].title
        }
        if ([string]::IsNullOrWhiteSpace($sourceLabel)) {
            $sourceLabel = 'playlist'
        }
    } else {
        $sourceItem = Get-VideoDetails -Url $url
        if (-not $sourceItem) {
            throw "Unable to read metadata for $url"
        }
        $sourceLabel = [string]$sourceItem.title
        if ([string]::IsNullOrWhiteSpace($sourceLabel)) {
            $sourceLabel = 'single video'
        }
    }

    Set-ProcessTitle -Title ("yt-dlp | {0} | {1}" -f $selectedMode, $sourceLabel)
    Write-Host ("Selected mode: {0}" -f $selectedMode)
    Write-Host ("Source: {0}" -f $sourceLabel)

    if ($choice -eq '1') {
        if ($sourceIsPlaylist) {
            $playlistFolderBase = Join-Path $downloadRoot (Get-SafeFileName -Name $sourceLabel)
            $playlistFolder = Get-UniqueFolderPath -BasePath $playlistFolderBase
            New-Item -ItemType Directory -Path $playlistFolder | Out-Null
            $logPath = Get-RunLogPath -Label $sourceLabel
            Invoke-CreatePlaylist -Items $sourceItems -PlaylistFolder $playlistFolder -Mode $selectedMode -LogPath $logPath
        } else {
            $logPath = Get-RunLogPath -Label $sourceLabel
            Invoke-CreateSingle -Item $sourceItem -TargetFolder $downloadRoot -Mode $selectedMode
            Merge-LogFiles -LogFiles @() -FinalLogPath $logPath
        }
    } else {
        $targetSelection = Select-UpdateTargetFolder -DownloadRoot $downloadRoot
        while ($null -eq $targetSelection) {
            Write-Host 'No folder selected. Choose the custom folder option again or use option 1.'
            $targetSelection = Select-UpdateTargetFolder -DownloadRoot $downloadRoot
        }
        $targetFolder = $targetSelection.Path
        $logPath = Get-RunLogPath -Label $targetSelection.Label

        if ($sourceIsPlaylist) {
            $sourceItems = @($sourceItems | Sort-Object { [int]$_.playlist_index }, id)
            Invoke-UpdatePlaylist -Items $sourceItems -TargetFolder $targetFolder -Mode $selectedMode -LogPath $logPath
        } else {
            Invoke-UpdateSingle -Item $sourceItem -TargetFolder $targetFolder -Mode $selectedMode
            Merge-LogFiles -LogFiles @() -FinalLogPath $logPath
        }
    }
}

if (-not $NoRun) {
    if ($Worker) {
        Invoke-WorkerMode
    } else {
        Invoke-InteractiveMode
    }
}
