# 1. アセンブリロード
Add-Type -AssemblyName System.Windows.Forms, System.Drawing

# 2. 変数初期化
$script:dragging = $false
$script:mouseOffset = New-Object Drawing.Point(0,0)

# --- 設定 ---
$ImageMagick = "magick.exe"
$script:ImageExtensions = '.jpg','.jpeg','.png','.bmp','.gif','.tif','.tiff','.webp'
$script:mergeDirection = "Horizontal"
$script:spacing = 0

# --- フォーム構築 ---
$form = New-Object Windows.Forms.Form
$form.Text = "画像連結ツール"
$form.Size = New-Object Drawing.Size(520, 320)
$form.StartPosition = "CenterScreen"
$form.BackColor = [Drawing.Color]::FromArgb(34, 34, 34)
$form.ForeColor = [Drawing.Color]::White
$form.FormBorderStyle = "None"
$form.AllowDrop = $true

# --- タイトルバー ---
$pnlTitle = New-Object Windows.Forms.Panel
$pnlTitle.Height = 32 ; $pnlTitle.Dock = "Top" ; $pnlTitle.BackColor = [Drawing.Color]::FromArgb(51, 51, 51)

$pnlTitle.Add_MouseDown({
    if ($_.Button -eq [Windows.Forms.MouseButtons]::Left) {
        $script:dragging = $true
        $script:mouseOffset = $_.Location
    }
})
$pnlTitle.Add_MouseMove({
    if ($script:dragging) {
        $currentPos = [Windows.Forms.Control]::MousePosition
        $form.Location = New-Object Drawing.Point(($currentPos.X - $script:mouseOffset.X), ($currentPos.Y - $script:mouseOffset.Y))
    }
})
$pnlTitle.Add_MouseUp({ $script:dragging = $false })

# 連結方向切替
$btnDirection = New-Object Windows.Forms.Button
$btnDirection.Text = "横連結" ; $btnDirection.Size = New-Object Drawing.Size(64, 32) ; $btnDirection.Dock = "Right"
$btnDirection.FlatStyle = "Flat" ; $btnDirection.FlatAppearance.BorderSize = 0 ; $btnDirection.TabStop = $false
$btnDirection.Add_Click({
    if ($script:mergeDirection -eq "Horizontal") { $script:mergeDirection = "Vertical"; $btnDirection.Text = "縦連結" }
    else { $script:mergeDirection = "Horizontal"; $btnDirection.Text = "横連結" }
})

# 間隔入力 (マイナス不可)
$lblSpacing = New-Object Windows.Forms.Label
$lblSpacing.Text = "間隔:" ; $lblSpacing.Location = New-Object Drawing.Point(275, 8) ; $lblSpacing.AutoSize = $true

$txtSpacing = New-Object Windows.Forms.TextBox
$txtSpacing.Text = "0" ; $txtSpacing.Size = New-Object Drawing.Size(35, 20) ; $txtSpacing.Location = New-Object Drawing.Point(320, 6)
$txtSpacing.BackColor = [Drawing.Color]::FromArgb(60, 60, 60) ; $txtSpacing.ForeColor = [Drawing.Color]::White
$txtSpacing.BorderStyle = "FixedSingle" ; $txtSpacing.TextAlign = "Center"
$txtSpacing.Add_TextChanged({ 
    if ($this.Text -match "^\d+$") { $script:spacing = [int]$this.Text }
    else { $this.Text = "0" } 
})

# 拡張操作機能 (ホイール・矢印キー)
$txtSpacing.Add_Enter({ $this.SelectAll() })
$txtSpacing.Add_MouseWheel({
    [int]$v = 0 ; if ([int]::TryParse($this.Text, [ref]$v)) {
        if ($_.Delta -gt 0) { $this.Text = ($v + 1).ToString() }
        elseif ($v -gt 0) { $this.Text = ($v - 1).ToString() }
    }
})
$txtSpacing.Add_KeyDown({
    [int]$v = 0 ; if ([int]::TryParse($this.Text, [ref]$v)) {
        if ($_.KeyCode -eq "Up") { $this.Text = ($v + 1).ToString(); $_.Handled = $true }
        elseif ($_.KeyCode -eq "Down" -and $v -gt 0) { $this.Text = ($v - 1).ToString(); $_.Handled = $true }
    }
})

# 上下ボタン
$btnUp = New-Object Windows.Forms.Button
$btnUp.Text = "▲" ; $btnUp.Size = New-Object Drawing.Size(18, 12) ; $btnUp.Location = New-Object Drawing.Point(357, 4)
$btnUp.FlatStyle = "Flat" ; $btnUp.FlatAppearance.BorderSize = 0 ; $btnUp.Font = New-Object Drawing.Font("MS Gothic", 5)
$btnUp.BackColor = [Drawing.Color]::FromArgb(60, 60, 60) ; $btnUp.TabStop = $false ; $btnUp.Tag = $txtSpacing
$btnUp.Add_Click({ [int]$v = 0 ; if ([int]::TryParse($this.Tag.Text, [ref]$v)) { $this.Tag.Text = ($v + 1).ToString() } })

$btnDown = New-Object Windows.Forms.Button
$btnDown.Text = "▼" ; $btnDown.Size = New-Object Drawing.Size(18, 12) ; $btnDown.Location = New-Object Drawing.Point(357, 16)
$btnDown.FlatStyle = "Flat" ; $btnDown.FlatAppearance.BorderSize = 0 ; $btnDown.Font = New-Object Drawing.Font("MS Gothic", 5)
$btnDown.BackColor = [Drawing.Color]::FromArgb(60, 60, 60) ; $btnDown.TabStop = $false ; $btnDown.Tag = $txtSpacing
$btnDown.Add_Click({ [int]$v = 0 ; if ([int]::TryParse($this.Tag.Text, [ref]$v) -and $v -gt 0) { $this.Tag.Text = ($v - 1).ToString() } })

# ピンボタン
$btnPin = New-Object Windows.Forms.Button
$btnPin.Text = "📌" ; $btnPin.Size = New-Object Drawing.Size(32, 32) ; $btnPin.Dock = "Right"
$btnPin.FlatStyle = "Flat" ; $btnPin.FlatAppearance.BorderSize = 0 ; $btnPin.TabStop = $false
$btnPin.Font = New-Object Drawing.Font("Segoe UI Emoji", 10)
$btnPin.ForeColor = [Drawing.Color]::FromArgb(120, 255, 255, 255)
$btnPin.Add_Click({
    $form.TopMost = -not $form.TopMost
    if ($form.TopMost) { $btnPin.Text = "📍"; $btnPin.ForeColor = [Drawing.Color]::White }
    else { $btnPin.Text = "📌"; $btnPin.ForeColor = [Drawing.Color]::FromArgb(120, 255, 255, 255) }
})

# 閉じるボタン
$btnClose = New-Object Windows.Forms.Button
$btnClose.Text = "✕" ; $btnClose.Size = New-Object Drawing.Size(32, 32) ; $btnClose.Dock = "Right"
$btnClose.FlatStyle = "Flat" ; $btnClose.FlatAppearance.BorderSize = 0 ; $btnClose.TabStop = $false
$btnClose.Font = New-Object Drawing.Font("MS Gothic", 10)
$btnClose.Add_Click({ $form.Close() })

$pnlTitle.Controls.AddRange(@($lblSpacing, $txtSpacing, $btnUp, $btnDown, $btnDirection, $btnPin, $btnClose))

$lblTitle = New-Object Windows.Forms.Label
$lblTitle.Text = "画像連結ツール (0/0)" ; $lblTitle.Location = New-Object Drawing.Point(10, 8) ; $lblTitle.AutoSize = $true
$lblTitle.Font = New-Object Drawing.Font("MS Gothic", 9, [Drawing.FontStyle]::Bold)
$pnlTitle.Controls.Add($lblTitle)

# --- ログエリア ---
$logBox = New-Object Windows.Forms.TextBox
$logBox.Multiline = $true ; $logBox.ReadOnly = $true ; $logBox.BackColor = [Drawing.Color]::FromArgb(17, 17, 17)
$logBox.ForeColor = [Drawing.Color]::White ; $logBox.Dock = "Fill" ; $logBox.BorderStyle = "None"
$logBox.ScrollBars = "Vertical" ; $logBox.Font = New-Object Drawing.Font("Consolas", 9)

$container = New-Object Windows.Forms.Panel
$container.Dock = "Fill" ; $container.Padding = New-Object Windows.Forms.Padding(10) ; $container.Controls.Add($logBox)
$form.Controls.AddRange(@($container, $pnlTitle))

# --- ロジック関数群 ---
function Add-Log([string]$Message) {
    $timestamp = (Get-Date).ToString("HH:mm:ss")
    $logBox.AppendText("$timestamp  $Message`r`n")
    $logBox.ScrollToCaret()
}

function Test-ImageByIdentify {
    param([string]$Path)
    $out = & $ImageMagick identify -format "%w %h" -- $Path 2>$null
    if ($LASTEXITCODE -eq 0 -and $out -match "^\d+\s+\d+") { return ,@($true, $out) }
    return ,@($false, "")
}

function Get-ValidImageFilesFromDrop {
    param([string[]]$Paths)
    $files = @()
    foreach ($p in $Paths) {
        if (Test-Path -LiteralPath $p -PathType Container) {
            $files += [System.IO.Directory]::EnumerateFiles($p) | ForEach-Object { Get-Item -LiteralPath $_ }
        } elseif (Test-Path -LiteralPath $p -PathType Leaf) { $files += Get-Item -LiteralPath $p }
    }
    $valid = @()
    foreach ($f in $files) {
        if ($script:ImageExtensions -notcontains $f.Extension.ToLower()) { continue }
        $result = Test-ImageByIdentify -Path $f.FullName
        if ($result[0]) { $valid += $f }
    }
    return $valid
}

function Merge-Images {
    param([System.IO.FileInfo[]]$Images, [int]$TotalSets)
    if ($Images.Count -lt 2) { return }

    # geometryの調整（隙間を画像間に設ける）
    $geoVal = [int]($script:spacing / 2)
    $geometryStr = "+${geoVal}+${geoVal}"

    for ($i = 0; $i + 1 -lt $Images.Count; $i += 2) {
        $f1 = $Images[$i] ; $f2 = $Images[$i + 1]
        $folder = $f1.DirectoryName ; $processDir = Join-Path $folder "Processed"
        if (-not (Test-Path -LiteralPath $processDir)) { New-Item -ItemType Directory -Path $processDir | Out-Null }

        $outName = "{0}-{1}{2}" -f $f1.BaseName, $f2.BaseName, $f1.Extension
        $outPath = Join-Path $folder $outName

        if ($script:mergeDirection -eq "Horizontal") {
            Add-Log "横連結: $($f2.Name) | $($f1.Name) (間隔: $script:spacing)"
            & $ImageMagick montage $f2.FullName $f1.FullName -geometry $geometryStr -tile 2x1 -background white $outPath 2>$null
        } else {
            Add-Log "縦連結: $($f1.Name) / $($f2.Name) (間隔: $script:spacing)"
            & $ImageMagick montage $f1.FullName $f2.FullName -geometry $geometryStr -tile 1x2 -background white $outPath 2>$null
        }

        if ($LASTEXITCODE -eq 0) {
            Move-Item -LiteralPath $f1.FullName -Destination $processDir -Force
            Move-Item -LiteralPath $f2.FullName -Destination $processDir -Force
        }
        
        $currentSet = [int]($i / 2) + 1
        $lblTitle.Text = "画像連結ツール ($currentSet/$TotalSets)" ; $lblTitle.Refresh()
    }
}

$form.Add_DragEnter({ if ($_.Data.GetDataPresent([Windows.Forms.DataFormats]::FileDrop)) { $_.Effect = "Copy" } })
$form.Add_DragDrop({
    $paths = $_.Data.GetData([Windows.Forms.DataFormats]::FileDrop)
    $validImages = Get-ValidImageFilesFromDrop -Paths $paths
    $totalSets = [Math]::Floor($validImages.Count / 2)
    
    if ($validImages.Count -gt 0) { 
        $lblTitle.Text = "画像連結ツール (0/$totalSets)"
        Merge-Images -Images $validImages -TotalSets $totalSets
        Add-Log "=== すべての処理が完了しました ==="
    }
})

[Windows.Forms.Application]::Run($form)