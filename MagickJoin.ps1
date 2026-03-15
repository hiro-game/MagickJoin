# 1. アセンブリロード (WPF関連を排除し起動を高速化)
Add-Type -AssemblyName System.Windows.Forms, System.Drawing

# 2. ドラッグ移動用の変数初期化 (コンパイル不要・最速)
$script:dragging = $false
$script:mouseOffset = New-Object Drawing.Point(0,0)

# --- 設定 ---
$ImageMagick = "magick.exe"
$script:ImageExtensions = '.jpg','.jpeg','.png','.bmp','.gif','.tif','.tiff','.webp'

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

# ドラッグ移動ロジック
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

# 1. 📌 ピンボタン
$btnPin = New-Object Windows.Forms.Button
$btnPin.Text = "📌" ; $btnPin.Size = New-Object Drawing.Size(32, 32) ; $btnPin.Dock = "Right"
$btnPin.FlatStyle = "Flat" ; $btnPin.FlatAppearance.BorderSize = 0
$btnPin.Font = New-Object Drawing.Font("Segoe UI Emoji", 10)
$btnPin.ForeColor = [Drawing.Color]::FromArgb(120, 255, 255, 255)
$btnPin.Add_Click({
    $form.TopMost = -not $form.TopMost
    if ($form.TopMost) { $btnPin.Text = "📍"; $btnPin.ForeColor = [Drawing.Color]::White }
    else { $btnPin.Text = "📌"; $btnPin.ForeColor = [Drawing.Color]::FromArgb(120, 255, 255, 255) }
})

# 2. ✕ 閉じるボタン
$btnClose = New-Object Windows.Forms.Button
$btnClose.Text = "✕" ; $btnClose.Size = New-Object Drawing.Size(32, 32) ; $btnClose.Dock = "Right"
$btnClose.FlatStyle = "Flat" ; $btnClose.FlatAppearance.BorderSize = 0 
$btnClose.Font = New-Object Drawing.Font("MS Gothic", 10)
$btnClose.Add_Click({ $form.Close() })

$pnlTitle.Controls.Add($btnPin)
$pnlTitle.Controls.Add($btnClose)

$lblTitle = New-Object Windows.Forms.Label
$lblTitle.Text = "画像連結ツール (ImageMagick)" ; $lblTitle.Location = New-Object Drawing.Point(10, 8) ; $lblTitle.AutoSize = $true
$lblTitle.Font = New-Object Drawing.Font("MS Gothic", 9, [Drawing.FontStyle]::Bold) ; $lblTitle.Enabled = $false
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
function Add-Log([string]$Message, [string]$Type = "info") {
    $timestamp = (Get-Date).ToString("HH:mm:ss")
    $logBox.AppendText("$timestamp  $Message`r`n")
    $logBox.ScrollToCaret()
}

function Test-ImageByIdentify {
    param([string]$Path)
    $out = & $ImageMagick identify -format "%w %h" -- $Path 2>$null
    if ($LASTEXITCODE -eq 0 -and $out -match "^\d+\s+\d+") {
        return ,@($true, $out)
    }
    return ,@($false, "")
}

function Get-ValidImageFilesFromDrop {
    param([string[]]$Paths)
    $files = @()
    foreach ($p in $Paths) {
        if (Test-Path -LiteralPath $p -PathType Container) {
            $files += [System.IO.Directory]::EnumerateFiles($p) | ForEach-Object { Get-Item -LiteralPath $_ }
        } elseif (Test-Path -LiteralPath $p -PathType Leaf) {
            $files += Get-Item -LiteralPath $p
        }
    }
    $valid = @()
    foreach ($f in $files) {
        if ($script:ImageExtensions -notcontains $f.Extension.ToLower()) { continue }
        $result = Test-ImageByIdentify -Path $f.FullName
        if ($result[0]) { $valid += $f }
        else { Add-Log "非対象(identify失敗): $($f.Name)" "error" }
    }
    return $valid
}

function Merge-Images {
    param([System.IO.FileInfo[]]$Images)
    if ($Images.Count -lt 2) { Add-Log "画像が2枚未満のためスキップ。" "error"; return }

    for ($i = 0; $i + 1 -lt $Images.Count; $i += 2) {
        $f1 = $Images[$i] ; $f2 = $Images[$i + 1]
        $folder = $f1.DirectoryName ; $processDir = Join-Path $folder "Processed"
        
        if (-not (Test-Path -LiteralPath $processDir)) { New-Item -ItemType Directory -Path $processDir | Out-Null }

        $result = Test-ImageByIdentify -Path $f1.FullName
        if (-not $result[0]) { Add-Log "サイズ取得失敗: $($f1.Name)" "error"; continue }

        $parts = $result[1].Trim().Split(" ")
        [int]$w1 = $parts[0] ; [int]$h1 = $parts[1]

        $outName = "{0}-{1}{2}" -f $f1.BaseName, $f2.BaseName, $f1.Extension
        $outPath = Join-Path $folder $outName

        if ($h1 -gt $w1) {
            Add-Log "横連結: 左=$($f2.Name) 右=$($f1.Name)"
            & $ImageMagick montage $f2.FullName $f1.FullName -geometry +0+0 -tile 2x1 $outPath 2>$null
        } else {
            Add-Log "縦連結: 上=$($f1.Name) 下=$($f2.Name)"
            & $ImageMagick montage $f1.FullName $f2.FullName -geometry +0+0 -tile 1x2 $outPath 2>$null
        }

        if ($LASTEXITCODE -eq 0) {
            Add-Log "連結成功: $outName"
            Move-Item -LiteralPath $f1.FullName -Destination $processDir -Force
            Move-Item -LiteralPath $f2.FullName -Destination $processDir -Force
        } else {
            Add-Log "連結失敗: $outName" "error"
        }
    }
    if ($Images.Count % 2 -eq 1) { Add-Log "奇数枚のため最後の1枚はスキップ: $($Images[-1].Name)" }
}

# --- ドロップイベント ---
$form.Add_DragEnter({ if ($_.Data.GetDataPresent([Windows.Forms.DataFormats]::FileDrop)) { $_.Effect = "Copy" } })
$form.Add_DragDrop({
    $paths = $_.Data.GetData([Windows.Forms.DataFormats]::FileDrop)
    $validImages = Get-ValidImageFilesFromDrop -Paths $paths
    if ($validImages.Count -eq 0) { Add-Log "有効な画像がありませんでした。" "error"; return }
    Merge-Images -Images $validImages
    Add-Log "=== すべての処理が完了しました ==="
})

# 起動
[Windows.Forms.Application]::Run($form)
