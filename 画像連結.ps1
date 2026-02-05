# Requires -Version 5.1
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase

# ImageMagick 実行ファイル
$ImageMagick = "magick.exe"

# --- XAML ---
$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="画像連結ツール (ImageMagick)"
        Height="260" Width="520"
        WindowStartupLocation="CenterScreen"
        ResizeMode="CanMinimize"
        AllowDrop="True"
        WindowStyle="None"
        Background="White">

    <Border BorderBrush="Gray" BorderThickness="1">
        <Grid>

            <!-- タイトルバー -->
            <Grid Background="#FFEEEEEE" Height="32" VerticalAlignment="Top">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>

                <TextBlock Text="画像連結ツール (ImageMagick)"
                           VerticalAlignment="Center"
                           Margin="10,0,0,0"
                           FontWeight="Bold"/>

                <Button x:Name="PinButton"
                        Grid.Column="1"
                        Content="📌"
                        Width="32" Height="32"
                        Background="Transparent"
                        BorderThickness="0"
                        FontSize="16"
                        ToolTip="最前面に固定"/>

                <Button x:Name="CloseButton"
                        Grid.Column="2"
                        Content="✕"
                        Width="32" Height="32"
                        Background="Transparent"
                        BorderThickness="0"
                        FontSize="14"
                        ToolTip="閉じる"/>
            </Grid>

            <!-- 本体 -->
            <Grid Margin="0,32,0,0">
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                </Grid.RowDefinitions>

                <Border Grid.Row="0" BorderBrush="Gray" BorderThickness="1" Padding="8" Margin="8,8,8,6" Background="#FFFAFAFA">
                    <TextBlock Text="ここに画像ファイルまたはフォルダをドラッグ＆ドロップしてください。" 
                               HorizontalAlignment="Center" VerticalAlignment="Center"
                               TextWrapping="Wrap" />
                </Border>

                <!-- RichTextBox + DropLayer -->
                <Grid Grid.Row="1" Margin="8,0,8,8">

                    <!-- 表示専用ログウィンドウ -->
                    <RichTextBox x:Name="LogBox"
                                 IsReadOnly="True"
                                 VerticalScrollBarVisibility="Auto"
                                 HorizontalScrollBarVisibility="Auto"
                                 Background="White"
                                 BorderBrush="Gray"
                                 BorderThickness="1"
                                 FontFamily="Consolas"
                                 FontSize="12" />

                    <!-- ドロップ専用レイヤー -->
                    <Border x:Name="DropLayer"
                            Background="Transparent"
                            AllowDrop="True" />
                </Grid>

            </Grid>

        </Grid>
    </Border>
</Window>
"@

# --- WPF Window ---
$reader = New-Object System.Xml.XmlNodeReader ([xml]$xaml)
$window = [System.Windows.Markup.XamlReader]::Load($reader)

# --- Window 全域ドロップ ---
$window.Add_DragOver({
    $_.Effects = [System.Windows.DragDropEffects]::Copy
    $_.Handled = $true
})

# --- タイトルバー移動 ---
$window.Add_MouseLeftButtonDown({
    if ($_.GetPosition($window).Y -lt 32) {
        $window.DragMove()
    }
})

# --- ピン固定 ---
$pinButton = $window.FindName("PinButton")
$pinButton.Add_Click({
    if ($window.Topmost) {
        $window.Topmost = $false
        $pinButton.Content = "📌"
    }
    else {
        $window.Topmost = $true
        $pinButton.Content = "📍"
    }
})

$logBox = $window.FindName("LogBox")
$closeButton = $window.FindName("CloseButton")
$dropLayer = $window.FindName("DropLayer")
$dropLayer.Add_PreviewMouseWheel({ $_.Handled = $false })
# --- DropLayer（RichTextBox 上のドロップを Window にバブル） ---
$dropLayer.Add_DragOver({
    if ($_.Data.GetDataPresent([Windows.DataFormats]::FileDrop)) {
        $_.Effects = [System.Windows.DragDropEffects]::Copy
    }
    $_.Handled = $false
})

$dropLayer.Add_Drop({
    $_.Handled = $false
})

# --- Add-Log ---
function Add-Log {
    param([string]$Message, [string]$Type = "info")

    $timestamp = (Get-Date).ToString("HH:mm:ss")
    $text = "$timestamp  $Message`r`n"

    $range = New-Object System.Windows.Documents.TextRange(
        $logBox.Document.ContentEnd,
        $logBox.Document.ContentEnd
    )
    $range.Text = $text

    switch ($Type) {
        "success" { $range.ApplyPropertyValue([System.Windows.Documents.TextElement]::ForegroundProperty, "Green") }
        "error"   { $range.ApplyPropertyValue([System.Windows.Documents.TextElement]::ForegroundProperty, "Red") }
        default   { $range.ApplyPropertyValue([System.Windows.Documents.TextElement]::ForegroundProperty, "Black") }
    }

    $logBox.ScrollToEnd()
}

# --- identify（復元：画像判定に必須） ---
function Test-ImageByIdentify {
    param([string]$Path)

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $ImageMagick
    $psi.Arguments = "identify -format `"%w %h`" `"$Path`""
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError  = $true
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true

    $proc = New-Object System.Diagnostics.Process
    $proc.StartInfo = $psi
    $null = $proc.Start()

    $out = $proc.StandardOutput.ReadToEnd()
    $err = $proc.StandardError.ReadToEnd()
    $proc.WaitForExit()

    if ($proc.ExitCode -eq 0 -and $out -match "^\d+\s+\d+") {
        return ,@($true, $out)
    }
    else {
        return ,@($false, $err)
    }
}

# --- ドロップされたファイルから有効画像を抽出 ---
function Get-ValidImageFilesFromDrop {
    param([string[]]$Paths)

    $files = @()

    foreach ($p in $Paths) {
        if (Test-Path -LiteralPath $p -PathType Container) {
            $files += [System.IO.Directory]::EnumerateFiles($p) | ForEach-Object {
                Get-Item -LiteralPath $_
            }
        }
        elseif (Test-Path -LiteralPath $p -PathType Leaf) {
            $files += Get-Item -LiteralPath $p
        }
    }

    $valid = @()

    foreach ($f in $files) {
        $result = Test-ImageByIdentify -Path $f.FullName

        if ($result[0] -eq $true) {
            Add-Log "対象画像: $($f.FullName)"
            $valid += $f
        }
        else {
            Add-Log "非対象(identify失敗): $($f.FullName)"
        }
    }

    return $valid
}

# --- Drop event ---
$window.Add_Drop({
    $paths = $_.Data.GetData([Windows.DataFormats]::FileDrop)
    Add-Log "ドロップ受付: $($paths -join ', ')" "info"

    $validImages = Get-ValidImageFilesFromDrop -Paths $paths

    if ($validImages.Count -eq 0) {
        Add-Log "有効な画像がありませんでした。" "error"
        return
    }

    Merge-Images -Images $validImages
    Add-Log "処理完了。" "success"
})

# --- マージ処理（Arguments 方式） ---
function Merge-Images {
    param([System.IO.FileInfo[]]$Images)

    if ($Images.Count -lt 2) {
        Add-Log "画像が2枚未満のためスキップ。" "error"
        return
    }

    for ($i = 0; $i + 1 -lt $Images.Count; $i += 2) {

        $f1 = $Images[$i]
        $f2 = $Images[$i + 1]

        $folder = $f1.DirectoryName
        $processDir = Join-Path $folder "Processed"
        
        # Processed が存在する場合、フォルダかどうか確認
        if (Test-Path -LiteralPath $processDir) {
            $item = Get-Item -LiteralPath $processDir
            if (-not $item.PSIsContainer) {
                # ファイルだった場合は削除してフォルダを作り直す
                Remove-Item -LiteralPath $processDir -Force
                New-Item -ItemType Directory -Path $processDir | Out-Null
            }
        }
        else {
            # 存在しない場合はフォルダを作成
            New-Item -ItemType Directory -Path $processDir | Out-Null
        }

        $result = Test-ImageByIdentify -Path $f1.FullName
        if ($result[0] -eq $false) {
            Add-Log "縦横比取得失敗のためスキップ: $($f1.FullName)" "error"
            continue
        }

        $parts = $result[1].Trim().Split(" ")
        [int]$w1 = $parts[0]
        [int]$h1 = $parts[1]

        if ($h1 -gt $w1) {
            # 横連結（2x1）
            $tile = "2x1"
            Add-Log "横に連結: 左=$($f2.Name) 右=$($f1.Name)" "info"
        }
        else {
            # 縦連結（1x2）
            $tile = "1x2"
            Add-Log "縦に連結: 上=$($f1.Name) 下=$($f2.Name)" "info"
        }

        $outName = "{0}-{1}{2}" -f $f1.BaseName, $f2.BaseName, $f1.Extension
        $outPath = Join-Path $folder $outName

        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = $ImageMagick
        if ($tile -eq "2x1") {
            # 横：左=f2, 右=f1
            $psi.Arguments = "montage `"$($f2.FullName)`" `"$($f1.FullName)`" -geometry +0+0 -tile 2x1 `"$outPath`""
        }
        else {
            # 縦：上=f1, 下=f2
            $psi.Arguments = "montage `"$($f1.FullName)`" `"$($f2.FullName)`" -geometry +0+0 -tile 1x2 `"$outPath`""
        }
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError  = $true
        $psi.UseShellExecute = $false
        $psi.CreateNoWindow = $true

        $proc = New-Object System.Diagnostics.Process
        $proc.StartInfo = $psi
        $null = $proc.Start()

        $out = $proc.StandardOutput.ReadToEnd()
        $err = $proc.StandardError.ReadToEnd()
        $proc.WaitForExit()

        if ($proc.ExitCode -eq 0) {
            Add-Log "連結成功: $outName" "success"
            Move-Item -LiteralPath $f1.FullName -Destination $processDir -Force
            Move-Item -LiteralPath $f2.FullName -Destination $processDir -Force
        }
        else {
            Add-Log "連結失敗: $outName  エラー: $err" "error"
        }
    }

    if ($Images.Count % 2 -eq 1) {
        Add-Log "奇数枚のため最後の1枚はスキップ: $($Images[-1].FullName)" "info"
    }
}

$closeButton.Add_Click({ $window.Close() })

$window.ShowDialog() | Out-Null
