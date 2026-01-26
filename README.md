# MagickJoin

当アプリは、ImageMagickのフロントエンドです。

---
## このアプリについて
本アプリは **Microsoft Copilot によって作成された Copilot 製アプリ**です。  
Windows 11 + PowerShell 7.5.4 で動作確認済みですが、5.1でも動作します。

![ImageMagick](https://github.com/user-attachments/assets/adb7da8e-002d-4d7f-aeaf-068a0919de25 "アプリウィンドウ")
---
## 本アプリはPowerShell と WPFで動作します、別途ランタイム等のインストールは必要ありません。
ImageMagickをインストールし、コマンドから使用できるようにしてください。

---
### 使用方法

- ２枚の画像ファイルか画像の含まれるフォルダをドロップすると対応形式の画像を連結します。
- 連結は１枚目の画像が縦長であれば横に、横長であれば縦に連結します。
- 連結はかならず２枚単位で行い、奇数ファイルをドロップした場合、最後の画像は無視されます。
- 対応形式以外のファイルは全て無視されます。
- ドロップした際に掴んでいたファイルを１枚目としてファイル名順に処理していきます。
- 結合後の画像は元フォルダに保存され、元の画像は自動的に `Processed` フォルダへ移動されます。

---

## 特徴

- ドラッグ＆ドロップのみで操作可能
- 横連結 / 縦連結を自動判定
- **右（または上）が一枚目、左（または下）が二枚目** となるように配置
- ImageMagick の montage を内部で使用
- PowerShell + WPF による軽量 GUI
- Processed フォルダへ自動整理
- MIT License で公開

---

## 動作要件

- Windows 10 / 11
- PowerShell 7.x
- ImageMagick（`magick.exe` が PATH に通っていること）
- .NET Framework / WPF が動作する環境

---

## インストール

1. リポジトリをクローンまたは ZIP ダウンロード  
2. `画像連結.ps1` を任意の場所に配置  
3. ImageMagick がインストールされていることを確認  
4. PowerShell 7 で起動

```powershell
pwsh ./画像連結.ps1
# PowerShell 5.1 の場合
powershell ./画像連結.ps1
```
```ショートカットで使用する場合
pwsh -WindowStyle Hidden -ExecutionPolicy Bypass -File .\画像連結.ps1
```
