# Export Repo Text

リポジトリのディレクトリ構造とテキストファイルの内容を、1つの `.txt` ファイルへ出力するPowerShellツールです。

## ファイル

- `Export-RepoText.ps1` — エクスポート本体
- `Export-RepoText.bat` — フォルダをドラッグ&ドロップして実行する入口

## 使い方

Windowsでリポジトリフォルダを `Export-RepoText.bat` にドラッグ&ドロップします。
複数フォルダをまとめてドロップした場合は、順番に処理します。

PowerShellから直接実行する場合:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Export-RepoText.ps1 -Root C:\repos\pub\watchbox -OutFile C:\temp\watchbox-repo.txt
```

PowerShell 7 / WSL から実行する場合:

```bash
pwsh ./Export-RepoText.ps1 -Root /home/ubuntu/repos/pub/watchbox -OutFile /tmp/watchbox-repo.txt
```

`-OutFile` を省略すると、対象フォルダの親ディレクトリへ `<repo-name>-repo-dump-<timestamp>.txt` を出力します。

## 出力内容

- `TREE` — ディレクトリ構造
- `SKIPPED FILES` — 除外・省略されたファイルと理由
- `FILE CONTENTS` — 読み取れたテキストファイルの本文

## 主な除外対象

- `.git`, `node_modules`, `bin`, `obj`, `dist`, `build`, `.venv` などの生成物・依存ディレクトリ
- `.env`, 秘密鍵、証明書などの秘匿ファイル
- 既定で1MBを超えるファイル
- バイナリ拡張子のファイル
- NUL byte を含むファイル

除外リストや上限サイズは、PowerShell引数で変更できます。
