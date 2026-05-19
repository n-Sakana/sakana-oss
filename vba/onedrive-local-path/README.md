# OneDrive / SharePoint URL to Local Path for VBA

Excel VBA で `ThisWorkbook.Path` が SharePoint / OneDrive の URL を返す場合に、OneDrive 同期済みのローカルパスへ変換するための単一モジュールです。

## 使い方

`ConvURLtoLocalPath.bas` を VBA プロジェクトへインポートします。公開関数は `convURLtoLocalPath` です。

```vb
Dim basePath As String
basePath = convURLtoLocalPath(ThisWorkbook)
```

`Workbook` オブジェクトを渡した場合は `.Path` を解決します。文字列を渡した場合は、その文字列を解決します。

```vb
Dim folderPath As String
Dim filePath As String

folderPath = convURLtoLocalPath(ThisWorkbook.Path)
filePath = convURLtoLocalPath(ThisWorkbook.FullName)
```

変換できない場合、既定では入力値をそのまま返します。失敗時に空文字を返したい場合は次のようにします。

```vb
folderPath = convURLtoLocalPath(ThisWorkbook, returnInputOnFail:=False)
```

## 方針

- Win32 API `Declare` は使いません。
- `Shell` / `WScript.Shell` / `Shell.Application` / PowerShell は使いません。
- レジストリは読みません。
- 外部通信はしません。
- ファイルやレジストリへの書き込みはしません。
- OneDrive の settings ファイルを `Open ... For Binary Access Read` で読み取ります。

参照する主な場所は次です。

```text
%LOCALAPPDATA%\Microsoft\OneDrive\settings\
```

主に次のファイルから、ローカル同期ルートと SharePoint / OneDrive URL ルートの対応を作ります。

```text
Business#\global.ini
Business#\<cid>.ini
Business#\ClientPolicy*.ini
Business#\<cid>.dat
Personal\global.ini
Personal\<cid>.ini
Personal\ClientPolicy*.ini
Personal\GroupFolders.ini
Personal\<cid>.dat
```

## 制約

このモジュールは、実務で使いやすいように安全側へ寄せた軽量実装です。`VBA-FileTools` や Guido Witt-Dörring 氏の Gist の考え方を参考にしていますが、完全移植ではありません。

特に、`SyncEngineDatabase.db` の解析は入れていません。OneDrive の環境によって `.dat` がなく `SyncEngineDatabase.db` のみになる場合、SharePoint のサブフォルダ同期や「OneDrive へのショートカットの追加」の一部は解決できない可能性があります。

失敗を許容できない業務処理では、変換結果が URL のままではないことと、対象パスが存在することを呼び出し側で確認してください。

```vb
Dim p As String
p = convURLtoLocalPath(ThisWorkbook, returnInputOnFail:=False)

If Len(p) = 0 Or LCase$(Left$(p, 8)) = "https://" Then
    Err.Raise vbObjectError + 1000, , "OneDrive同期ローカルパスを解決できません。"
End If

If Dir(p, vbDirectory) = "" Then
    Err.Raise vbObjectError + 1001, , "解決したローカルパスが存在しません。"
End If
```

## 参考

- Cristian Buse, `VBA-FileTools`
  - https://github.com/cristianbuse/VBA-FileTools
- Guido Witt-Dörring, `GetLocalOneDrivePath`
  - https://gist.github.com/guwidoe/038398b6be1b16c458365716a921814d
- Stack Overflow answer
  - https://stackoverflow.com/a/73577057/12287457

## ライセンス

MIT License。
