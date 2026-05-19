# OneDrive / SharePoint URL to Local Path for VBA

Excel VBA で `ThisWorkbook.Path` が SharePoint / OneDrive の URL を返す場合に、OneDrive 同期済みのローカルパスへ変換するための VBA モジュール群です。

## モジュール構成

通常業務では本番用だけを使います。

| ファイル | 用途 |
| --- | --- |
| `ConvURLtoLocalPath.bas` | 本番用。DBなしの軽量実装です。 |
| `ConvURLtoLocalPath_Debug.bas` | 検証用。本番用モジュールと一緒にインポートしてログを出します。 |
| `ConvURLtoLocalPath_WithDb.bas` | 旧経路・重い検証用。`SyncEngineDatabase.db` も読む実装を隔離しています。 |
| `Demo_OneDrivePathIssue.bas` | デモ用。FSO/Dirでファイル作成まで比較します。 |

## 使い方

`ConvURLtoLocalPath.bas` を VBA プロジェクトへインポートします。通常は `convURLtoLocalPath` を使います。

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

`convURLtoLocalPath` は `SyncEngineDatabase.db` を読みません。明示的に同じ挙動を呼びたい場合は、別名の軽量版も使えます。

```vb
folderPath = convURLtoLocalPathLight(ThisWorkbook, returnInputOnFail:=False)
```

`SyncEngineDatabase.db` も読む旧経路を使いたい場合は、追加で `ConvURLtoLocalPath_WithDb.bas` をインポートし、明示的に次を使います。

```vb
folderPath = convURLtoLocalPathWithDb(ThisWorkbook, returnInputOnFail:=False)
```

## 診断ログ

変換できない場合や、成功/失敗の条件差を調べたい場合は、`ConvURLtoLocalPath.bas` と `ConvURLtoLocalPath_Debug.bas` をインポートして次を実行します。

```vb
DebugThisWorkbookLocalPath
```

これは `Alt + F8` のマクロ一覧から実行できる引数なしの診断マクロです。Immediate ウィンドウと `OneDrive Path Debug` シートに、次の情報を出力します。

- `ThisWorkbook.Path` / `ThisWorkbook.FullName`
- `%LOCALAPPDATA%\Microsoft\OneDrive\settings\` の検出結果
- `Business#` / `Personal` アカウントフォルダの検出結果
- 生成された `webRoot -> localRoot` 対応表
- 入力URLに一致した `webRoot`、候補ローカルパス、存在確認結果

シート出力が不要な場合は次のようにします。

```vb
convURLtoLocalPathDebug ThisWorkbook, False
```

`convURLtoLocalPathDebug` は引数付きのため、Excel のマクロ一覧には表示されない場合があります。

DBも読む診断が必要な場合は、追加で `ConvURLtoLocalPath_WithDb.bas` をインポートし、`Alt + F8` から次を実行します。

```vb
DebugThisWorkbookLocalPathWithDb
```

ショートカット追加された場所を調べる場合は、検証用モジュールの次のマクロを実行します。

```vb
DebugOneDriveAddedScopes
```

`OneDrive AddedScope Debug` シートに、`AddedScope` 行の検出元、URL、folderId、相対パス、`OneDriveCommercial` 等から作った候補パスと存在確認結果を出します。

## デモ

`Demo_OneDrivePathIssue.bas` は報告・説明用のデモモジュールです。`ConvURLtoLocalPath.bas` と一緒にインポートして、次のマクロを実行します。

```vb
Demo_PathIssue_Compare
```

実行すると `Path Demo` シートを作成し、次の2ケースを同じ表で比較します。

```text
Raw ThisWorkbook.Path
convURLtoLocalPath(ThisWorkbook)
```

各ケースで、ブックと同じ場所に `_vba_path_demo\path-demo.txt` を作成しようとします。表には作成先フォルダ、作成先ファイル、`FolderExists`、`CreateFolder`、`CreateTextFile`、`Dir(file)` の結果を出します。SharePoint / OneDrive URL として開かれている場合、通常の `ThisWorkbook.Path` 側は FSO / Dir で失敗し、`convURLtoLocalPath(ThisWorkbook)` 側はローカル同期パスへ解決できれば成功します。

動画や画面共有では、`ThisWorkbook.Path` が `https://...` になっていること、FSO処理の失敗、変換後パスでの成功を同じシート上で見せられます。

## 方針

- Win32 API `Declare` は使いません。
- `Shell` / `WScript.Shell` / `Shell.Application` / PowerShell は使いません。
- レジストリは読みません。
- 外部通信はしません。
- 本番用・検証用・DBありモジュールは外部ファイルやレジストリへの書き込みをしません。検証用はブック内に `OneDrive Path Debug` シートを書きます。デモモジュールだけは動作確認のため `_vba_path_demo\path-demo.txt` を作成します。
- OneDrive の settings ファイルを `Open ... For Binary Access Read` で読み取ります。
- `global.ini` の `cid` が空の場合は、同じアカウントフォルダ内の `.ini` を調べ、`libraryScope` 等を含むGUID形式の設定ファイルを `<cid>.ini` として扱います。
- `<cid>.ini` の `libraryScope` / `AddedScope` にURLが含まれている場合は、その固定位置のURLを `ClientPolicy*.ini` より優先して使います。

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

`SyncEngineDatabase.db` は本番用では読みません。`ConvURLtoLocalPath_WithDb.bas` を使った場合だけ参照します。

## 制約

本番用モジュールは、実務で使いやすいように安全側へ寄せた軽量実装です。`VBA-FileTools` や Guido Witt-Dörring 氏の Gist の考え方を参考にしていますが、完全移植ではありません。

`convURLtoLocalPath` / `convURLtoLocalPathLight` / `DebugThisWorkbookLocalPath` / `DebugThisWorkbookLocalPathLight` は `SyncEngineDatabase.db` を読みません。DBも読む場合は `ConvURLtoLocalPath_WithDb.bas` の `convURLtoLocalPathWithDb` または `DebugThisWorkbookLocalPathWithDb` を使います。SQLite API や外部DLLは使わず、VBAのバイナリ読み取りだけで必要なフォルダID、親ID、フォルダ名を抽出します。

OneDrive クライアントの設定ファイル形式は更新される可能性があります。失敗を許容できない業務処理では、変換結果が URL のままではないことと、対象パスが存在することを呼び出し側で確認してください。

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
