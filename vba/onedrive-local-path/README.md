# OneDrive / SharePoint URL to Local Path for VBA

Excel VBA で `ThisWorkbook.Path` が SharePoint / OneDrive の URL を返す場合に、OneDrive 同期済みのローカルパスへ変換するための単一モジュールです。

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

`convURLtoLocalPath` は既定で `SyncEngineDatabase.db` を読みません。明示的に同じ挙動を呼びたい場合は、別名の軽量版も使えます。

```vb
folderPath = convURLtoLocalPathLight(ThisWorkbook, returnInputOnFail:=False)
```

`SyncEngineDatabase.db` も読む旧経路を使いたい場合は、明示的に次を使います。

```vb
folderPath = convURLtoLocalPathWithDb(ThisWorkbook, returnInputOnFail:=False)
```

## 診断ログ

変換できない場合や、成功/失敗の条件差を調べたい場合は次を実行します。

```vb
DebugThisWorkbookLocalPath
```

これは `Alt + F8` のマクロ一覧から実行できる引数なしの診断マクロです。Immediate ウィンドウと `OneDrive Path Debug` シートに、次の情報を出力します。

- `ThisWorkbook.Path` / `ThisWorkbook.FullName`
- `%LOCALAPPDATA%\Microsoft\OneDrive\settings\` の検出結果
- `Business#` / `Personal` フォルダごとの `cid`
- `ClientPolicy*.ini` の `DavUrlNamespace` / `SiteID` / `WebID` / `IrmLibraryId`
- `<cid>.dat` と `SyncEngineDatabase.db` から取得できたフォルダ件数
- 生成された `webRoot -> localRoot` 対応表
- 入力URLに一致した `webRoot`、候補ローカルパス、存在確認結果

シート出力が不要な場合は次のようにします。

```vb
convURLtoLocalPathDebug ThisWorkbook, False
```

`convURLtoLocalPathDebug` は引数付きのため、Excel のマクロ一覧には表示されない場合があります。

`DebugThisWorkbookLocalPath` は既定で `SyncEngineDatabase.db` を読みません。DBも読む診断が必要な場合は、`Alt + F8` から次を実行します。

```vb
DebugThisWorkbookLocalPathWithDb
```

## デモ

`Demo_OneDrivePathIssue.bas` は報告・説明用のデモモジュールです。`ConvURLtoLocalPath.bas` と一緒にインポートして、次のマクロを実行します。

```vb
Demo_PathIssue_Compare
```

実行すると `Path Demo` シートを作成し、次の2ケースを横並びで比較します。

```text
Raw ThisWorkbook.Path
convURLtoLocalPath(ThisWorkbook)
```

各ケースで、ブックと同じ場所に `_vba_path_demo\path-demo.txt` を作成しようとします。SharePoint / OneDrive URL として開かれている場合、通常の `ThisWorkbook.Path` 側は FSO / Dir で失敗し、`convURLtoLocalPath(ThisWorkbook)` 側はローカル同期パスへ解決できれば成功します。

動画や画面共有では、`ThisWorkbook.Path` が `https://...` になっていること、FSO処理の失敗、変換後パスでの成功を同じシート上で見せられます。

## 方針

- Win32 API `Declare` は使いません。
- `Shell` / `WScript.Shell` / `Shell.Application` / PowerShell は使いません。
- レジストリは読みません。
- 外部通信はしません。
- ファイルやレジストリへの書き込みはしません。
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
Business#\SyncEngineDatabase.db
Personal\global.ini
Personal\<cid>.ini
Personal\ClientPolicy*.ini
Personal\GroupFolders.ini
Personal\<cid>.dat
Personal\SyncEngineDatabase.db
```

## 制約

このモジュールは、実務で使いやすいように安全側へ寄せた軽量実装です。`VBA-FileTools` や Guido Witt-Dörring 氏の Gist の考え方を参考にしていますが、完全移植ではありません。

`.dat` からフォルダ対応を取得できない場合は、`SyncEngineDatabase.db` も読み取ります。SQLite API や外部DLLは使わず、VBAのバイナリ読み取りだけで必要なフォルダID、親ID、フォルダ名を抽出します。

`convURLtoLocalPath` / `convURLtoLocalPathLight` / `DebugThisWorkbookLocalPath` / `DebugThisWorkbookLocalPathLight` は `SyncEngineDatabase.db` を読みません。DBも読む場合は `convURLtoLocalPathWithDb` または `DebugThisWorkbookLocalPathWithDb` を使います。

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
