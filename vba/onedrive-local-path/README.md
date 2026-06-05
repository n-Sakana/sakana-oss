# OneDrive / SharePoint URL to Local Path for VBA

Excel VBA で `ThisWorkbook.Path` が SharePoint / OneDrive の URL を返す場合に、OneDrive 同期済みのローカルパスへ変換するための VBA モジュール群です。


## なぜ必要か

Excel VBA の `Workbook.Path` / `Workbook.FullName` は、ブックを基準に関連ファイルを読む処理でよく使われます。たとえば、次のような処理です。

```vb
baseDir = ThisWorkbook.Path
Open baseDir & "\settings.ini" For Input As #1
```

ローカル保存のブックであれば、`ThisWorkbook.Path` は通常 `C:\Users\...` のようなローカルファイルシステム上のパスになります。その前提で、`Dir`、`MkDir`、`Open`、`FileSystemObject` などを使う既存マクロは多くあります。

一方、ブックを OneDrive / SharePoint / Teams 配下に置くと、Excel は `ThisWorkbook.Path` や `ThisWorkbook.FullName` に `https://...sharepoint.com/...` 形式の URL を返す場合があります。Microsoft の `Workbook.Path` の説明自体は「ブックを表す完全なパス」としていますが、クラウド上のブックでは、その「パス」がローカル同期フォルダではなく SharePoint / OneDrive の URL になることがあります。

この URL は、Excel が共同編集や自動保存を扱う上では意味があります。しかし、VBA の通常のファイル操作から見ると、これはローカルフォルダではありません。`Dir` や `FileSystemObject` に `https://...` を渡しても、ブックと同じフォルダにある設定ファイル、CSV、出力フォルダ等をそのまま扱うことはできません。

SharePoint ライブラリや Teams のファイルは OneDrive 同期アプリによってエクスプローラー上のフォルダとして扱えますが、URL とローカル同期先の対応は単純な文字列置換では決まりません。ユーザー名、組織名、個人用 OneDrive / 職場または学校アカウント、複数アカウント、SharePoint ライブラリ、OneDrive へのショートカット追加などでローカル側の配置が変わるためです。

このモジュールは、その差を埋めるためのものです。Excel から見えている SharePoint / OneDrive URL を、OneDrive クライアントが同期しているローカルパスへ変換し、既存の「ブックと同じフォルダを基準にファイルを読む・書く」VBA を動かしやすくします。

このモジュールが必要になる典型例は次の条件が重なる場合です。

- `.xlsm` ブックを OneDrive / SharePoint / Teams 配下で使う。
- その場所を OneDrive 同期アプリでローカルPCにも同期している。
- マクロが `ThisWorkbook.Path` / `ThisWorkbook.FullName` を基準に、同じフォルダや下位フォルダのファイルを操作する。
- Microsoft Graph API や SharePoint API へ作り替えるのではなく、既存のローカルファイル前提の VBA をできるだけ維持したい。

逆に、ブックがローカルディスクや通常のネットワークドライブ上にある場合、またはマクロ側が最初からクラウドURL/APIを扱う設計になっている場合、この変換は不要です。OneDrive 同期されていない SharePoint 上のファイルを、魔法のようにローカルファイル化するものでもありません。紅茶で言えば、茶葉を増やす道具ではなく、棚番を読み替える札です。


## OneDrive 同期方式の前提

SharePoint / Teams のファイルをエクスプローラーで扱う方法は、大きく分けて2種類あります。Microsoft の説明でも、SharePoint ライブラリや Teams のファイルをローカルPCで扱う方法として、`Sync` ボタンによる同期と、`Add shortcut to My files` による OneDrive ショートカット追加が分けられています。

どちらも最終的にはエクスプローラーや Finder にフォルダとして見えるため、利用者からは同じ「同期」に見えます。しかし、`ThisWorkbook.Path` が返す SharePoint / OneDrive URL をローカルパスへ戻すマクロの文脈では、両者は別物です。

### 1. SharePoint / Teams の「同期」

SharePoint サイトまたは Teams のファイル画面で、対象のドキュメントライブラリやフォルダを開き、`同期` / `Sync` を押して OneDrive 同期アプリに登録する方式です。設定はそのPCごとに行います。別PCでも使う場合は、そのPCでも同じ同期操作が必要です。

ローカルでは、典型的には次のような場所に見えます。

```text
C:\Users\<user>\<organization>\<site> - <library>\...
```

この方式では、OneDrive 同期アプリの設定画面から確認しやすいです。

```text
OneDrive 雲アイコン
  → 設定
  → アカウント
  → 同期中の SharePoint サイト / ライブラリ一覧
```

ここに対象のサイトやライブラリが表示され、`フォルダーの選択` や `同期の停止` の対象になっていれば、SharePoint / Teams の直接同期です。

マクロ上の注意点は、ブックがローカルに同期されていても、Excel の `ThisWorkbook.Path` / `ThisWorkbook.FullName` が次のようなクラウドURLを返す場合があることです。

```text
https://tenant.sharepoint.com/sites/site/Shared%20Documents/...
```

この場合、マクロでは「SharePoint サイト / ライブラリのURLルート」と「そのPC上の同期ルート」を対応させる必要があります。URLの末尾だけを `C:\Users\...` に置換する処理では、サイト名、ライブラリ名、表示名、エンコード差で簡単に壊れます。

### 2. OneDrive に「ショートカットを追加」

SharePoint / Teams / 共有フォルダ側で、対象フォルダに対して `OneDrive にショートカットを追加` / `Add shortcut to My files` を選ぶ方式です。追加されたフォルダは OneDrive Web の `自分のファイル` / `My files` に表示され、OneDrive 同期アプリを使っているPCでは個人の OneDrive 配下にも現れます。

ローカルでは、典型的には次のような場所に見えます。

```text
C:\Users\<user>\OneDrive - <organization>\<shortcut name>\...
```

この方式の簡単な見分け方は、OneDrive Web です。

```text
OneDrive Web
  → 自分のファイル / My files
  → 対象フォルダを選択
  → 「ショートカットの削除」/ "Remove shortcut" が出る
```

この表示なら、直接同期ではなく OneDrive ショートカットです。Microsoft のトラブルシュートでも、ショートカットは OneDrive サイトの `My files` で確認し、削除できるものとして扱われています。

この方式では、元の SharePoint / Teams 上のフォルダ名と、ローカル上のフォルダ名が一致するとは限りません。追加したショートカットはユーザー側で名前を変更でき、その変更名は本人の OneDrive 上でだけ見えます。ショートカットを OneDrive 内の別フォルダへ移動できる場合もあります。

そのため、Excel が返すURLは元の SharePoint / Teams の場所を指しているのに、ローカル実体は個人 OneDrive 配下のショートカット名で見えている、という状態になります。

```text
ThisWorkbook.Path:
https://tenant.sharepoint.com/sites/site/Shared%20Documents/案件A/...

ローカル側:
C:\Users\<user>\OneDrive - <organization>\審査資料\...
```

この2つは文字列として似ている保証がありません。特にショートカット方式では、URLからローカルパスを単純復元する発想は危険です。このモジュールでは、OneDrive 設定ファイル内の `AddedScope` 等からショートカット側の対応も拾う方針にしています。

### 簡単な見分け方

実務では、先にUIで判定するのが確実です。ローカルパスの形から推測するのは最後で構いません。

| 確認場所 | 見え方 | 判断 |
| --- | --- | --- |
| OneDrive Web の `自分のファイル` / `My files` | 対象フォルダに `ショートカットの削除` / `Remove shortcut` が出る | OneDrive ショートカット追加 |
| OneDrive 同期アプリの `設定` → `アカウント` | 対象の SharePoint サイト / ライブラリが同期対象として表示される | SharePoint / Teams の直接同期 |
| エクスプローラー | `C:\Users\<user>\OneDrive - <org>\...` 配下に対象がある | ショートカット追加の可能性が高い |
| エクスプローラー | `C:\Users\<user>\<org>\<site> - <library>\...` 配下に対象がある | 直接同期の可能性が高い |

注意点として、同じ SharePoint ドキュメントライブラリに対して、`Sync` と `Add shortcut to My files` を同時に使うことはできません。片方で登録済みの対象をもう片方でも登録しようとすると、OneDrive 同期アプリ側でエラーになることがあります。

マクロの利用者に確認する場合は、次の順で聞くと混乱が少なくなります。

1. OneDrive Web の `自分のファイル` に対象フォルダがあり、`ショートカットの削除` が出るか。
2. OneDrive 同期アプリの `アカウント` タブに、対象の SharePoint サイト / ライブラリが出るか。
3. エクスプローラー上のローカルパスが、個人 OneDrive 配下か、組織名 / サイト名配下か。

この判定をしないまま `ThisWorkbook.Path` のURLを処理すると、直接同期向けの対応表でショートカットを探したり、逆にショートカット名を SharePoint ライブラリ名だと誤認したりします。見た目は同じフォルダでも、茶葉の棚番が違います。

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

- Microsoft Learn, `Workbook.Path` property
  - https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.path
- Microsoft Learn, `Workbook.FullName` property
  - https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.fullname
- Microsoft Learn, How AutoSave impacts add-ins and macros
  - https://learn.microsoft.com/en-us/office/vba/library-reference/concepts/how-autosave-impacts-addins-and-macros
- Microsoft Learn, Sync in SharePoint and OneDrive
  - https://learn.microsoft.com/en-us/sharepoint/sharepoint-sync
- Microsoft Support, Sync SharePoint and Teams files with your computer
  - https://support.microsoft.com/en-gb/office/sync-sharepoint-and-teams-files-with-your-computer-6de9ede8-5b6e-4503-80b2-6190f3354a88
- Microsoft Support, Add shortcuts to shared folders in OneDrive
  - https://support.microsoft.com/en-us/office/add-shortcuts-to-shared-folders-in-onedrive-d66b1347-99b7-4470-9360-ffc048d35a33
- Microsoft Learn, You can't sync a shared folder that's set as a shortcut in OneDrive sync app
  - https://learn.microsoft.com/en-us/troubleshoot/sharepoint/sync/cannot-sync-shortcut-folder
- Microsoft Tech Community, ThisWorkbook.FullName / local path discussion
  - https://techcommunity.microsoft.com/t5/excel/thisworkbook-fullname-determine-local-path-with-vba/td-p/3816000
- Microsoft Q&A, `ThisWorkbook.Path` / `ThisWorkbook.FullName` returns URL discussion
  - https://learn.microsoft.com/en-us/answers/questions/5271931/the-onedrive-nightmare-continues-thisworkbook-path
- Cristian Buse, `VBA-FileTools`
  - https://github.com/cristianbuse/VBA-FileTools
- Guido Witt-Dörring, `GetLocalOneDrivePath`
  - https://gist.github.com/guwidoe/038398b6be1b16c458365716a921814d
- Stack Overflow answer
  - https://stackoverflow.com/a/73577057/12287457

## ライセンス

MIT License。
