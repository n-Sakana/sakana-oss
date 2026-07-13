# toolrack

Windows の右クリックメニュー（Shift+右クリック）と、全アプリ共通のショートカットから起動する
社内ツールプラットフォーム。
PowerShell / Python のツールを「1フォルダ + tool.json」で追加できる。

マニフェスト規格の正本は [toolrack基盤設計.md](../tools/toolrack基盤設計.md)。

## インストール

1. `install.bat` をダブルクリック（管理者権限不要）
2. ファイル、フォルダ、またはフォルダ内の空白を **Shift+右クリック**
3. 一般ツールは `Tool Rack`、AIチャット連携ツールは `Tool Rack AI` を選ぶ

インストールすると、Explorerメニューに加えて非表示のTool Rack Hostが現在のユーザーで起動し、
次回サインイン時の自動起動もHKCUへ登録される。exeや管理者権限は使わず、Windows PowerShell 5.1が
起動時にリポジトリ内のC#を1回だけ読み込む。ツールを追加・削除したら`install.bat`を再実行する。
アンインストールは`uninstall.bat`で、Host、HKCU Run、Explorerメニュー、ローカルHost状態だけを削除し、
`bindings.json`、`tool`、`output`は削除しない。

## どこからでも起動

既定のグローバル操作は次の3つ。Explorer上に限らず、Chrome、VS Code、Office等でも同じように動く。

| 操作 | 動作 |
|---|---|
| `Ctrl`+右クリック | Captureをマウス付近に表示 |
| `Ctrl`+`Alt`+`C` | Captureを表示 |
| `Ctrl`+`Alt`+`T` | Transcribeを開始 |

修飾なしの右クリックは奪わない。`Ctrl`+右クリックは全アプリ共通で予約するため、その操作を使うアプリと
競合する場合は、ルートの`bindings.json`で該当bindingを変更または削除する。設定は保存後約500msで
自動再読込される。不正JSONや一時的な書込み途中は直前の正常設定を維持する。

```json
{
  "schema": 1,
  "bindings": [
    {
      "id": "capture-mouse",
      "trigger": { "type": "mouse", "button": "right", "modifiers": ["ctrl"] },
      "invoke": { "tool": "capture", "action": "default" }
    },
    {
      "id": "capture-hotkey",
      "trigger": { "type": "hotkey", "key": "C", "modifiers": ["ctrl", "alt"] },
      "invoke": { "tool": "capture", "action": "default" }
    }
  ]
}
```

v1のbindingはすべてグローバルで、`when`や`process`によるアプリ限定指定は受け付けない。
`RegisterHotKey`はシステム全体の登録なので、前面切替に合わせた登録解除では短時間のキー奪取や
取りこぼしを保証できないためである。Hostは前面アプリ、ウィンドウタイトル、キー履歴、マウス座標履歴、
入力内容を記録しない。

hotkeyがWindowsや他アプリと競合した場合、そのbindingだけをinactiveにして他の操作は継続する。
状態確認と手動再読込は次のコマンド。詳細ログは`%LOCALAPPDATA%\ToolRack\log\host.log`にある。

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\common\host-control.ps1 -Status
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\common\host-control.ps1 -Reload
```

グローバル経路はExplorerの選択対象を推測せず、ツールの`Target`へ常に空文字列を渡す。ファイルパスが
必要なツールはExplorerメニューから起動するか、ツール自身のpickerで選ぶ。Hostが起動できなくても
Explorerメニューは独立して残り、従来経路で各ツールを起動できる。

## 同梱ツール

| 分類 | ツール | 対象 | 内容 |
|---|---|---|---|
| Tool Rack | Capture | 空白 | 範囲／ウィンドウを選び、画像コピー／PNG保存＋パスコピー／ローカルOCRを即実行 |
| Tool Rack | Tree | フォルダ／空白 | フォルダツリーを md 化（深さ3／5／無制限／Custom）。完了後は `output` フォルダを Explorer で開く |
| Tool Rack | Timestamp | ファイル／フォルダ | 名前に `_yyyyMMdd` または `_yyyyMMddHHmmss` を付与。コンソールを表示せず実行 |
| Tool Rack | KeySend | 空白 | 無操作時にキー送信（スリープ防止）。Ctrl+C で停止 |
| Tool Rack | Timer | 空白 | WPF カウントダウンタイマー（5／10／25分／Custom） |
| Tool Rack AI | Clip | フォルダ／空白 | Codex／Claude向けの起動コマンドをクリップボードへコピー |
| Tool Rack AI | MD Mirror | ファイル／フォルダ／空白 | コード一式をAIチャット用のmdへ写す、またはmdから新しい `output` フォルダへ復元 |
| Tool Rack AI | MD Patch | フォルダ／空白 | AIが返した行番号付き差分を事前検査し、バックアップ後にフォルダへ適用 |
| Tool Rack AI | MD Extract | ファイル／フォルダ | テキスト、Office、PDFを従来互換のExtract Bundleへ抽出 |
| Tool Rack | VBA DevKit | ファイル／フォルダ | VBAを解析／抽出／差分／無害化／保護解除。Analyze／Extract／Diff／Sanitize／Unlockをメニューから選択 |
| Tool Rack | Transcribe | 空白 | 完全ローカルの日本語文字起こし。開始／停止、コピー／クリア／保存、感度、レベル表示に対応 |

## キャプチャ

`Ctrl`+右クリック、`Ctrl`+`Alt`+`C`、またはフォルダ内の空白をShift+右クリックして
`Tool Rack` → `Capture`を選ぶ。マウスの近くに一枚のパレットが表示され、`範囲`／`ウィンドウ`の
2列と`画像`／`パス`／`テキスト`の3行から、6通りの組合せをワンクリックで選ぶ。

- `画像`: 選択した画面を画像としてクリップボードへコピーする
- `パス`: PNGを`output/capture_<日時>.png`へ保存し、その絶対パスをコピーする
- `テキスト`: Windows標準OCRで画面内の文字を認識し、結果だけをコピーする

内側のカード枠や確認ダイアログはなく、light/darkテーマ、100～200% DPI、負座標を含む複数モニターへ
対応する。Tab、矢印、Enterでも選択でき、`Esc`またはパレット外のクリックで閉じる。
選択後の編集画面や保存確認は表示せず、処理成功時は短い通知だけを表示する。
画像とOCRデータは外部へ送信しない。OCRを利用できない場合はWindowsの言語OCR機能を追加する。

## ローカル文字起こし

フォルダ内の空白をShift+右クリックし、`Tool Rack` → `Transcribe` → `Start (Japanese)`を選ぶ。
黒いコンソールを表示せずWPF画面が開き、発話の切れ目ごとに日本語を追記する。句読点の自動付与は行わない。

- 音声認識、発話区切り、マイク処理はすべてPC内で行い、音声を外部へ送信しない
- 感度は0.5～4倍のソフトウェアゲイン。メーターは-60～0dBFSを表示範囲へ正規化
- 停止時は取り込み済みの音声を処理し、最後の発話を確定してから停止する
- コピーは全文をクリップボードへ、クリアは録音を止めずに現在の文字欄だけを空にする
- 保存はUTF-8 BOM付きで`output/transcribe_<日時>[_n].txt`へ出力する
- 音声そのものは保存しない

モデルとDLLを同梱するため、リポジトリは約180MB大きくなる。通常Gitの単体ファイル制限を避けるため、
最大のencoderモデルは2分割して管理している。Git LFSや追加ソフトは不要。初回起動時にツール内で結合し、
SHA-256が一致したモデルだけを使用する。2回目以降は検証済みの結合モデルを再利用する。

## AIチャット連携

### 既存プロジェクトを編集する

1. 対象フォルダを右クリックし、`MD Mirror` → `Create MD from File/Folder`
2. 生成された `output/md-mirror_<名前>_<日時>.md` をAIチャットへ渡す
3. AIが返した `MD Patch v1` 全文をクリップボードへコピー
4. 元のプロジェクトフォルダを右クリックし、`MD Patch` → `Apply from Clipboard`
5. 表示された変更対象を確認して適用

MD Mirrorに同梱される指示は、AIへ「変更したファイル全体」ではなく、1始まりの行番号、
変更前行、変更後行を持つMD Patchを返すよう要求する。コード行には差分記号や行番号を
付けず、宣言された行数で区画を判定するため、コードブロックや固定終了マーカーとの衝突はない。

MD Patchは、変更、新規テキスト／バイナリ、削除、空フォルダの作成／削除を扱う。
対象行とOLD内容、現在ファイルのSHA-256、文字コード、パス安全性、読み取り専用、ロック、
reparse point、パス長を全件先に検査し、1件でも不正なら何も書かない。既存ファイルの変更と
削除は `output/backup_<日時>/data/` に元のバイト列を保存してから適用する。

### 新しいプロジェクトを作る

1. `MD Mirror` → `Copy Format Instructions` で新規作成用の説明をクリップボードへコピー
2. 説明と依頼をAIチャットへ渡す
3. 返されたMD Mirror全文をコピー
4. `MD Mirror` → `Restore from Clipboard`

復元先は常に `output/md-mirror_restore_<日時>[_n]/`。MD内の `Source Name` は表示情報であり、
保存先の決定には使わない。元のプロジェクトへ直接上書きするのはMD Patchだけ。

### MD Mirrorの保存規則

- テキスト本文はUTF-8／LFへ統一し、元のEncoding、最初のEOL、末尾改行をヘッダへ記録
- CRLFとLFが混在するファイルは、最初のEOLへ統一して復元し、警告を表示
- 現在ファイル用のOriginal SHA-256と、正規化後のRestore SHA-256を別々に記録
- 画像、フォント、WASM等のアプリ資産はBase64で格納し、サイズとSHA-256を検証
- バイナリ上限は1ファイル5 MiB、合計10 MiB。動画、音声、実行ファイル等と上限超過分は目録のみ
- `.git`、`node_modules`、`.venv`、`__pycache__` とreparse pointは走査しない

形式の大小文字、ヘッダ順、必須項目は厳格。AIの説明文や表記ゆれを推測して適用することはない。

### 文書をAIへ読ませる

`MD Extract`は旧 `extract-md.ps1` の機能とExtract Bundle形式を維持している。DOCX／XLSX／PPTXは
OfficeなしでOpenXMLを直接読み、DOC／XLS／PPT／PDFはOfficeがある場合に隔離ワーカーで処理する。
既定timeoutは120秒、`Timeout 300`は300秒、`Custom...`はtimeout、再試行timeout、Deferred時の
動作を対話設定する。timeoutしたファイルはbundleの `Deferred Files` に残り、他の抽出を止めない。

出力名は従来どおり `output/extract-md_<対象名>_<日時>.md`、ログは `log/extract-md_<日時>.log`。

OpenXML経路は未信頼文書を隔離ワーカー内で扱い、DTD／外部エンティティ、重複part名、
10,000件を超えるZIP entry、1 part 64 MiB／XML合計256 MiB超を拒否する。XLSXの列参照は
Excel上限のXFDまでとし、細工された疎な列番号による過大メモリ確保を行わない。
Office COM経路は生成したApplicationのPID・プロセス名・開始時刻を記録し、完全一致する
隔離プロセスだけを終了する。同時に起動された既存のOfficeプロセスは終了対象にしない。

## ツールを作る・持ち込む

> VBA DevKitのSanitize／Unlock／Analyzeは、正当な権限を持つファイルへの防御的調査用です。
> 不審なファイルは隔離環境で扱ってください。HTMLレポートは元コードを含むため、外部共有に注意してください。

`tool/<id>/` にフォルダを作り、`tool.json` と実行ファイルを置く。

```json
{
  "schema": 1,
  "id": "mytool",
  "name": "My Tool",
  "on": ["folder"],
  "run": { "type": "powershell", "entry": "main.ps1" }
}
```

`run.window` は次の3モード。

| 値 | 動作 |
|---|---|
| `console` | 既定。コンソール内で同期実行し、既定では終了後にキー待ち |
| `hidden` | コンソールを表示せず、切り離して実行。対話不要の短いツール向け |
| `gui` | コンソールを表示せず、切り離して GUI アプリを実行 |

`hidden` / `gui` の起動前エラーは MessageBox で表示される。既存アプリは規定の
`-Target`（PowerShell）または `--target`（Python）を受けるアダプタで載せられる。

`hidden` は対話入力のないツール向け。ツール起動後の標準出力は表示されず、ランチャーも
終了を待たないため、実行時エラーはツール自身が MessageBox 等で通知する必要がある。
PowerShell ツールでは `common/ui.ps1` の `Show-ErrorDialog` を利用できる。

## 右クリックメニューの分類

ツールの保存場所はすべて `tool/<id>/` のままにし、表示分類とページへの詰め込み順をルートの
`menu.json` で指定する。Explorer上の最終的な表示順はWindowsがレジストリキー名から決める。

```json
{
  "schema": 1,
  "default_category": "general",
  "categories": [
    { "id": "general", "label": "Tool Rack", "tools": ["timer", "tree"] },
    { "id": "ai", "label": "Tool Rack AI", "tools": ["md-mirror", "md-patch"] }
  ]
}
```

Windows 11の静的カスケードメニューには16項目の上限がある。toolrackは、直接実行するツールを
1枠、選択肢を持つツールを「ツール見出し1枠＋選択肢数」と数え、分類ごと・クリック対象ごとに
16枠以内へ自動分割する。続きは `Tool Rack AI 2` のような別メニューになり、1つのツールを途中で
分割することはない。単独で16枠を超えるツールは理由を表示して登録しない。

`menu.json` にない正常なツールは警告付きで `default_category` の末尾へ入る。設定に書かれた
未導入のツールIDは警告して無視する。設定ファイル自体が不正な場合、インストールは失敗し、
現在登録済みのメニューには触れない。`menu.json` が存在しない場合は、全ツールを `Tool Rack` に
まとめた設定として扱い、必要なら同じ規則で自動分割する。

## テスト

PowerShell 5.1で全テストを実行する。

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\test\run-tests.ps1
```

完了時は `ALL TEST FILES PASSED`。テストには実クリップボード、実HKCUレジストリ、OpenXML fixture、
Officeが利用可能な場合のCOM経路が含まれる。レジストリを使うテストの終了後は `install.bat` または
`common\install.ps1` を再実行してメニューを復元する。

3つのAIチャット連携ツールを右クリックメニューから手動確認する場合は、次のコマンドで
`output\manual-test-fixtures` に反復利用できるテスト一式を作成する。

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\test\create-manual-fixtures.ps1 -IncludeLegacyOffice
```

正常例と拒否例、各文字コード、改行、画像、OpenXML、PDF、利用可能な旧Office形式を含む。
使い方と期待結果は [手動確認用ファイル](test/MANUAL-FIXTURES.md) を参照。生成スクリプトを
再実行すると、MD Patch適用後の対象を含めて初期状態へ戻る。
