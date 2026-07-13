# AIチャット連携3ツール 手動テスト手順

この文書は、MD Mirror、MD Patch、MD Extractを、`output\manual-test-fixtures`のテストデータで確認するための手順書です。

テストデータは次の場所に生成されます。

`C:\repos\pub\toolrack\output\manual-test-fixtures`

## 1. 3ツールの役割

| ツール | 入力 | 結果 | 入力元を変更するか |
|---|---|---|---|
| MD Mirror | `mirror-source` | 1個のMDを作成、またはMDから別フォルダへ復元 | 変更しない |
| MD Patch | `patch-target`とPatch MD | `patch-target`を書き換える | 変更する。既存ファイルは事前にバックアップ |
| MD Extract | `extract-source` | 文書内容をまとめた1個のMD | 変更しない |

3ツールは別々のツールです。MD Mirror用の`ready-mirror.md`をMD Patchへ渡したり、MD Patch用の`apply-demo.md`をMD Mirrorへ渡したりしないでください。

## 2. テストを始める

テストデータがすでにある場合、生成コマンドを実行する必要はありません。

1. エクスプローラーで次を開きます。

   `C:\repos\pub\toolrack\output\manual-test-fixtures`

2. 次のフォルダとファイルがあることを確認します。

   - `mirror-source`：MD Mirror用の入力
   - `ready-mirror.md`：MD Mirrorの復元用
   - `patch-target`：MD Patchの書き換え対象
   - `patches`：MD Patchへ渡すMD
   - `extract-source`：MD Extract用の入力

3. Tool Rackのメニューを出すときは、対象を選択してからShiftを押しながら右クリックします。
4. `Tool Rack`が見えない場合は「その他のオプションを確認」を開きます。

## 3. MD Mirror：フォルダを1個のMDにする

1. `mirror-source`フォルダ自体をShift＋右クリックします。フォルダ内のファイルではありません。
2. `Tool Rack`を選びます。
3. `MD Mirror`を選びます。
4. `Create MD from File/Folder`を選びます。
5. 処理が終わると、toolrackの`output`フォルダがエクスプローラーで開きます。
6. `md-mirror_mirror-source_日時.md`というファイルが新しく作られたことを確認します。

元の`mirror-source`は変更されません。作成されるMDには次のテスト項目が入っています。

- UTF-8
- BOM付きUTF-8
- UTF-16LE
- CP932
- CRLFとLF
- 末尾改行なし
- 空ファイル
- 空フォルダ
- 深いフォルダ
- 日本語ファイル名
- PNG画像
- 区切り記号に見える文章

## 4. MD Mirror：MDファイルから復元する

1. `manual-test-fixtures`へ戻ります。
2. `ready-mirror.md`をShift＋右クリックします。
3. `Tool Rack`を選びます。
4. `MD Mirror`を選びます。
5. `Restore from MD File`を選びます。
6. `output\md-mirror_restore_日時`という新しい復元先が作成され、そのフォルダがエクスプローラーで開くことを確認します。

成功時は12ファイルと8フォルダが復元されます。次を確認します。

- `日本語.txt`がある
- `empty.txt`が0バイト
- 空の`empty-dir`がある
- `assets\logo.png`が画像として開ける
- `deep\a\b\c\deep.txt`がある
- `marker-looking-text.txt`の文章が欠けていない
- `docs\cp932.txt`などが文字化けしていない

`mixed-eol.txt`だけは、混在していた改行が最初に現れるCRLFへ統一されます。それ以外は元のバイト列と一致することを、生成時の自動検証で確認しています。

復元先は必ず新しい`output`フォルダです。`mirror-source`へは上書きしません。

## 5. MD Mirror：クリップボードから復元する

1. `ready-mirror.md`をメモ帳などで開きます。
2. Ctrl＋Aで全文を選択します。
3. Ctrl＋Cでコピーします。
4. エクスプローラーのフォルダ内で、空白部分をShift＋右クリックします。
5. `Tool Rack`を選びます。
6. `MD Mirror`を選びます。
7. `Restore from Clipboard`を選びます。
8. 新しい`output\md-mirror_restore_日時`が作られることを確認します。

先頭から末尾まで、MD全文をコピーする必要があります。

## 6. MD Mirror：新規作成用の説明をコピーする

1. エクスプローラーのフォルダ内で、空白部分をShift＋右クリックします。
2. `Tool Rack`を選びます。
3. `MD Mirror`を選びます。
4. `Copy Format Instructions`を選びます。
5. メモ帳へ貼り付け、AI向けの形式説明がクリップボードへ入ったことを確認します。

実運用では、この説明と作成依頼をAIへ渡し、AIが返したMD全文を`Restore from Clipboard`で復元します。

## 7. MD Patch：拒否テスト

MD Patchは`patch-target`を実際に変更します。最初に、何も変更しない3つの拒否テストを実行します。

共通操作は次のとおりです。

1. `patch-target`フォルダ自体をShift＋右クリックします。
2. `Tool Rack`を選びます。
3. `MD Patch`を選びます。
4. `Apply from MD File`を選びます。
5. `Patch MD file:`と表示されたら、使用するMDのフルパスを貼り付けます。
6. Enterを押します。

### 7.1 フォルダ外への書き込みを拒否する

入力するパス：

`C:\repos\pub\toolrack\output\manual-test-fixtures\patches\reject-path-attack.md`

期待結果：

- パスが不正というエラーになる
- `patch-target`は一切変わらない
- `manual-test-fixtures`の外にもファイルは作られない
- バックアップは作られない

### 7.2 SHA-256が違う場合に拒否する

入力するパス：

`C:\repos\pub\toolrack\output\manual-test-fixtures\patches\reject-hash-mismatch.md`

期待結果：

- `Expected SHA-256 mismatch: obsolete.txt`と表示される
- `obsolete.txt`は削除されない
- ほかのファイルも変更されない
- バックアップは作られない

### 7.3 変更前の行内容が違う場合に拒否する

入力するパス：

`C:\repos\pub\toolrack\output\manual-test-fixtures\patches\reject-old-content-mismatch.md`

期待結果：

- `OLD content mismatch in src/app.ps1 at line 2`と表示される
- `src\app.ps1`は変更されない
- ほかのファイルも変更されない
- バックアップは作られない

3つとも書き込み前に止まるため、初期化せず続けて実行できます。

## 8. MD Patch：正常なMDファイルを適用する

1. `patch-target`をShift＋右クリックします。
2. `Tool Rack`を選びます。
3. `MD Patch`を選びます。
4. `Apply from MD File`を選びます。
5. 次のパスを入力します。

   `C:\repos\pub\toolrack\output\manual-test-fixtures\patches\apply-demo.md`

6. 変更予定の5操作が表示されることを確認します。
7. `Apply 5 operation(s)? [y/N]`と聞かれたら、`y`を入力してEnterを押します。
8. 適用後、作成されたバックアップフォルダがエクスプローラーで開くことを確認します。

`patch-target`には次の変更が入ります。

- `src\app.ps1`の`$name = "world"`が`$name = "toolrack"`になる
- `created.txt`が新しくできる
- `obsolete.txt`が削除される
- `empty-created`という空フォルダができる
- `assets\logo.png`が別の画像へ置き換わる

バックアップは`output\backup_日時\data`に作られます。次の変更前ファイルがあることを確認します。

- `data\src\app.ps1`
- `data\obsolete.txt`
- `data\assets\logo.png`

新規作成された`created.txt`と`empty-created`には変更前の実体がないため、バックアップファイルはありません。

## 9. MD Patch：クリップボードから適用する

先に「13. テストデータを初期状態へ戻す」を実行してください。正常適用後の対象へ同じPatchをもう一度適用すると、安全検査によって拒否されます。

1. `patches\apply-demo.md`をメモ帳で開きます。
2. Ctrl＋Aで全文を選択します。
3. Ctrl＋Cでコピーします。
4. `patch-target`をShift＋右クリックします。
5. `Tool Rack`を選びます。
6. `MD Patch`を選びます。
7. `Apply from Clipboard`を選びます。
8. 変更予定の5操作が表示されることを確認します。
9. `y`を入力して適用します。

結果は`Apply from MD File`の場合と同じです。

## 10. MD Extract：一式を抽出する

1. `extract-source`フォルダ自体をShift＋右クリックします。
2. `Tool Rack`を選びます。
3. `MD Extract`を選びます。
4. `Default`を選びます。
5. ファイルの走査と抽出が始まることを確認します。

現在のテストデータでは、除外対象の`.git`内を除いた18ファイルが対象です。次のデータが含まれます。

- UTF-8、BOM付きUTF-8、UTF-16LE、CP932
- 空ファイル
- 深いフォルダのテキスト
- Pythonコード
- DOCX、XLSX、PPTXの本文とノート
- 旧形式のDOC、XLS、PPT
- PDF
- PNG画像
- 未対応のBINファイル

抽出が終わると、`output\extract-md_extract-source_日時.md`が作られます。VS Codeの`code`コマンドが利用可能ならVS Codeで、なければメモ帳で開きます。

出力MD内で次の文字を検索します。

- `DOCX MANUAL FIXTURE`
- `XLSX MANUAL FIXTURE`
- `PPTX MANUAL FIXTURE`
- `NOTES MANUAL FIXTURE`
- `WORD LEGACY MANUAL FIXTURE`
- `EXCEL LEGACY MANUAL FIXTURE`
- `POWERPOINT LEGACY MANUAL FIXTURE`

次も確認します。

- `.git\hidden.txt`の内容である`this must be excluded`が出力されていない
- `screenshot.png`と`unsupported.bin`がUnsupportedとして記録される
- 通常のテキストが文字化けしていない
- `Failed Files`に予期しない失敗がない

画像と未対応BINがUnsupportedになるのは想定どおりで、ツール全体の失敗ではありません。

## 11. MD Extract：PDFがDeferredになった場合

PDFはWordを使って読み込むため、環境によって時間がかかります。既定の待ち時間は1ファイル120秒です。

時間内に完了しなかった場合、次の質問が表示されます。

`Retry deferred files now? [R] Retry / [S] Skip (default: S):`

- `R`：Deferredになったファイルを300秒で再試行する
- `S`またはEnter：再試行せず、Deferredとして出力MDへ残す

Deferredになっても、それ以前に抽出できたほかのファイルは出力MDへ保存されます。

## 12. MD Extract：Timeout 300とCustom

### Timeout 300

`MD Extract`の`Timeout 300`を選ぶと、最初から各Office/PDFファイルを最大300秒待ちます。通常確認では`Default`で十分です。

### Custom...

`MD Extract`の`Custom...`では、起動時に次を指定できます。

1. `Timeout seconds`：最初の処理で待つ秒数
2. `Retry timeout seconds`：Deferred再試行時に待つ秒数
3. `Deferred action`

Deferred actionの番号は次のとおりです。

- `0`：Deferredがあれば、その場で再試行するか質問する
- `1`：質問せず自動再試行する
- `2`：質問せずDeferredのまま終了する

## 13. テストデータを初期状態へ戻す

MD Patchの正常例を適用した後などに実行します。

1. Windowsキー＋Rを押します。
2. 次の1行を貼り付けます。

   `powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "C:\repos\pub\toolrack\test\create-manual-fixtures.ps1" -IncludeLegacyOffice`

3. Enterを押します。
4. 黒いPowerShell画面が表示されないことを確認します。
5. 作り直しが終わると、`manual-test-fixtures`がエクスプローラーで開きます。

この処理は`manual-test-fixtures`全体を削除して作り直します。手動で追加したファイルも消えるため、必要なものは別の場所へ移してから実行してください。

表示付きで原因を確認したい場合は、toolrack直下で次を実行します。

`powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\test\create-manual-fixtures.ps1 -IncludeLegacyOffice`

## 14. 実際にAIと使う場合

### 既存プロジェクトを変更する

1. MD Mirrorで対象コードを1個のMDにする
2. そのMDをAIへ渡す
3. AIからMD Patchを受け取る
4. MD Patchで元のプロジェクトへ適用する
5. 自動バックアップを確認する

### 新しいプロジェクトを作る

1. `Copy Format Instructions`を実行する
2. コピーされた説明と作成依頼をAIへ渡す
3. AIからMD Mirror形式のMDを受け取る
4. `Restore from Clipboard`で新しい`output`フォルダへ復元する

### 文書をAIへ渡す

1. MD Extractで文書群を1個のMDにする
2. 作られたExtract BundleをAIへ渡す

## 15. テスト完了後の片付け

今回の手動確認状況は、リポジトリ直下の`TODO-MANUAL-TEST.md`で管理します。

確認完了後に削除してよいもの：

- `TODO-MANUAL-TEST.md`
- `output\manual-test-fixtures`
- 手動確認で作られた`output\md-mirror_*`
- 手動確認で作られた`output\backup_*`
- 手動確認で作られた`output\extract-md_*`
- 手動確認で作られた`log\extract-md_*`

今後の回帰確認に使うため、次は残します。

- `test\MANUAL-FIXTURES.md`
- `test\create-manual-fixtures.ps1`
- `test\test-manual-fixtures.ps1`
