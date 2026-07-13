# xltoolrack

Excelの中で、重い処理・常駐処理・複数処理を別Excelプロセスへ逃がして実行するアドイン基盤です。FEはユーザー操作と描画、各BEは1ジョブを担当します。

## 特徴

- 全toolの`Run`をAppGuard / Logger / Status / ErrorHandlerで包む必須ハーネス
- 1ジョブ=1不可視Excelプロセスによる真の並行実行
- worker結果はsession単位のaggregate snapshotへ集約。input / stop / errorはjob単位、worker→FEのCOM呼び出しはゼロ
- FE所有のOnTime 1秒pumpはaggregateを1回だけ読み、各jobのversion進行とOnResultを配送
- OnTimeはセル編集・ダイアログ中はExcelがネイティブに保留するため、ユーザー操作と衝突しない
- pumpはcursorを変更せず、tick前後のUI yieldでExcel標準のcontext cursorを維持
- worker側は1秒自己ペーシングループ。toolによるOnTime直接使用は引き続き禁止
- stopフラグ、shutdownフラグ、FEリースファイルのロック探知によるworker自己Quit（クラッシュ即検知）
- tool単位上限とFEインスタンス全体上限（既定10）
- toolメタデータからRibbonと静的dispatch registryを自動生成
- 非ASCII、Declare、Shell、WMI、迂回エントリ、不正signatureをビルド前に拒否

WinAPI Declare、Shell、WMI、PID推定、taskkillは使用しません。VBAソースはASCIIです。

## 前提

- Windows 11
- Excel 16.0
- Windows PowerShell 5.1
- Excelの「VBAプロジェクト オブジェクト モデルへのアクセスを信頼する」が有効

## 使う

`xltoolrack.bat`をダブルクリックします。アドインの登録後にExcelが開き、リボンの
`xltoolrack`タブへ次の3ツールが表示されます。

- Conway Life
- Parallel Pi Race
- Multi Stopwatch

ツールはリボンのボタンから起動します。マクロ一覧、VBAエディター、
`xltoolrack-test.xlsm`、`xltoolrack-worker.xlsm`をユーザーが開く必要はありません。
登録を解除する場合は`uninstall.bat`を実行します。

## ビルド

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts\Build-Addin.ps1 -OutputFormat all
```

`dist`へ次を生成します。

- `xltoolrack.xlam` — FEアドイン
- `xltoolrack-test.xlsm` — COM検証用ホスト
- `xltoolrack-worker.xlsm` — BE workerテンプレート

アドインとworkerは同じフォルダーに置いてください。ジョブ開始時にworkerをジョブ固有名で`%TEMP%`へコピーします。

## tool契約

`src/tools`へ標準モジュールを追加します。VBA module名がtoolIdです。

```vba
Attribute VB_Name = "my_tool"
Option Explicit
Option Private Module
'@name My Tool
'@ribbon My Tool
'@group Data
'@maxjobs 3

Public Sub Run(ByVal ctx As InfraContext)
    Dim jobId As String
    jobId = ctx.RunJob("my_tool.Worker", "{}")
End Sub

Public Sub Worker(ByVal job As InfraJob)
    ' One bounded tick. The infrastructure owns the loop and cadence.
    Call job.Push("result", "A1", 123)
End Sub

Public Sub OnResult(ByVal ctx As InfraContext, ByVal jobId As String, ByVal version As Long)
    Dim value As Variant
    value = ctx.ReadJob(jobId, "A1")
End Sub
```

`Run`だけが必須です。BEを使うtoolは`Worker`、push結果をFEへ描画するtoolは`OnResult`も定義します。Workerは常駐ループを書かず、1回分のboundedな仕事をして戻ります。

主なctx API:

- `ctx.LogMessage message, [level]`
- `ctx.Status fraction, message`
- `ctx.RunJob(workerEntry, argsJson) As String`
- `ctx.StopJob jobId`
- `ctx.Target As Workbook`
- `ctx.ReadJob(jobId, address)`
- `ctx.BindInput jobId, worksheet, address`
- `ctx.RefreshInput jobId`（結果描画中に入力範囲も更新する場合）
- `ctx.Fail message`

主なjob API:

- `job.Args`
- `job.FEWorkbook`
- `job.Alive`
- `job.Push(logicalName, address, values)`
- `job.ReadInput(address)`
- `job.LogMessage message`

`logicalName`はtool側の意味名です。最新payloadはメモリから直接OnResultへ渡し、任意範囲読み取り時だけ隔離sheetへ遅延反映します。

## 上限と失敗処理

- `@maxjobs`は1toolの同時BE数。1～10を指定します。
- FEインスタンス全体の既定上限は10です。
- Starting / Running / Stoppingも上限枠へ算入します。
- StartJob途中で失敗した場合、BE、COM参照、チャネル、tempコピー、予約を巻き戻します。
- 1回のRun中で後続StartJobが失敗した場合、そのRunが開始済みのjobだけを停止します。
- BE例外はerror payloadとversionでFEへ通知し、ログとfailed状態を残して自己Quitします。

ログは既定で`%LOCALAPPDATA%\xltoolrack\log`へ保存します。テストでは`XLTOOLRACK_LOG_DIR`で一時フォルダーへ切り替えます。

## サンプル

- Multi Stopwatch — 3つの独立した常駐job
- Parallel Pi Race — 3つの計算BEを並行稼働
- Conway Life — BEで世代計算し、実行中のセル編集を入力チャネルへ反映

## 自動検証

全COMハーネスをWindows PowerShell 5.1で実行します。

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File test\Run-All.ps1
```

主な検証内容:

- AppGuardの通常・ネスト・例外復元
- error/log/status必須ハーネス
- N並行の独立ルーティング
- tool/global上限と部分起動rollback
- 3並行を30秒観測し、各1秒更新・最大gap 5秒未満
- BE例外のFE捕捉、stop、FE消失時の自己Quit
- build契約と迂回パターン拒否
- 3サンプルの動作と全部盛りE2E

## 手動チェックリスト

1. `Build-Addin.ps1 -OutputFormat all`を実行し、xlam / test xlsm / worker xlsmを確認する。
2. `xltoolrack.xlam`をExcelアドインへ登録し、workerを同じフォルダーへ置く。
3. RibbonからParallel Pi Raceを起動し、タスクマネージャーで複数コアが動くことを確認する。
4. Conway Lifeを起動し、盤面が1秒ごとに進み、実行中に描いたセルが次世代へ反映されることを確認する。
5. Multi Stopwatchを起動し、3本が別々に1秒刻みで進むことを確認する。
6. 3サンプルを同時起動し、横のセルへ入力してもFEが固まらないことを確認する。
7. `@maxjobs`超過と全体10本超過が明示的に拒否されることを確認する。
8. tool／BE例外がログと通知へ到達し、ScreenUpdating / DisplayAlerts / Calculationが元の値へ戻ることを確認する。
9. stopおよびFE終了後に不可視worker Excelが残らないことを確認する。
10. 通信、Shell、WMI、taskkillが発生しないことを確認する。
