# 実装指示書 第2フェーズ: カーソルちらつき — 計測修理と Defender 実験

作成: 2026-07-12 夕 / 作成者: 診断マネージャー(Claude Fable 5) / 実装担当: Codex(新規セッション)

本書は `docs/cursor-flicker-fix-directive.md`(第1フェーズ)の **§2 実装タスクを置換**する。
診断(§0)・計測地雷(§4)・制約(§5)は第1フェーズのものが引き続き有効。

---

## 前提(このファイル単体を渡された場合はここから)

対象リポジトリ: `C:\repos\pub\xltoolrack`

**最初に読むこと(この順で):**

1. 本書(最後まで)
2. `docs/cursor-flicker-fix-directive.md` — 確定診断(症状= IDC_WAIT フラッシュ、
   フラッシュ長≒tick実行時間、勝利条件 P95<10ms+体感消失)、制約、計測地雷
3. `docs/measurements/2026-07-12-cursor-fix-results.md` — 第1フェーズ実装の結果報告。
   **ただしその A/B 判定は本書 §2 のとおり無効と裁定済み。数値の再解釈をせず §3 の
   再計測からやり直すこと**
4. `docs/edit-freeze-and-fe-pump.md` — アーキテクチャ・用語・ビルド/テスト手順(§6)
5. ソース: `src/addin/JobPump.bas`(T0プロファイラ入り)、`src/common/JobHost.cls`、
   `src/tools/stopwatch.bas`、`src/common/ChannelFile.bas`、`src/common/PathFs.bas`、
   `scripts/Measure-Desktop.ps1`、`scripts/Build-Addin.ps1`(ToolRegistry はここが生成)

**主要コマンド**: ビルド= `scripts\Build-Addin.ps1` / テスト15ケース= `test.bat` /
インストール= `scripts\Install-Addin.ps1`。source 変更後は dist 再ビルド+SHA-256確認。

**環境注意**: 複数 AI エージェントが並行作業する。main worktree は未コミット作業を含む。
reset / clean / checkout での破壊、既存 Excel PID の kill 禁止。オーナー在席中の
前面強制・マウス合成禁止(計測はオーナーの通常使用中に受動で行う)。

## 1. 現在の状態(2026-07-12 15:30 時点)

- インストール済み dist: xlam `F09F5916A5059B5A396C213EBDFC74D75A172520B2960C5C4EA04A134C01D02F`
  / worker `BBBEB7369696EE4E59C98FD899332A84B9220A998EF7F7C22450CD876F099D8A`
- 15ケース all passed(`docs/measurements/2026-07-12-full-suite.log`)。既知の注意:
  `test-ribbon-ui` は検証green後の終了時に VBA「ファイルが見つかりません」モーダルが
  出ることがあり、無人実行がクリーンでない
- **T1(coalesce)の runtime は rollback 済み**。ただし `OnFlush` の生成基盤
  (Build-Addin の `Registry_HasFlush`/`Registry_OnFlush` 生成+シグネチャ検証)は
  残置。対応ツール0件、`PumpOnce` からの呼出しなし → **T1 は再有効化が容易**
- **T0 プロファイラは残置**(JobPump 内 ring buffer 256件、`XLTOOLRACK_TEST=1` or
  `XLTOOLRACK_PROFILE=1` で有効、dump は `JobTest_PumpProfile`)
- T2(lazy context)/T3(チャネルのサブフォルダ移動)/life payload 圧縮(420→21列)は
  試行後 rollback。**ただしこれらの棄却判断もノイズ由来の可能性が高い**(§2)
- 計測スクリプト: `scripts/Measure-Cursor.ps1`(汎用カーソル観測)、
  `scripts/Measure-Desktop.ps1`(製品を起動して tick プロファイル+WAIT を測る)
- `HostMain.bas` の diff は worker 終了管理(リース/tombstone/掃除/Pump_Stop)で
  カーソルとは無関係(確認済み)

## 2. 第1フェーズ A/B 判定を無効と裁定した理由(同じ過ちを繰り返さないこと)

1. **WAIT P95 の標本数が 2〜12 個**。n<30 の P95 比較は成立しない。onset 捕捉率は
   オーナーのマウスの動きに依存するため、窓ごとの操作差で数字が勝手に動く。
2. **窓ごとの実行 tick 数が 77 / 24 / 39 と不揃い**(90秒なら本来約85)。OnTime が
   大量保留された窓は FE 状態が別物で、比較不能。保留明け tick は重くなるので
   体感悪化もこれで説明がつく。
3. **数値が VBA `Timer` の分解能 3.906ms(1/256秒)に量子化**されている
   (7.812/11.719/15.625/19.531 は全て 3.906 の整数倍)。1目盛差での採否判定は
   ノイズ判定。lazy-context・life圧縮の棄却(いずれも1目盛差)も同様に無効。

## 3. 有効性ルール付き計測プロトコル(全実験の前提)

**P2-1: Measure-Desktop.ps1 の WAIT サンプリングを修理する**

- サンプリングを Sleep ベース(実効64Hz)から **Stopwatch spin-wait 200Hz以上+
  プロセス優先度 High** に変更(90秒×数回なら desktop でも許容。計測終了で即解放)
- WAIT 区間の長さは遷移検出時刻の差で算出し、サンプリング周期(≦5ms)を分解能として明記

**窓の有効性ルール(1つでも欠けたらその窓は棄却し、取り直す):**

- 90秒窓で実行 tick 数 ≥ 70(プロファイラの行数で判定)
- 計測開始/終了時に非ロック(`GetForegroundWindow() != 0`)
- FE が可視・前景で、オーナーが通常操作中(マウスがセル領域上を動く時間があること)

**比較の有効性ルール:**

- 各構成 90秒×3窓、**A/B/A 交互**(時間帯ドリフトの統制)
- 構成ごとの WAIT onset 合計 ≥ 30 個。足りなければ窓を追加
- 判定は分布(P50/P95/max と onset/分)で行い、単一数値の1目盛差(tick系は
  3.906ms、WAIT系はサンプリング周期)以内の差は「差なし」と扱う
- 結果は `docs/measurements/` 配下に raw+サマリで保存

## 4. 実験タスク(この順で)

### P2-2: 現行 rollback ビルドの正規ベースライン

§3 のプロトコルで、3ツール同時(stopwatch+pi_race+life、7 jobs)の
tick 分布と WAIT 分布を取り直す。以降の全判定の基準線。

### P2-3: Defender 除外実験【オーナー承認済み 2026-07-12・実験目的・可逆】

背景: ステージ分解で tick の支配項は **aggregate read(P95 4〜8ms)+ dispatch
(4〜8ms)**であり、書込み coalesce ではない(error/cleanup は 0ms)。note 実機では
同じ製品のフラッシュが約5msだったため、desktop の aggregate read はディスクではなく
**Defender オンアクセススキャン(毎tick、worker が書き換えた直後のファイルを open)**
の疑いがある。第1フェーズ T3 の「サブフォルダ移動のみ」は除外なしでは無意味な実験だった。

手順:
1. チャネルファイル群を `%TEMP%\xltoolrack\` 配下へ移す(`PathFs` のパス生成を変更。
   第1フェーズ T3 の実装を復活させてよい)。ビルド→インストール→15ケース green 確認
2. 除外前の計測(§3 プロトコル)
3. `Add-MpPreference -ExclusionPath "$env:TEMP\xltoolrack"` で**そのフォルダだけ**除外
   (実行前後に `Get-MpPreference | Select ExclusionPath` を記録)。
   **`%TEMP%` 全体や他のパスを除外することは絶対禁止**
4. 除外後の計測(§3 プロトコル)
5. 完了後、**除外は `Remove-MpPreference -ExclusionPath ...` で必ず戻す**
   (恒久採用はオーナーの別途判断。本実験はデータ取りのみ)

判定:
- aggregate read P95 が 4〜8ms → 2ms 以下級に落ちる → **犯人確定**。恒久対応
  (フォルダ+除外の採用、またはスキャンに引っかかりにくい I/O 形態の検討)を
  オーナーに提案
- 変わらない → Defender 説棄却。aggregate read の中身(Open/Line Input/decode)を
  T0 プロファイラの計時点を増やして分解し、次を決める

### P2-4: T1・life圧縮の再判定(P2-3 の結果が出てから)

- T1(OnFlush coalesce)と life payload 圧縮は実装済み・容易に再有効化できる。
  §3 の有効な統計でベースラインと比較し、採否を**データで**決め直す
- 1変数ずつ。組み合わせる場合も追加は1つずつ

### P2-5: T4 の提示(最後の手段)

有効な計測で「床(入場+read+dispatch)が P95 10ms を下回れない」と確定した場合のみ、
第1フェーズ §2 T4 の2案(2秒tick / callback中の矢印ピン)を、実測見込みと UX 影響を
添えてオーナーに提示する。**承認なしに実装しない。**

## 5. 報告フォーマット

- 構成ごとの表: 窓数 / 有効窓数(棄却理由も) / tick 実行数 / tick P50/P95 /
  aggregate P50/P95 / WAIT onset 合計 / WAIT P50/P95/max / onset/分
- 判定は §3 の有効性ルールを通過した比較のみに基づくこと
- Defender 除外の実施記録(前後の `Get-MpPreference` 出力、戻したことの確認)
- 全 raw を `docs/measurements/` へ。オーナー体感の聴取結果も記録(見えた/見えない)

## 6. 経緯の要約(新規セッション向けタイムライン)

1. **朝**: セル編集フリーズはファイルチャネル+FE OnTime ポンプ化で根治済み(第0フェーズ、
   `edit-freeze-and-fe-pump.md`)。残課題=「2〜3秒に1回のカーソルちらつき」
2. **昼**: 診断で正体を確定。IDC_APPSTARTING は誤認で、実体は **OnTime callback 中の
   IDC_WAIT フラッシュ(長さ≒tick時間)**。note 統制実験で「空tickでも数ms光る=
   ゼロ不可」「5msは不可視」。desktop 実測 11〜30ms をオーナーが目視確認
3. **午後**: 第1フェーズ実装(T0 プロファイラ+T1 coalesce+T2/T3試行)。ビルド・
   テストは成功したが症状改善なしと報告され、T1〜T3 runtime は rollback
4. **夕(本書)**: 第1フェーズの A/B 判定は計測不備で無効と裁定。ステージ分解から
   「支配項= aggregate read+入場床」という新しい手がかり。計測修理(§3)→
   Defender 実験(P2-3、承認済み)→再判定(P2-4)→必要なら T4 提示(P2-5)の順で進める
