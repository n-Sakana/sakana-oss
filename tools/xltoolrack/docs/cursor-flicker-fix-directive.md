# 実装指示書: カーソルちらつきの知覚閾値未満化

作成: 2026-07-12 / 作成者: 診断マネージャー(Claude Fable 5) / 実装担当: Codex

この文書は `docs/cursor-flicker-review-prompt.md` の設問群に対する**最終回答**であり、
以降の実装作業の唯一の指示書である。旧文書と矛盾する場合は本書が優先する。

> **2026-07-12 夕 更新**: 本書 §2 の実装タスクは第1フェーズとして実施済み
> (結果: `docs/measurements/2026-07-12-cursor-fix-results.md`。ただし A/B 判定は
> 計測不備で無効と裁定)。**以降の作業は `docs/cursor-flicker-phase2-directive.md`
> (第2フェーズ指示書)に従うこと。** 本書の §0 診断・§4 地雷・§5 制約は引き続き有効。

---

## 前提(このファイル単体を渡された場合はここから)

対象リポジトリ: `C:\repos\pub\xltoolrack`

**最初に読むこと(この順で):**

1. `docs/edit-freeze-and-fe-pump.md` — 用語(FE/BE/worker/pump/tick)、アーキテクチャ全体、
   凍結根治の経緯、ビルド/テスト/インストール手順(§6)、環境上の注意(§7)
2. `docs/HANDOFF-codex.md` — ちらつき調査の履歴と地雷(ただし症状名 IDC_APPSTARTING は
   本書 §0 のとおり誤認と確定済み)
3. ソース: `src/addin/JobPump.bas`(OnTimeポンプ)、`src/common/JobHost.cls`(PumpOnce)、
   `src/common/Infra_Dispatch.bas`(OnResult配送)、
   `src/tools/stopwatch.bas`(改修対象ツール)、`src/common/ChannelFile.bas`(ファイル
   プロトコル)、`scripts/Build-Addin.ps1`(ビルド時validator。**注意: `ToolRegistry` は
   ソースファイルではなく、このスクリプトがツール注釈から生成する**。Registry_* 群の
   生成コードは同スクリプト内 231〜271 行付近、`Add-GeneratedModule` で注入)

**主要コマンド**(詳細は edit-freeze-and-fe-pump.md §6):
ビルド= `scripts\Build-Addin.ps1`(または `build.bat`) / 全テスト15ケース= `test.bat` /
インストール= `scripts\Install-Addin.ps1` / アンインストール= `uninstall.bat`。
source を変更したら dist を再ビルドしてから検証すること(旧バイナリ検証は既知の事故)。

**環境注意**: このマシンでは複数の AI エージェントが並行作業する。main worktree は
意図的に未コミットの変更を含む。reset / clean / checkout による破壊、既存 Excel PID の
kill、`%TEMP%` の他人の実験物への接触は禁止(詳細 §5)。

---

## 0. 確定した診断(2026-07-12、実測ベース)

### 症状の正体は IDC_WAIT フラッシュ。IDC_APPSTARTING ではない

desktop 実機・実使用中(3ツール同時、Excel 8プロセス、13:12:40〜13:14:10 の90秒、
64Hz `GetCursorInfo` サンプリング)の実測:

- **IDC_APPSTARTING (0x10019): 出現 0 回**
- **IDC_WAIT (0x10007): 13回**。マウスがセル領域上にあった約40秒間に約2〜4秒間隔
- 1回のフラッシュ長: **11〜30ms**
- フラッシュの合間は Excel の文脈カーソル(白十字=非共有ハンドル、I字等)が正常
- **同区間でオーナーが肉眼でちらつきを確認** → 体感と計器の同定完了

生データ: `docs/measurements/2026-07-12-cursor-baseline/`(これが改修前ベースライン)

従来ドキュメントの「IDC_APPSTARTING(矢印+spinner)」という症状名は**誤認**。
Win11 の WAIT(青い輪)を APPSTARTING と誤記していた。

### 機構モデル(note 実機の統制実験で確認済み)

note(win-note、非ロック・カーソル可視・SendInput実入力・200Hz spin-wait観測)にて:

| 実験 | 内容 | 結果 |
|---|---|---|
| N-B | OnTime 毎秒 300ms busyループ(単体ブック) | WAIT 56 onset/60s ≒ 毎tick |
| N-C | 同 12ms | 捕捉はビート周期(約5秒毎)に低下 |
| N-D | **次回予約のみの空tick** | それでも WAIT フラッシュ発生(約5ms級) |
| N-E | 製品 Multi Stopwatch(rollback build 0AE109F5...9746) | WAIT のみ、5ms級 |

結論:
1. **VBA が OnTime callback で実行されている間、Excel は WAIT カーソルを掲示する。**
   フラッシュ長 ≒ tick 実行時間(+メッセージポンプまでの復帰遅延)。
2. **空 tick でも数msは光る。完全ゼロは 1Hz FE ポンプ構成では不可能**
   (review-prompt 設問7 への回答)。
3. **5ms級のフラッシュは人間に見えない。15〜30ms級は数秒に1回目に留まる。**
   desktop で見えるのは tick が重くフラッシュが長いから。「2〜3秒に1回」の周期は
   毎秒のフラッシュに対する捕捉確率(マウス移動×WM_SETCURSORタイミング)の現れ。
4. `DoEvents` 2個除去で悪化した理由: 復帰(通常カーソルへの回復)がポンプ依存の
   ため、出口 DoEvents を失うとフラッシュが伸びる。**出口 DoEvents は必須。維持。**

### 勝利条件(定量)

> **desktop 実使用(3ツール同時)で、WAIT フラッシュ長の P95 < 10ms、
> かつオーナーが「見えなくなった」と確認すること。**

S1〜S10 の累積 matrix は不要になった(原因段= tick 長そのものと確定したため)。

---

## 1. 計測プロトコル(全タスク共通)

- 観測器: `scripts/Measure-Cursor.ps1`(`-Seconds N -OutDir path`)。
  遷移CSV・秒別busy-ms・サマリを出力。**配布VBAには一切手を入れない外部観測器。**
- 改修の before/after は必ず desktop で、オーナー通常使用中に 90〜120秒計測して
  WAIT onset 頻度と長さ分布(P50/P95)をベースラインと比較する。
- 計測時の地雷(§4)を必ず守ること。

## 2. 実装タスク(優先順・1変数ずつ・各々に rollback 基準)

### T0: tick プロファイラ(test-only 計測基盤)

- `PumpOnce`/`Pump_Tick` に **メモリ上 ring buffer**(固定長配列)で
  「tick開始 Timer 値、所要ms、fresh record 数、aggregate read 所要ms」を記録する
  test-only 機構を追加。**steady-state のファイルI/Oは禁止**。dump は
  `Pump_Stop`/ジョブ全終了時に1回だけ、または test 用エントリポイントから。
- 目的: desktop 実使用での tick 長分布と、フラッシュ長との相関確認。
  aggregate read の所要が跳ねるなら Defender 等の I/O レイテンシ(T3)へ。
- rollback 基準: 15ケーステストに1つでも回帰があれば即戻す。計測値が
  取れないだけなら本体に影響しないこと(On Error で握る)。

### T1(本命): 可視シート書込みの coalesce

- 現状: 同一 tick に複数 job が fresh だと `OnResult` が job ごとに呼ばれ、
  stopwatch は 1x3 `Value2` を最大3回書く。3ツール同時ではさらに増える。
  フラッシュ長 ≒ tick 長なので、書込み回数の削減が最も直接効く。
- 要件: **同一 tick 内の可視シート書込みを、ツールあたり最大1回にする。**
- 推奨設計: Registry にオプショナルな `OnFlush(ctx)` フックを追加。
  `PumpOnce` の dispatch ループ後、その tick で1件以上 result を受けたツールに
  対して1回だけ呼ぶ。stopwatch は `OnResult` でモジュール内バッファに溜め、
  `OnFlush` で `B3:D5` へ 3x3 を1回書く。フックを持たないツールは従来どおり。
- 注意: **`ToolRegistry` はソース非実在。`scripts/Build-Addin.ps1` がツールモジュールの
  注釈/プロシージャ検出から生成する**ため、`OnFlush` 対応は (a) 生成器に
  `Registry_HasFlush` / `Registry_OnFlush` の生成を追加、(b) ツールの `OnFlush`
  プロシージャ検出とシグネチャ検証を追加、の2点セットになる。既存の
  `Registry_HasResult` / `Registry_OnResult` の生成コード(231〜271行付近)を雛形に
  すること。VBA は ASCII のみ(validator が非ASCIIを拒否)。
- 検証: 15ケース green + coalesce の unit test 追加 + desktop before/after 計測。
- rollback 基準: P95 が改善しない、またはテスト回帰、または編集フリーズ
  受け入れテスト(実編集 F2 10秒)が FAIL なら戻す。

### T2: tick 末尾/内部のスリム化(T0 のプロファイル駆動でのみ)

- T1 後も P95 ≥ 10ms の場合のみ。プロファイルで支配的な成分を特定してから
  1変数ずつ削る。当てずっぽうの削減は禁止。
- 既知の候補: aggregate decode(job数比例)、`AppStateSnapshot` 生成、
  dispatch scaffolding(Logger/Status/InfraContext の毎回 New)。
- 参考(診断で確認済みの無罪): `Status.Clear` は dispatch 経路では no-op。
  error file チェックは payload age ≥ 2s のみ。sweep は Timer 演算のみ。

### T3: チャネルファイル I/O レイテンシ(T0 で aggregate read の跳ねが観測された場合のみ)

- 仮説: FE は毎 tick「worker が直前に書き換えたばかりのファイル」を open するため、
  Defender のオンアクセススキャンで open が延びる tick がある。
- 対応候補: チャネルファイルを `%TEMP%` 直下から専用サブフォルダ
  (`%TEMP%\xltoolrack\`)へ移す改修(パス生成は `PathFs` に集約済み)。
  その上でオーナーに当該フォルダの Defender 除外を**提案**する。
  **除外設定を勝手に入れない。** パス変更自体の rollback 基準: 15ケース green +
  worker 死活(リース/tombstone)の実インストール検証 PASS。

### T4(最終手段・オーナー承認必須)

- tick 頻度削減(1s→2s、stopwatch の秒表示が飛ぶ UX 劣化あり)
- tick 中のみ `Application.Cursor = xlNorthwestArrow` ピン
  (スピナー点滅→矢印点滅への置換。validator が現在 `Application.Cursor` を
  拒否しているため validator 変更も必要)
- いずれも T1〜T3 で勝利条件未達の場合に、オーナーへ選択肢として提示するに留める。

## 3. ドキュメント更新タスク(実装と同時に)

1. `docs/HANDOFF-codex.md` と `docs/edit-freeze-and-fe-pump.md` の症状名を訂正:
   「IDC_APPSTARTING」→「IDC_WAIT(APPSTARTING は実測0)」。
2. 「note の 0/0/0 計測」への注記: 当時 note が**ロック画面状態だった疑いが濃く、
   ロック中は WM_SETCURSOR が発生せずカーソル形状が更新されないため、
   ゼロ計測は無効**。以後のカーソル計測はロック状態検証(`GetForegroundWindow != 0`
   等)を開始/終了時に記録すること。
3. §4 の地雷リストに本書 §4 の新規地雷を追記。
4. S1〜S10 累積 matrix の節に「2026-07-12 診断確定により不要」と注記。

## 4. 新たに確認された計測上の地雷(必読)

- **セッションロック**: ロック中はカーソル形状が一切更新されず、全計測が無音で
  無効化される。計測の前後でロック状態を検証・記録する。
- **`Application.OnTime` の非修飾マクロ名は無音で不発**
  (エラーも出ない)。必ず `'<book名>'!Module.Proc` 形式で修飾する。
- **CURSOR_SUPPRESSED**: タッチ入力後は CURSORINFO.flags=2 でカーソル非表示になり
  計測不能。計測用のマウス移動は `SetCursorPos` でなく `SendInput` 実入力で行う。
- **バックグラウンドプロセスのタイマー間引き**: Sleep ベースのサンプリングは
  約64Hz が上限(15.6ms)。それ以上は spin-wait が必要(無人機でのみ許容)。
- **onset 回数だけで比較しない**(既存地雷の再掲)。長さ分布(P50/P95)を主指標に。

## 5. 継承する制約(変更なし)

- 配布 VBA に `Declare` / Shell / WMI / `taskkill` 禁止。ASCII のみ。
- `Application.Cursor` / `ScreenUpdating` toggle をポンプ経路に入れない(T4 例外は
  オーナー承認時のみ)。
- **`Pump_Tick` の入口/出口 `DoEvents` は維持**(出口除去は実証済みの悪化要因)。
- BE worker から FE への同期 COM 禁止。`Application.OnTime` は
  `WorkerBridge.bas` / `JobPump.bas` のみ。
- main worktree は複数エージェントの未コミット作業を含む。reset / clean /
  checkout での破壊禁止。既存 Excel PID への不干渉。
- source 変更後は dist を再 build / install し、SHA-256 で実体確認。
- 改修前の既知良好ビルド: SHA-256
  `0AE109F5F7BE0D904AE6DCAA424769CA7E6843B86A7EB864D4675C0199909746`
