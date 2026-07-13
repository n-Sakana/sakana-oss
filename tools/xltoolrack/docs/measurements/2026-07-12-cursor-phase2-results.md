# カーソルちらつき第2フェーズ結果（進行中、2026-07-12）

## 現在の結論

P2-1 と P2-2 は完了し、P2-3 は Defender 除外前まで完了した。

- 250 Hz の WAIT 観測器は `note` で有効性を確認済み。
- rollback 基準は 90 秒 x 3 窓すべて有効、WAIT onset 合計 359。
- `%TEMP%\xltoolrack\` へのチャネル移動後・Defender 除外前も 90 秒 x 3 窓
  すべて有効、WAIT onset 合計 304。
- 集約分布では、rollback → サブディレクトリのみで tick P95 は
  39.062 → 31.250 ms、aggregate read P95 は 23.438 → 15.625 ms だった。
  ただし時間帯をまたぐ非交互比較なので、これだけでパス移動の効果とは断定しない。
- Defender 除外はまだ一度も追加されていない。UAC は 2 回とも Windows から
  「ユーザーによって取り消されました」と返り、昇格スクリプトは開始前に停止した。
- P2-4 は、P2-3 の除外後データが未取得のため未着手。

## 実行条件

- デバイス: Helm の `note`
- サブディレクトリ版 worktree:
  `C:\repos\_codex_work\xltoolrack-phase2-codex-20260712`
- rollback 基準 worktree:
  `C:\repos\_codex_work\xltoolrack-phase2-baseline-20260712`
- base: detached `52f8dde69808743ab7aac823f84c6d291d57beb4`
- Excel の前景化と SendInput mouse-move-only は、オーナーの明示承認後に使用。
  click、key、セル編集は送信していない。
- 各正式窓は Excel 前景率 100%、開始終了とも非ロック、可視、カーソル移動あり。
- 各正式窓は 7 worker を開始から終了まで維持。
- 既存 Excel PID は終了していない。各操作前に Excel 0 件を確認し、試験が生成した
  PID のみを正常停止または、終了不能時に PID/ウィンドウを再確認して停止した。
- raw: `phase2-note-evidence-20260712/`

オーナーの通常操作ではなく統制入力による無人比較なので、体感結果は未聴取。

## P2-1: WAIT サンプリング修理

`Measure-Cursor.ps1` / `Measure-Desktop.ps1` に以下を実装した。

- `Stopwatch` deadline + spin-wait、既定 250 Hz。
- 観測中だけ sampler process を High priority にし、`finally` で復帰。
- WAIT の開始・終了遷移時刻差から duration を算出。
- 実効レート、nominal resolution、gap P50/P95/max、missed deadline を記録。
- 対象 Excel の前景 endpoint と前景率、可視 endpoint、カーソル移動を検証。
- 90 秒あたり 70 pump tick、7 worker start/end を有効性条件に追加。

`note` 上の最終 sampler smoke（3 秒、製品 workload なし）:

| 要求 | sample | 実効 | 分解能 | gap P50 | gap P95 | gap max | missed | priority |
|---:|---:|---:|---:|---:|---:|---:|---:|---|
| 250 Hz | 750 | 249.9 Hz | 4.000 ms | 4.000 ms | 4.006 ms | 5.496 ms | 0 | High |

正式窓でも cursor gap P95 は 4.15--4.18 ms 級、前景率はすべて 100% だった。

## P2-2: rollback 正規基準

build（`note`、再起動後）:

| artifact | SHA-256 |
|---|---|
| `xltoolrack.xlam` | `4E9F30ACF485CA7A47C2CF516975222E0CAF106AC0787B022AEF7A86A5C24D9D` |
| `xltoolrack-worker.xlsm` | `C8E5A2DE48BDDB443CD4CDBA2917B1EFB045570C2EEDD9B2C78C148D774EBCE8` |

90 秒窓:

| 窓 | 有効 | tick | tick P50/P95 | aggregate P95 | WAIT onset | WAIT P50/P95/max |
|---:|---|---:|---:|---:|---:|---:|
| 1 | yes | 90 | 23.438 / 39.063 ms | 15.625 ms | 141 | 4.011 / 8.000 / 12.030 ms |
| 2 | yes | 90 | 23.438 / 39.063 ms | 15.625 ms | 112 | 4.006 / 8.030 / 11.973 ms |
| 3 | yes | 91 | 23.438 / 39.063 ms | 23.438 ms | 106 | 4.008 / 7.996 / 12.009 ms |

raw を連結した構成集約:

| 窓/有効 | tick | tick P50/P95 | aggregate P50/P95 | WAIT onset | WAIT P50/P95/max | onset/min |
|---:|---:|---:|---:|---:|---:|---:|
| 3 / 3 | 271 | 23.438 / 39.062 ms | 7.812 / 23.438 ms | 359 | 4.000 / 8.000 / 12.000 ms | 79.78 |

再起動後の一括 15-case は `test-samples all` まで green だったが、最後の
`Validate-E2E` で life generation=0 となり exit 1（14/15）。同じ実装の
`test-samples all` では stopwatch 3、pi 3、life 1 の全 worker が進んだため、
実装破損ではなく 7-worker publish 競合の時間依存再現としてログを保存した。

## P2-3: チャネルサブディレクトリ

実装:

- aggregate/input/stop/error と lock を `%TEMP%\xltoolrack\` 配下へ移動。
- lease、shutdown tombstone、done、worker copy は従来どおり TEMP 直下。
- 起動時 stale sweep は旧 root-level channel と新サブディレクトリの両方を掃除。
- Defender helper は除外可能なパスを `%TEMP%\xltoolrack` に hard-code し、
  TEMP 全体や任意パスを受け取らない。

build（`note`、再起動後）:

| artifact | SHA-256 |
|---|---|
| `xltoolrack.xlam` | `392AB69FAC4013ACB1399A237C8D59E28DB21AFE6EF710AD21E2543703CAD5C4` |
| `xltoolrack-worker.xlsm` | `7BA0D04A91F7A18621101EFD50CD487ECE75088B1F902D0361CC776A6A1F95F0` |

### 自動テスト

一括 Run-All は先頭 12 case green 後、`test-worker-busy-fe` の回復確認で
stopwatch/probe version が 47→47 のままになり exit 1。worker 生存、stop、sweep は
green だった。同一 build の単独再試行では当該 case が green、その後の
`test-samples all` と `Validate-E2E` も green となった。

E2E の主な値:

- 7 worker 全員進行: stopwatch `[13,12,11]`、pi `[8,7,5]`、life `5`
- 連続 30 cell input 成功
- FE cell round-trip max 31 ms（上限 750 ms）

したがって同一 build で 15 case 全件の green は確認したが、一括初回の transient
失敗も棄却せず raw に残している。既知の `test-ribbon-ui` 終了時モーダルは assertion
green 後に test-owned PID だけを VBE reset して継続した。

### Defender 除外前測定

| 窓 | 有効 | tick | tick P50/P95 | aggregate P95 | WAIT onset | WAIT P50/P95/max |
|---:|---|---:|---:|---:|---:|---:|
| 1 | yes | 91 | 15.625 / 31.250 ms | 15.625 ms | 100 | 4.008 / 7.919 / 12.070 ms |
| 2 | yes | 90 | 23.438 / 31.250 ms | 23.438 ms | 107 | 4.013 / 4.218 / 8.035 ms |
| 3 | yes | 90 | 23.438 / 31.250 ms | 15.625 ms | 97 | 4.000 / 4.119 / 7.993 ms |

raw を連結した構成集約:

| 窓/有効 | tick | tick P50/P95 | aggregate P50/P95 | WAIT onset | WAIT P50/P95/max | onset/min |
|---:|---:|---:|---:|---:|---:|---:|
| 3 / 3 | 271 | 23.438 / 31.250 ms | 7.812 / 15.625 ms | 304 | 4.000 / 4.300 / 12.100 ms | 67.56 |

## Defender 除外記録

現時点では除外なし。

- UAC attempt 1: user-cancelled before elevated process start
- UAC attempt 2: user-cancelled before elevated process start
- `exclusion.ready`: なし
- `before.json` / `excluded.json` / `final.json`: wrapper 未起動のため未生成
- `Add-MpPreference`: 未実行
- `Remove-MpPreference`: 追加がないため未実行

次の UAC はオーナーが `note` の前にいる時だけ再実行する。成功時は
`before.json` → `excluded.json` → 除外後 90 秒 x 3 → `measurement.done` →
`final.json` / `exclusion.removed` の順で確認する。

## P2-4 / P2-5

- P2-4: Defender 除外後の有効比較がないため未着手。
- P2-5: T4 は提示・実装とも未実施。承認なしに変更しない。

## raw / 集計ファイル

`phase2-note-evidence-20260712/`:

- `p22-rollback-baseline/`
- `p23-subdir-pre-exclusion/`
- `logs/`
- `phase2-summary.csv` / `phase2-summary.json`（最終 archive で追記予定）
- `phase2-windows.csv`（最終 archive で追記予定）

個別 sampler smoke は `phase2-note-p21-smoke-final/` に保存済み。
