# xltoolrack: セル編集フリーズの根治とFEポンプ方式（知見まとめ）

最終更新: 2026-07-12 / 対象コミット: 作業ツリー未コミット（`52f8dde` の上）

このドキュメントは「タイマー実行中にセルを編集するとアプリ全体が固まる」問題の
調査・根治・残課題（カーソルのちらつき）を、後から入る人（人間 / Codex）が
最短で状況を把握できるようにまとめたもの。

---

## 0. 用語

- **FE (front-end)**: ユーザーが操作している通常の Excel。`xltoolrack.xlam` アドインが載る。
- **BE / worker**: 不可視の Excel プロセス。`xltoolrack-worker.xlsm` を実行する計算役。
- **ジョブ**: ツール（`stopwatch`, `pi_race`, `life` など）1 個の実行単位。worker 1 個に対応。
- **ポンプ (pump)**: FE 側で `Application.OnTime` により 1 秒ごとに回るループ。
  worker が書いたファイルを読み、FE のシートへ反映する（`JobPump.bas` / `JobHost.PumpOnce`）。

---

## 1. 何が起きていたか（根本原因）

### 症状
FE でタイマー系ツールを走らせている最中に、ユーザーがセルを編集状態（F2 やダブルクリック、
文字入力）にすると、**その瞬間に FE 全体が固まり、ESC でも抜けられない**。

### 確定した因果連鎖（旧アーキ）
旧アーキは **worker（別プロセス）→ FE のセルへ毎秒 同期 COM 書き込み** をしていた。

1. Excel は STA（単一メインスレッド）。**セル編集モードに入ると FE のメッセージポンプが実質停止**する。
2. その状態へ worker からの **クロスプロセス同期 COM 呼び出しがパーク**される。
3. パークされた呼び出しに対し **OLE チャネルが高頻度でリトライ**を投げる。
4. このリトライは Windows のメッセージキューに**ポストメッセージ**として積まれ、
   `GetMessage` の優先度で **`WM_KEYDOWN`/`WM_CHAR` より先に処理**される。
5. → FE のキー入力が**飢餓（starvation）**。**ESC が届かない**ので編集モードから出られない。
6. 編集モードが続く限り 1〜5 が自己維持し、**永久フリーズ**。

### 実測での裏付け
- FE の STA スレッド CPU 約 60%、worker 各 15% でスピン。
- `IsHungAppWindow` は **False**（＝OS 的には「応答なし」ではない。だからタスクマネージャで殺す以外に見えない）。
- ESC は送達できているのに無効。
- **worker 側の VBA 対策（busy 分類 / Ready ゲート / バックオフ）は原理的に無効**だった。
  パーク中は worker の VBA そのものが動けないため。

> ここが最大の教訓: **「編集中に外からセルを書く」設計は、対策を積んでも直らない。**
> 書き込みの方向を反転させるしかなかった。

---

## 2. 採用した解（案 A: ファイルチャネル + FE ポンプ）

**worker は FE を一切 COM で触らない。ファイルに publish し、FE が自分の都合で pull する。**

```
  worker (別プロセス)                    FE (ユーザーの Excel)
  ─────────────────                     ────────────────────
  計算する                               Application.OnTime 1秒ごと (JobPump)
    │                                       │
    │ lock下でaggregate snapshotを更新         │ aggregateを1回読む
    │ tmp 書き→Kill→Name (原子的)             │   ↓
    ▼                                       ▼
  %TEMP%\xltoolrack\xltoolrack_ch_<sid>_payloads.dat
                                              → memory cache + OnResult 配送
```

- **編集中は `Application.OnTime` が Excel によりネイティブに保留**される。
  ポンプが動かない＝FE を触らない＝衝突が物理的に消える。編集が終われば自動で再開。
- worker→FE の COM がゼロなので、パークもリトライ洪水も起きない。

### 主要ファイル
| ファイル | 役割 |
|---|---|
| `src/common/ChannelFile.bas` (新規) | aggregate payload snapshot、job別input/stop/error、N/S/E cell encoding、worker間lock、原子的publish。 |
| `src/common/ChannelRecord.cls` (新規) | aggregate内のjobId/version/address/values 1件。 |
| `src/addin/JobPump.bas` (新規) | FE の `Application.OnTime` ポンプ。`Pump_Start/Stop/Tick`、イベントからの `Pump_EnsureArmed` 再アーム。 |
| `src/common/JobHost.cls` | `PumpOnce`（1 tick の本体）、`StartJob`/`StopJob`、`MaybePump`、sweep。 |
| `src/common/InfraJob.cls` | worker 側。FE COM 参照ゼロ。`Push` はaggregateをlock付き更新。死活は `CheckChannelState`。 |
| `src/common/PathFs.bas` | `SessionId()`、各種パス生成。 |
| `src/addin/HostMain.bas` | リース確保 / tombstone / 起動時の残骸掃除。 |

### ファイルプロトコル（ChannelFile.bas）
- パスは全て `%TEMP%` 配下、`SessionId`（後述）でセッション名前空間化。
  高頻度に更新するチャネル群だけは Defender 実験を狭い範囲に限定できるよう
  `%TEMP%\xltoolrack\` に集約する（lease / tombstone / done / worker copy は従来どおり）。
  - `..._ch_<sid>_payloads.dat` … 全workerの最新結果を1行1jobで保持するaggregate snapshot
  - `..._ch_<sid>_payloads.dat.lock` … worker間のread/replaceを直列化するlock
  - `..._ch_<sid>_<ch>_input.dat` … FE→worker の入力（ユーザーがセルに入れた値など）
  - `..._ch_<sid>_<ch>_stop.flag` … FE→worker の停止指示
  - `..._ch_<sid>_<ch>_error.dat` … worker→FE のエラー通知
- **原子的 publish**: tmp に書く → 目的ファイルを `Kill` → `Name tmp As dest`。読み手が途中を見ない。
- **読みは寛容**: `Open For Input Shared`。失敗（書き込み最中など）したらそのサイクルは skip。
- セルエンコード: `N`+数値 / `S`+文字列 / `E`+空。ヘッダに `version, address, rows, cols`。

### FE 死活検知（tombstone + リースファイル）
worker は「FE がもう居ない」を検知して自分も終了する必要がある。

- **tombstone**: FE 正常終了時に `..._fe_gone_<sid>.flag` を置く。
- **リースファイル**: `HostMain` が起動時に `..._fe_lease_<sid>.lock` を
  `Open ... For Output Lock Read Write` で開き**ハンドルを保持し続ける**。
  FE がクラッシュするとこのロックは OS が即解放する。
- worker は毎周、リースファイルへの**書き込みロックを試す**。ロックできた（＝誰も保持していない）＝FE 死亡と判断。
- **なぜリースが必須だったか**: spike で「読み取り専用 `Workbooks.Open` は書き込みロックを保持しない」ことを実証。
  だから「FE がファイルを開いているか」では死活を判定できず、**専用のリースを明示的に持つ**必要があった。

### SessionId（PathFs）— hWnd を使ってはいけない
- 当初は `Application.hWnd` をセッションキーにしていたが、
  **`Workbook_Open` 時と後の呼び出しとで別の値を返す**ことがあった（後述バグ 2 の原因）。
- 現在は `PathFs.SessionId()` = `Format$(Now,"yyyymmddhhnnss") & <Timer ミリ秒>` を
  **一度だけ生成してキャッシュ**し、全チャネル / リース / done / worker パスに使う。
- 起動時 `InitAddin` で 1 日以上前の残骸ファイルを掃除。

---

## 3. 3 つのバグ（時系列と証拠）

### バグ 1: フリーズ本体
上記 §1。worker 側 VBA 対策では直らず、§2 のアーキ変更で根治。
**検証**: 実編集モード（probe 検証つき F2 を 10 秒）で FE CPU ~0%、ESC ≤3 回で回復、
タイマーが正確な経過値へジャンプ再開（旧: 回復不能）。E2E で 7 並行ジョブ + ユーザー入力 30 連続応答。

### バグ 2: worker の自殺（実インストールでのみ発現）
- `Application.hWnd` がリース生成時（`Workbook_Open`）と worker 起動時（`StartJob`）で
  **異なる値**を返し、worker に渡すリースパスが不一致に。
- → worker が起動直後に「リースが誰にも保持されていない＝FE 死亡」と**誤判定して自殺**。
- → payload が一切書かれず、画面は **"starting" / "0" のまま**（ユーザー報告のスクショと一致）。
- **なぜテストで捕まらなかったか**: テストは 250ms 間隔で COM ポーリングしており、
  その COM が `MaybePump` を誘発してバージョンが進んでしまう。**worker が死んでいても進んで見える**。
- **修正**: `Application.hWnd` 廃止 → `PathFs.SessionId()`。
- **決定的検証**: 実インストール経由で stopwatch 起動 → **COM を完全に解放** → 純粋な画面観測で
  `Seconds 114 → 120` と自走前進を確認（OnTime ポンプ単独でセルが更新される証明）。

### バグ 3: VBA コンパイルエラー「変数が定義されていません」
- `Private m_sessionId As String` を**プロシージャの後**に置いていた。
- VBA は**モジュールレベル宣言を全プロシージャより前**に置く必要がある。→ モジュール先頭へ移動して解決。

---

## 4. 対策更新: IDC_WAITカーソルのちらつき（毎秒）

### 症状と実測

従来「IDC_APPSTARTING（矢印＋spinner）」と記載していた症状名は誤りだった。desktop実使用中
（3ツール同時、Excel 8プロセス、90秒、約64Hz）の外部観測結果は次のとおり。

- **IDC_APPSTARTING: 0回**
- **IDC_WAIT: 13回**
- WAIT interval: 11〜30ms
- 同じ区間でオーナーがちらつきを目視

改修前データは `docs/measurements/2026-07-12-cursor-baseline/` に保存した。

### 確定した機構

noteで非ロック・カーソル可視・SendInput実入力・200Hz spin-waitの統制実験を行った。

| 実験 | 結果 |
|---|---|
| OnTime毎秒300ms busy loop | WAIT 56 onset / 60秒（ほぼ毎tick） |
| 同12ms | 約5秒ごとに捕捉 |
| 次回予約だけの空tick | 約5msのWAITが発生 |
| 製品Multi Stopwatch | WAITのみ、約5ms級 |

**OnTime callbackでVBAが実行されている間、ExcelはWAITを掲示する。** フラッシュ長はtick実行時間と
復帰までの遅延にほぼ対応する。空tickでも約5msの床があり完全ゼロにはできないが、5ms級は知覚
されず、15〜30ms級は見える。勝利条件はdesktop実使用でP95 < 10ms、かつオーナーが見えなく
なったと確認すること。

過去のnoteの `0 / 0 / 0` はロック画面中だった疑いが濃い。ロック中はWM_SETCURSORが発生せず
形状が更新されないため、このゼロ計測は無効とする。

### 現状の実装

`AppStateSnapshot` からcursor保存・復元を削除し、validatorも `Application.Cursor` を拒否する。
入口・出口の `DoEvents` は維持する。両方の削除はdesktopで悪化し、出口yieldは通常カーソルへの
復帰に必要という実測と整合する。

T0として、`JobPump` に256件のメモリ内リングプロファイラを追加した。`XLTOOLRACK_TEST=1` または
`XLTOOLRACK_PROFILE=1` のときだけ、tick開始Timer、callback全体、aggregate read、PumpOnce、
dispatch、error poll、cleanupの各所要msとfresh record数を記録する。定常tick中のファイルI/Oはない。

T1の生成Registry / `Infra_DispatchFlush` 基盤は実装したが、desktop A/BでtickとWAITが改善せず、
pi_raceまで変換した版は明確に悪化したためruntimeをrollbackした。現在のtoolは `OnFlush` を持たず、
`PumpOnce` もflushを呼ばない。比較値と不採用にしたT2/T3案は
`docs/measurements/2026-07-12-cursor-fix-results.md` に記録した。

カーソル観測器 `scripts/Measure-Cursor.ps1` はWAITのP50/P95/max、秒別WAIT-ms、APPSTARTING陰性対照、
開始・終了時foreground handle、CURSOR_SUPPRESSEDを記録する。配布VBAは `Declare` もcursor操作も
持たない。

---

## 5. 検証方法論（ここでハマった / 効いた）

- **FE ポンプ単独の挙動は「COM を触らない画面観測」でしか検証できない。**
  COM read は 1 回の `MaybePump` を誘発し、worker が死んでいても版が進むので**偽陽性**になる。
  → スクショ / カーソルサンプリングで観測すること。
- **カーソル計測前後にセッションが非ロックか確認する。** ロック中はWM_SETCURSORが発生せず、
  形状が更新されないまま偽の0回になる。`Measure-Cursor.ps1` は開始・終了時foreground handleを記録する。
- **CURSOR_SUPPRESSED中は計測不能。** タッチ入力後など `CURSORINFO.flags=2` のsampleが出た計測は無効。
  合意済みの無人計測でmouseを動かす場合は `SetCursorPos` でなく `SendInput` 実入力を使う。
- **`Application.OnTime` は完全修飾名で予約する。** `'<book名>'!Module.Proc` でない呼出しは無音で
  不発になることがある。
- Sleepベースのsamplingはこの環境で約64Hzが上限。より高頻度のspin-waitは合意済みの無人機だけで使う。
- onset回数だけで比較せず、WAIT intervalのP50/P95/maxを主指標にする。
- **dist の再ビルドを忘れると旧バイナリで検証**してしまう。`xltoolrack.bat` / `Install-Addin.ps1` は
  `dist` を使う。ソースを直したら dist を作り直す。
- VBA: **モジュールレベル宣言は全プロシージャより前**（バグ 3）。
- **ユーザーが作業中のマシンで Excel を前面強制 / カーソル移動する計測をしてはいけない。**
  ユーザーの入力を奪う（今セッションで実際にやってしまい叱られた）。計測は**ユーザーが席を外している合意が取れているときだけ**、
  または前面強制せずに済む方法で。

### busy 系 HRESULT の一覧（編集中 COM で出るもの）
| コード | 名前 | .NET int |
|---|---|---|
| `0x80010001` | RPC_E_CALL_REJECTED | -2147418111 |
| `0x8001010A` | RPC_E_SERVERCALL_RETRYLATER | -2147417846 |
| `0x800AC472` | VBA_E_IGNORE | -2146777998 |
| VBA 50290 | 編集モードでの書き込み拒否 | — |

`0x800AC472` は**戻り値の HRESULT**でありメッセージフィルタの対象外。だから
テストハーネスは `Invoke-XlRun`（Run 側の busy リトライ）で包む必要があった。

---

## 6. ビルド / テスト / インストール

| 操作 | コマンド |
|---|---|
| dist をビルド | `scripts\Build-Addin.ps1`（または `build.bat`） |
| テスト全実行 | `test.bat`（= `test\Run-All.ps1`、ケース 15 本） |
| インストール（FE へ登録して起動） | `scripts\Install-Addin.ps1` |
| アンインストール | `uninstall.bat` |

### ビルド検証（`scripts/Build-Addin.ps1`）が弾くもの
- 非 ASCII の VBA
- `Declare`（Win32 API 宣言）
- `Shell` / WMI / `taskkill`
- `Application.OnTime` を `@('WorkerBridge.bas','JobPump.bas')` 以外のファイルで使うこと
- ツールの厳格なシグネチャ違反（`OnFlush(ctx)` を含む）

---

## 7. 環境上の注意（重要）

- **このマシン / リポジトリは複数の AI エージェントが並行作業することがある。**
  コミット `52f8dde`（"fix: keep FE responsive during result dispatch"）は別セッションのもの。
- 深夜に外部要因で **Excel の一斉 kill** が観測されている。テスト中に FE が消えても
  コードのせいとは限らない（過去、無関係な zombie 8 個も同時に死んだ）。
- **既存の EXCEL PID には触れない。** 新規に起動したものだけを片付ける。
- `%TEMP%` にユーザー自身の実験物（`xltoolrack_gui_*`, `spike_*` 等）が残っていることがある。触らない。

---

## 8. 現在の状態（2026-07-12 時点）

- **フリーズ**: 根治済み。スイート 15 本 all passed + COM ハーネス all passed。実編集受け入れ PASS。
- **worker 自殺 (hWnd)**: 修正済み。実インストールで自走前進を画面観測で確認。
- **コンパイルエラー**: 修正済み。ユーザーが「動いた」と確認。
- **カーソルちらつき**: 症状をIDC_WAITと確定。T0リングプロファイラとT1 tool単位flushを実装し、
  stopwatchの可視書込みを同一tick最大1回へcoalesceした。desktop 90〜120秒のafter計測と
  オーナー体感確認を残す（§4）。
- **Git**: **未コミット**。main worktreeには複数作業由来の既存差分がある。
