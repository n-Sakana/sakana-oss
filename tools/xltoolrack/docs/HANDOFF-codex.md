# 引き継ぎ書: xltoolrack のカーソルちらつき残課題

更新: 2026-07-12

最終判断と実装順は `docs/cursor-flicker-fix-directive.md` を正とする。凍結根治を含む
アーキテクチャ全体は `docs/edit-freeze-and-fe-pump.md` を参照すること。

## 1. 確定した診断

セル編集中に FE Excel 全体が固まる問題は、worker から FE への同期 COM 書込みを廃し、
ファイルチャネルを FE の `Application.OnTime` で pull する構成へ反転して解消した。

残っていたカーソルちらつきは **IDC_WAIT（青い輪）** であり、従来記載していた
IDC_APPSTARTING ではない。desktop 実使用中の90秒計測では IDC_APPSTARTING は0回、
IDC_WAIT は13回、1回11〜30msだった。同区間のちらつきをオーナーも目視している。

note の統制実験により、OnTime callback の実行中は Excel が WAIT を掲示し、フラッシュ長が
tick実行時間にほぼ対応することが確定した。次回予約だけの空tickでも約5msの床があるため、
1Hzポンプを維持したまま出現を完全にゼロにはできない。一方、約5msは知覚されず、15〜30msは
見える。したがって勝利条件は次のとおり。

> desktop実使用（3ツール同時）で WAITフラッシュ長 P95 < 10ms、かつオーナーが
> 「見えなくなった」と確認する。

`AppStateSnapshot` が callback 中の一時的なbusy cursor modeを保存・復元する別バグは削除済み。
配布add-inは `Application.Cursor` を変更せず、Excelの文脈カーソルをExcel自身に管理させる。

## 2. 現在の callback

```text
JobPump.Pump_Tick
  DoEvents
  HostServices.Jobs().PumpOnce
    Application.EnableEvents = False
    aggregate file を1回 open/decode
    fresh result ごとに memory cache を更新
    Infra_DispatchResult
      tool.OnResult
        stopwatch / pi_race は従来どおりjobごとに表示行を書込み
    error/stale/terminal の必要な確認
    cursor以外のapplication stateを復元
    Application.EnableEventsを復元
  必要なら次のOnTimeを予約
  DoEvents
```

入口・出口の `DoEvents` は維持する。両方を削除した版はdesktopで明確に悪化し、特に出口yieldを
失うと通常カーソルへの復帰が遅れるという機構と整合する。

## 3. T0 tickプロファイラ

`JobPump` は256件の固定長リングに、tick開始時の `Timer`、callback全体の所要ms、fresh record数、
aggregate read、PumpOnce、dispatch、error poll、cleanupの各所要msを保持できる。通常配布時は無効で、
次のどちらかをExcel起動前に設定した場合だけ記録する。

- `XLTOOLRACK_TEST=1`
- `XLTOOLRACK_PROFILE=1`

定常tick中のファイルI/Oはない。`JobTest_ResetPumpProfile` と `JobTest_PumpProfile` がtest用の
reset/dump入口で、dumpはCSV文字列として返す。

## 4. T1 可視sheet書込みのcoalesce（実測不採用・runtime rollback済み）

ビルド生成器はオプショナルな `Public Sub OnFlush(ByVal ctx As InfraContext)` を検出・検証し、
`Registry_HasFlush` / `Registry_OnFlush` を生成する。`OnFlush` は `OnResult` を持つtoolでのみ許可する。

生成器と `Infra_DispatchFlush` のオプショナルhook基盤は残しているが、現在の `PumpOnce` は呼ばず、
対応toolも0件である。stopwatch / pi_raceは従来どおり `OnResult` 内でjobごとに表示行を書込む。

desktop 3ツール同時の90秒A/Bでは、stopwatchだけcoalesceした版のtick P95が15.625ms、
stopwatch+pi_race版が19.531ms、全rollback版が15.625msだった。pi_raceを狙いどおり3 result→1 flushへ
変換してもWAIT P95は16.9→32.5msへ悪化し、オーナー体感も「明らかに悪化」だったため不採用とした。
詳細rawと比較表は `docs/measurements/2026-07-12-cursor-fix-results.md` を参照すること。

## 5. 計測プロトコル

観測には `scripts/Measure-Cursor.ps1 -Seconds 90 -OutDir <path>` を使う。観測器はExcelへCOM接触せず、
WAITのonset、各interval、P50/P95/max、秒別WAIT-msを記録する。APPSTARTINGは陰性対照として記録する。
開始・終了時のforeground window handle、CURSOR_SUPPRESSED sample数もsummaryへ残す。

before/afterはdesktopでオーナーの通常使用中、3ツール同時、90〜120秒で測る。主指標はWAIT interval
のP95であり、onset回数だけでは判定しない。ユーザー作業中に前面強制やmouse移動は行わない。
実Excelのpump ringとcursorを同一区間で回収する場合は `scripts/Measure-Desktop.ps1` を使う。

過去のnoteの `0 / 0 / 0` は、noteがロック画面だった疑いが濃い。ロック中はWM_SETCURSORが
発生せず形状が更新されないため、このゼロ計測は解消の証拠として無効とする。

## 6. S1〜S10累積matrix（不要）

旧レビュー文書のS1〜S10は、2026-07-12の統制実験で「OnTime callback中のWAIT長 ≒ tick長」と
確定したため、以後は実施しない。T1後もP95が10ms以上なら、T0プロファイルで支配成分を特定して
からT2/T3へ進む。aggregate readの跳ねがなければチャネルpathやDefender設定は変更しない。

## 7. 検証上の地雷

- セッションロック中はカーソル形状が更新されない。計測の開始・終了時にforeground状態を記録する。
- `Application.OnTime` は `'<book名>'!Module.Proc` の完全修飾名で予約する。非修飾名は無音で不発になる。
- `CURSORINFO.flags=2`（CURSOR_SUPPRESSED）中の計測は無効。必要なmouse移動は、合意済みの無人計測で
  `SendInput` 実入力を使い、`SetCursorPos` は使わない。
- Sleepベースの観測はこの環境で約64Hz。より高頻度のspin-waitは合意済みの無人機でだけ使う。
- sampling中にFE COMを触らない。`MaybePump` を誘発して測定系を変える。
- onset数だけで比較しない。P50/P95とmaxを主に見る。
- `Application.Cursor`、`ScreenUpdating` toggle、入口・出口 `DoEvents` の削除を再投入しない。
- 配布コードの `Declare` / Shell / WMI / `taskkill` 禁止を守る。
- 既存Excel PIDを終了しない。自分が生成したPIDだけを所有関係確認後に片付ける。
- source変更後はdistを再build/installし、SHA-256で実体を確認する。

## 8. rollback基準と既知良好版

T1はdesktop P95とオーナー受入れの両方でFAILし、runtimeをrollback済み。T2のdispatch object遅延生成、
T3のchannel専用サブフォルダ、life payload圧縮も各々1変数で測ったが改善せずrollbackした。
最終候補は15ケースgreen（終了コード0）。ただし `test-ribbon-ui` のassertion後の終了時に、この環境固有の
VBA「ファイルが見つかりません」modalを外部から閉じ、test所有Excelを終了する介入が1回必要だった。
全出力は `docs/measurements/2026-07-12-full-suite.log` にある。

勝利条件P95 < 10msと「見えなくなった」は未達。次はT4であり、tick 1s→2sまたはcursor pinの
どちらもオーナー承認なしに実施しない。

改修前の既知良好ビルドSHA-256:
`0AE109F5F7BE0D904AE6DCAA424769CA7E6843B86A7EB864D4675C0199909746`

main worktreeは既存の未コミット変更を含む。reset / clean / checkoutで他人の変更を消さないこと。
