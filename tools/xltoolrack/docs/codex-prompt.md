# Codex 用プロンプト（そのままコピペして渡す）

> 2026-07-12 Codex 対応更新: cursor固定案は撤回し、ポンプからcursor操作を完全除去。
> worker側aggregate snapshot、memory直結、steady-state sweep除去でfresh tickを
> 約17.7ms→約9.3msへ短縮。`Declare`なし、全15ケース一括green。
> 以下は対応前の履歴として残す。

以下をそのまま Codex に渡してください。

---

あなたは `C:\repos\pub\xltoolrack`（Excel の FE-BE アドイン）で作業します。
まず `docs/edit-freeze-and-fe-pump.md` と `docs/HANDOFF-codex.md` を読んでから着手してください。

## 背景（要約）
「タイマー実行中にセルを編集すると FE 全体が固まる」問題は前任が根治済み。
worker（別プロセス）→ FE への同期 COM 書き込みを廃止し、worker はファイルに publish、
FE が `Application.OnTime` の 1 秒ポンプで pull する方式（案 A）に変えた。
これによりフリーズは消えたが、**FE ポンプの tick が毎秒メインスレッドを一瞬ビジーにするため、
Excel が前面のときにマウスカーソルが毎秒「考え中」(IDC_APPSTARTING) に一瞬なる**という残課題がある。

前任が `Application.Cursor` のピン（tick 中 `xlNorthwestArrow` / tick 後 `xlDefault`）を入れ、
実測でちらつきを 12 秒あたり 8 回 → 2〜4 回に減らしたが、**ゼロにはできていない**。

## あなたのゴール
フリーズ根治を壊さずに、この「考え中」カーソルのちらつきを**さらに減らす（できればゼロ化）**。
不可能なら、その根拠を計測で示し、**妥協ライン or 設計変更（下記制約の緩和）をユーザーに提案**する。

## 厳守する制約
- **`Declare`（Win32 API 宣言）禁止**。`scripts/Build-Addin.ps1` の validator が弾く。
  これを緩和したい場合は**実装前にユーザーへ承認を求める**（`SetTimer` は候補だが、WM_TIMER も
  メインスレッド処理なので効くとは限らない点も伝える）。
- **非 ASCII の VBA 禁止**、`Shell`/WMI/`taskkill` 禁止、`Application.OnTime` は
  `WorkerBridge.bas` と `JobPump.bas` でのみ許可。
- **worker → FE の COM 呼び出しを絶対に復活させない**（それがフリーズの原因だった）。

## 検証のしかた（重要・ここでハマりやすい）
- 変更後は必ず `test.bat`（スイート 15 本）が緑であることを確認。
- **カーソルちらつきは Excel が前面のときだけ再現する。背面の自動化では観測できない。**
  計測は前面強制 + カーソルをグリッド上に置く必要があり、**ユーザーの操作を奪う**。
  → **ユーザーが席を外している合意があるときだけ**実施する。合意がなければ計測せず、
     コード上の理屈と `test.bat` 緑で説明し、体感確認はユーザーに依頼する。
- 参考計測スクリプト:
  `C:\Users\ynisi\AppData\Local\Temp\claude\C--Users-ynisi\f641ca14-4ced-4bd1-8caf-30d61eb19c9c\scratchpad\measure-cursor2.ps1`
  （test-host 経由、前面強制 + 50Hz サンプリング、A/B は `JobHost.PumpOnce` のピン 2 行を出し入れ）
- **COM read で pump の動作を検証しない**（`MaybePump` が誘発され偽陽性になる）。画面/カーソル観測で。

## 試す順（前任の推奨）
1. `JobHost.PumpOnce` 1 tick のメインスレッド占有時間を `Timer` 差分で計測。
   `Infra_DispatchResult` が重い疑い。占有を数十 ms 未満に抑えられればビジー表示が消える可能性。
2. payload 反映（`Range(address).Value2 = values`）を差分書き込みに。
3. `DoEvents` を tick 冒頭に 1 回（再入注意、`m_pumping`/`EnableEvents=False` との相互作用を確認）。
4. どうしても消えないなら計測結果を添えて「6〜7 割減で妥協」or「`Declare` 緩和」をユーザーに提案。

## 進め方
- 修正は 1 つずつ、`test.bat` 緑を保ちながら。ソース修正後は `scripts\Build-Addin.ps1` で dist を作り直す
  （`xltoolrack.bat`/`Install-Addin.ps1` は dist を使うため、忘れると旧バイナリで検証してしまう）。
- **既存の EXCEL プロセスに触れない**（この repo は複数エージェント並行、深夜に外部 Excel kill あり）。
- コミットはユーザーの明示指示があるまでしない（現在 24 変更 + 3 新規が未コミット）。
- 判断に迷う設計事項（`Declare` 緩和・妥協ライン）は**勝手に決めずユーザーに聞く**。

---
