# TODO: グローバル起動とCaptureの手動確認

自動テストと実HKCU検証は完了済みです。このファイルは、目視・実操作と職場PCのEDR/GPO確認が
終わるまで残します。

## 自動確認済み

- [x] `test\run-tests.ps1`を一度の通し実行で`ALL TEST FILES PASSED`
- [x] installで11ツール、7メニューページ、HKCU Runを登録
- [x] Host ready、active binding 3、rejected 0、hotkey 2、mouse 1
- [x] Ctrl+右クリック判定のcallback p99 5ms未満
- [x] Capture実描画20回: 初回800ms以下、同一session p95 500ms以下
- [x] Capture light/dark、100/150/200% DPI、負座標、キーボード状態を自動検証
- [x] install失敗時の旧Host・active世代・Run・メニューrollbackを実HKCUで検証

## 手動チェックリスト

- [ ] 1. Explorerの空白を通常右クリックすると従来メニュー、Ctrl+右クリックではCaptureだけが開く
- [ ] 2. Chrome、VS Code、メモ帳、利用可能なOfficeアプリでCtrl+右クリックすると同じCaptureがマウス付近へ開く
- [ ] 3. 上記アプリで修飾なし右クリックが元どおり完全に動作する
- [ ] 4. Ctrl+Alt+CでCapture、Ctrl+Alt+TでTranscribeが黒いコンソールなしで開く
- [ ] 5. `bindings.json`の正常な変更が自動反映され、不正JSON中は直前のbindingが動き続けて理由が通知される
- [ ] 6. 他アプリと同じglobal hotkeyを競合させた場合、そのbindingだけinactiveになり理由が通知される
- [ ] 7. Captureをlight/dark双方で確認し、hover、Tab、矢印、Enter、Esc、外側クリックを試す
- [ ] 8. Windows表示倍率100/150/200%と複数モニター端でCaptureが欠けない
- [ ] 9. 範囲／ウィンドウ × 画像／パス／テキストの6操作を各1回確認する
- [ ] 10. HostをshutdownするとCtrl+右クリックを含む予約操作が元に戻り、installで復帰する
- [ ] 11. サインアウト／サインイン後、黒いコンソールなしでHostが自動起動しstatusがreadyになる
- [ ] 12. uninstall後、Host、HKCU Run、Tool Rackメニュー、ローカルbootstrap/stateが消える
- [ ] 13. 職場PCで次のprobeを実行し、サインイン後もEDR/GPOに遮断されず、最後にprobeを削除できる

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\test\probe-global-host.ps1 -Install
# サインアウトして再度サインイン
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\test\probe-global-host.ps1 -Verify
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\test\probe-global-host.ps1 -Latency
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\test\probe-global-host.ps1 -Remove
```

## Captureプレビュー

実装から生成したlight/dark PNGは一時フォルダの次のファイルです。

- `%TEMP%\toolrack-capture-preview-light.png`
- `%TEMP%\toolrack-capture-preview-dark.png`

## 完了後

全項目を確認したら、このファイルを削除します。既存の`TODO-MANUAL-TEST.md`はAIチャット連携3ツールの
別チェックリストなので、その確認が終わるまでは残します。
