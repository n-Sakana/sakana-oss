# TODO: AIチャット連携3ツールの手動確認

このファイルは、今回の手動確認が終わるまでリポジトリ直下に置く一時的なTODOです。

完全な操作手順： [test/MANUAL-FIXTURES.md](test/MANUAL-FIXTURES.md)

テストデータ： `C:\repos\pub\toolrack\output\manual-test-fixtures`

## 自動確認済み

- [x] 全テストが`ALL TEST FILES PASSED`
- [x] MD Mirrorの生成、解析、復元を自動検証
- [x] MD Patchの5操作、拒否例、バックアップ一致を自動検証
- [x] MD Extractの文字コードとOpenXML抽出を自動検証
- [x] 実Office形式のテストデータを生成
- [x] 7ツールのHKCU登録を検証し、右クリックメニューを復元

## 手動確認待ち

- [ ] MD Mirrorで`mirror-source`からMDを作成
- [ ] `ready-mirror.md`をファイルから復元
- [ ] `ready-mirror.md`をクリップボードから復元
- [ ] `Copy Format Instructions`のクリップボード内容を確認
- [ ] MD Patchのパス攻撃拒否を確認
- [ ] MD PatchのSHA-256不一致拒否を確認
- [ ] MD Patchの変更前行不一致拒否を確認
- [ ] `apply-demo.md`をファイルから正常適用
- [ ] 正常適用後のバックアップ内容を確認
- [ ] テストデータを初期化
- [ ] `apply-demo.md`をクリップボードから正常適用
- [ ] MD ExtractのDefault出力を確認
- [ ] PDFがDeferredになった場合の表示と出力を確認
- [ ] 必要ならTimeout 300またはCustomを確認

## 完了後に削除するもの

手動確認がすべて終わったら、次を削除します。

- この`TODO-MANUAL-TEST.md`
- `output\manual-test-fixtures`
- 今回の手動確認で作られた`output\md-mirror_*`
- 今回の手動確認で作られた`output\backup_*`
- 今回の手動確認で作られた`output\extract-md_*`
- 今回の手動確認で作られた`log\extract-md_*`

`test\MANUAL-FIXTURES.md`、生成スクリプト、自動テストは今後も使うため削除しません。
