# Slack text reactions 20260701 v2

Slack用の文字ベース絵文字セットです。

- 128x128px PNG
- 透過背景
- 1ファイル 128KB 以下
- 縁取り・シャドウなしのフラット色文字
- 基本は極太ゴシック
- 1〜3文字は横一列、3文字は横方向に圧縮して高さを確保
- 4文字は2×2配置
- 余白は最小化
- 高解像度で描画してから縮小し、ラスタ消し込みは使わない

## Files

- `:youkoso:` — ようこそ
- `:kansha:` — 感謝
- `:kansha_kangeki:` — 感謝感激
- `:sasuga:` — さすが
- `:sugosugi:` — すごすぎ
- `:tensai:` — 天才
- `:nice:` — ナイス
- `:kakusan:` — 拡散
- `:uwaa:` — うわぁ
- `:majika:` — マジか
- `:saikyou:` — 最強
- `:komatta:` — 困った
- `:kanben_shite:` — 勘弁して
- `:onajiku:` — 同じく
- `:sorena:` — それな
- `:doui:` — 同意
- `:urayama:` — うらやま
- `:yakekuso:` — やけくそ
- `:namusan:` — 南無三
- `:ouen:` — 応援
- `:donmai:` — どんまい
- `:daishouri:` — 大勝利
- `:iwai:` — 祝
- `:noroi:` — 呪
- `:kami:` — 神
- `:muri:` — 無理
- `:tsurai:` — つらい

## Generation

`generate.py` で再生成できます。Pillow と日本語ゴシックフォントが必要です。
