# Dataverse Stack Overview

Web通信プロトコルから Microsoft Power Platform / Dataverse まで、フロント・バックエンド・認証・ホスティングを通しで整理した学習メモ。下から上に積み上げる順序で読むと理解しやすい。

---

## 1. 通信プロトコルの階層

「通信の決まりごと」のこと。お互い何のルールで喋るかの取り決め。

```
アプリ層: HTTP, HTTPS, SSH, SMTP, DNS, WebSocket
↓
暗号化層: TLS（HTTPSやWebSocket wssに被さる）
↓
トランスポート層: TCP, UDP
↓
ネットワーク層: IP
```

下の層は上の層を「運ぶ」役割で、上の層は下の層に「乗っかってる」。WebSocketも最初はHTTPで始まって途中でアップグレードできるのは、上の層同士の中で切り替えてるから。

### 主要プロトコル

| プロトコル | 役割 | 主な用途 |
|---|---|---|
| **HTTP** | Webの基本通信 | API、Web閲覧 |
| **HTTPS** | HTTP + TLS暗号化 | 現代Web標準 |
| **SSH** | リモートに安全に入る | `ssh vps`、GitHub接続 |
| **SFTP / SCP** | SSHの上でファイル送る | VPSへのアップロード |
| **DNS** | ドメイン名→IPアドレス | 電話帳的役割 |
| **TLS** | 暗号化のお作法 | HTTPSやWebSocket wssに被さる |
| **TCP** | 信頼性重視、順序保証 | Webアプリ標準 |
| **UDP** | 速度優先 | ゲーム、動画ストリーミング |
| **SMTP / IMAP / POP3** | メール用 | 送受信 |

---

## 2. HTTPの中の世界

### HTTPメソッド

HTTPは通信プロトコルで、その中に「動詞」がある。

- **GET** = 「取ってきて」（読み）
- **POST** = 「これを送るよ」（書き）
- **PUT** = 更新
- **DELETE** = 削除

### REST

HTTPの「使い方の作法・流派」。プロトコルそのものじゃない。

「リソース（データ）にはURLを割り当てて、GETで取得、POSTで作成、PUTで更新、DELETEで削除する」って綺麗な書き方ルール。REST APIはその作法に従って作られたAPIのこと。

### WebSocket

HTTPとは別の通信プロトコル。「電話線張りっぱなし」型で、一度繋いだら双方向にいつでも喋れる。

最初はHTTPで接続を始めて、途中で「WebSocketに切り替えるよ」ってアップグレードする仕組み。だから既存のWeb基盤（80番、443番ポート、Cloudflare Tunnel等）にそのまま乗っかれる。

### REST vs WebSocket

| 観点 | REST (HTTP) | WebSocket |
|---|---|---|
| 通信 | 手紙型、毎回繋ぎ直し | 電話型、張りっぱなし |
| 方向 | クライアント→サーバ | 双方向 |
| リアルタイム性 | 弱い（ポーリング要） | 強い |
| 実装の重さ | 軽い | 接続維持が要る |
| 典型例 | API呼び出し、Web閲覧 | チャット、ライブ更新 |

```
プロトコル
├─ HTTP ─── REST APIで使う、毎回手紙
│   ├─ GET (読む)
│   ├─ POST (送る)
│   └─ ...
└─ WebSocket ─── 張りっぱなし、双方向
```

---

## 3. 認証の世界 (OAuth)

通信できるようになった次は「誰が通信していいか」。

### OAuth (オーオース)

Open Authorizationの略。「パスワードを渡さずに、別のサービスに自分の権限を貸す」仕組み。

例えるとホテルのルームキー。フロントにマスターキー（パスワード）は渡さないで、「302号室に入っていいよ」って限定的なカードキー（トークン）だけ渡す。失くしても鍵を変えれば済むし、有効期限も付けられる。

### APIキーとOAuthの違い

| 観点 | APIキー | OAuth |
|---|---|---|
| 認証する相手 | アプリ | 人間（経由でアプリ） |
| 期限 | なし or 長期 | 短期（1h）+ refresh |
| 権限粒度 | キー単位 | スコープごとに細かく |
| ユーザー識別 | 不可 | 可能 |
| 漏洩リスク | 致命的 | 範囲限定で被害抑制 |
| 実装 | 1行で済む | 同意フロー実装が要る |
| 典型 | 公開データ取得 | 個人データ操作 |

GoogleのコンソールにAPIキーとOAuthの両方があるのは、「誰でも使える公開データ」と「ユーザー個人のプライベートデータ」の両方をサービスが提供してるから。

- YouTube Data API（公開動画情報） → APIキーで十分
- Google Maps → APIキー
- Googleドライブ（個人ファイル） → OAuth必須
- Gmail → OAuth必須

### OAuthのフロー

OAuthには複数のフローがあって、用途で使い分ける。

**Authorization Code Flow（人間が主役）**

```
人間（Entra IDアカウント持ち）
↓ アプリを使い始める
アプリ「この人の代わりにアクセスしていいですか？」
↓ ブラウザで Entra ID にリダイレクト
人間「ログイン+同意します」
↓
Entra ID「OK、アクセストークン発行」
↓
アプリが人間の権限でAPIを叩く
```

権限は人間のセキュリティロールで決まる。アプリは単なる代理人。

向いてる用途: Power Apps、SPA、人間が操作するWebアプリ

**Client Credentials Flow（アプリが主役）**

```
バックエンドアプリ（人間ログイン要らない）
↓ 自分のClient ID + Secret 提示
Entra ID「OK、アクセストークン発行」
↓
アプリ自身の権限でAPIを叩く
```

権限はApplication Userのセキュリティロールで決まる。

向いてる用途: 定期バッチ、サーバー間連携、自前バックエンドからの自動処理

**その他のフロー**

- Device Code Flow … スマートTV等、入力しづらいデバイス用
- Implicit Flow … 古い、今は非推奨

「OAuth」って単語だけだと2.0のことを指すのが今標準。

---

## 4. Microsoft 365 の階層構造

Power Platform に入る前に押さえておく前提。

```
[ テナント ]                ← Microsoft 365契約全体
└ 会社全体の器
├ [ 環境 A: 開発用 ]
├ [ 環境 B: 本番用 ]
├ [ 環境 C: Default ]
└ [ 環境 D: テスト用 ]
```

### テナント

Microsoft 365 契約ごとに1つ作られる、会社全体の器。

- Entra IDのユーザー一覧はテナント内で管理される
- ドメインも紐付く（`@yourcompany.onmicrosoft.com` 等）
- **Tenant ID**（GUID）で一意に識別される
- 1社1テナントが基本

例えるなら会社の建物全体。入口の警備員（Entra ID）が「この建物に入っていい人」を管理してる。

### 環境

テナントの中に作る、Power Platform の作業空間。

- 1テナント内に複数作れる（開発・検証・本番・個人サンドボックス）
- 各環境に独立した Dataverse が入る
- Power Apps / Power Automate / Power BI も環境単位で管理
- **Environment ID** で識別

例えるなら建物の中のフロア。用途別に分けられた区画。

| 観点 | テナント | 環境 |
|---|---|---|
| 単位 | Microsoft 365契約全体 | Power Platform作業空間 |
| 数 | 1社で基本1つ | 1テナント内に複数 |
| 管理対象 | ユーザー、ドメイン、ライセンス | Dataverse、アプリ、フロー |
| URL例 | `yourcompany.onmicrosoft.com` | `org12345.crm.dynamics.com` |
| 主な管理画面 | Azure Portal、M365 admin center | Power Platform Admin Center |

**Default環境**: Microsoft 365契約すると自動で作られる共有環境。個人がちょっと試すには楽だが、業務利用は別途専用環境を作るのが推奨。

---

## 5. Power Platform の全体像

Power Platform は「製品」というより複数のサービス群をまとめた箱・総称。

```
Power Platform（傘）
├ Power Apps      ← 画面作るとこ
├ Power Automate  ← 自動処理を組むとこ
├ Power BI        ← データ可視化
├ Power Pages     ← 外部公開Webサイト
├ Copilot Studio  ← AIチャットボット
└ 共通基盤
├ Dataverse   ← データの倉庫
├ Connectors  ← 各種SaaS接続口
└ AI Builder  ← AI部品
```

### 各サービスの役割

**Power Apps** — 画面を作って人間が操作する役割
- Canvas Apps: GUIで自由レイアウト、Excel数式系の言語
- Model-driven Apps: Dataverseスキーマから自動生成される定型アプリ
- Code Apps: 2026年2月GAのReact製フルコード版

**Power Automate** — 人間が触らなくても勝手に処理が走る役割
- 「Dataverseに行が追加されたらTeamsに通知してExcelに追記」のようなフロー
- スケジュール実行（毎朝9時に処理）
- Power Appsから呼び出されて走るフロー
- 旧名 Microsoft Flow

**Dataverse** — 共通データ倉庫
- 各サービスがここを参照・更新する

| 名前 | 役割 | 例え |
|---|---|---|
| Power Platform | 全体の傘 | デパート全体 |
| Power Apps | 画面作る | 売り場フロア |
| Power Automate | 自動処理 | バックヤード |
| Power BI | 可視化 | 分析室 |
| Dataverse | データ保管 | 倉庫 |

---

## 6. Entra ID と Dataverse の 2 世界構造

ここがいちばん繋がりにくいポイント。認証世界と権限世界が分かれてる。

```
[ Entra ID（旧 Azure AD）]  ← 認証の世界
└ 「あなた誰？」を管理する身分証発行所

[ Dataverse ]  ← 権限とデータの世界
└ 「あなた何できる？」を管理する建物
```

アプリで業務やる時って、この2つの世界を行き来する。Entra IDで身分証もらって、Dataverseの入館証と権限が付いて、ようやくデータ触れる。

### 4 つの要素

| 要素 | どこの世界 | 役割 | 例え |
|---|---|---|---|
| **App Registration** | Entra ID | アプリ登録、Client ID/Secret発行 | 身分証の発行申請 |
| **Service Principal** | Entra ID | アプリのテナント内実体 | テナント支店の身分証 |
| **Application User** | Dataverse | アプリを「ユーザー」として認識 | 建物の入館証 |
| **Security Role** | Dataverse | 何ができるかの権限束 | フロアごとの鍵 |

### 自前バックエンドから Dataverse を叩く流れ (Client Credentials Flow)

**ステップ1: Entra ID で App Registration**
- Azure Portal で「App Registrations → 新規登録」
- **Client ID** 発行
- 必要なら **Client Secret** か Certificate 発行

**ステップ2: Service Principal が自動生成**
- App Registration を作ると同じテナント内に Service Principal が自動で生まれる
- App Registration が定義、Service Principal が実体

**ステップ3: Dataverse 側で Application User として登録**

一番見落とされるポイント。Entra ID 側のアプリは生まれても、Dataverse はまだそのアプリを知らない。

- Power Platform Admin Center → 環境を選ぶ
- 「設定 → ユーザー + アクセス許可 → アプリケーション ユーザー」
- 「新しいアプリ ユーザー」追加
- Client ID を貼り付けて、Entra ID 側の App Registration と紐付ける

**ステップ4: Security Role を割り当てる**

Application User 作っただけでは何の権限もない。

- System Administrator: 全権限
- System Customizer: スキーマ変更可能
- Sales Manager / Salesperson: 業務テンプレロール
- カスタムロール: 必要な権限だけ束ねた自前ロール

最小権限の原則で、必要なエンティティの必要な操作だけ付けるのが筋いい。

**ステップ5: アプリから API を叩く**

```python
# 1. Entra IDからアクセストークンを取得
token = requests.post(
    f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token",
    data={
        "grant_type": "client_credentials",
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope": f"https://<org>.crm.dynamics.com/.default"
    }
).json()["access_token"]

# 2. そのトークンでDataverseを叩く
requests.post(
    f"https://<org>.crm.dynamics.com/api/data/v9.2/accounts",
    headers={"Authorization": f"Bearer {token}"},
    json={"name": "新規顧客"}
)
```

裏で起きてること:

```
アプリ
↓ Client ID + Secret 提示
Entra ID（身分証発行所）
↓ アクセストークン発行
↓ トークン持参
Dataverse（建物の受付）
↓ トークン検証 → Application User 特定
↓ Security Role で権限確認
↓ OK → レコード作成
```

### 人間が Power Apps から Dataverse を触る時 (Authorization Code Flow)

- Application User は出てこない
- 人間自身の System User と Security Role で動く
- 認証は裏で Microsoft が自動でやってくれる
- 開発者は意識しない

### 階層イメージ

```
[Entra ID 世界]
App Registration「My Backend」
├ Client ID: abc123...
├ Client Secret: ...（秘密）
└ Service Principal（実体）
│
│ Client ID で紐付け
▼
[Dataverse 世界]
Application User「My Backend」
├ systemuser テーブルに1行
└ Security Roles:
├ My Backend Role（カスタム）
│   ├ account: Create / Read / Update
│   └ contact: Read
└ ...
```

### よく詰まるポイント

1. **App Registration 作っただけで動くと思ってる** … Entra ID と Dataverse は別世界。Application User 作り忘れで 401 になる
2. **Security Role 不足で 403** … ロールに必要なエンティティ権限が足りない
3. **スコープ指定間違い** … `https://<org>.crm.dynamics.com/.default` であって `api/data/v9.2/.default` ではない
4. **Client Secret 失効** … 有効期限あり。長期運用なら Certificate 認証が筋いい
5. **マルチテナントでの Service Principal** … 別テナントの Dataverse に繋ぐ時、同意プロセスが必要

### マルチ環境での注意

1つのアプリで複数環境の Dataverse に繋ぐなら、**各環境ごとに Application User を別々に登録**する必要がある。Dataverse は環境単位で別物だから。

---

## 7. フロント開発の 4 分類

Dataverse をバックエンドにする時のフロント実装は、ローコード度合いで 4 層に分かれる。

```
ローコード度 ←──────────────→ プロコード度

(1) Canvas Apps / Model-driven    ローコード、Power Platform内
(2) Canvas Apps + PCF             基本ローコード、画面の一部だけHTML差し込み
(3) Code Apps                     フルReact、でもPower Platform内に居続ける
(4) 完全独立 React + Web API      Power Platform外、自由なホスティング
```

### 通信観点の整理

| 層 | FE⇄BE(Dataverse)の繋ぎ方 | 認証 |
|---|---|---|
| (1) Canvas/Model-driven | Dataverseコネクタ（裏でWeb API） | 自動 |
| (2) Canvas + PCF | (1)と同じ。コネクタ経由 | 自動 |
| (3) Code Apps | Power Platform SDK経由でコネクタ呼ぶ | 自動 |
| (4) 完全独立React | Web APIを直接fetch | OAuth自前(MSAL.js等) |

(1)〜(3) は Power Platform の認証経路に乗っかれる、(4) だけ自前で認証実装が要る。

### Power Apps Component Framework (PCF)

- **TypeScript + HTML + CSS** でカスタムコンポーネントを作る
- Reactも使える（React版テンプレートが推奨）
- Canvas Apps / Model-driven Apps に埋め込める
- 標準コントロールでは表現できない凝った UI 用

開発フロー:
```
pac pcf init       # プロジェクト生成
↓ TypeScript で実装
↓ pac pcf push でDataverseにデプロイ
↓ Power Apps 内で配置
```

**配布の許可**
- Canvas Apps で使うには、環境管理者が「Power Apps component framework for canvas apps」設定をオン
- Model-driven Apps はこの設定不要
- コンポーネント自体の import は System Customizer 相当の権限が要る

### Code Apps (2026 年 2 月 GA)

- 完全にコードで書く Power Apps
- TypeScript + React + Vite
- ローコードと決別、Pro Code 開発者向け
- Power Platform の恩恵（コネクタ、認証、ガバナンス、ライセンス）はそのまま使える

開発フロー:
```
pac code init      # Code Appプロジェクト生成
↓ Reactで普通にフロント実装
↓ Power Platformコネクタを呼び出すSDK経由でデータアクセス
↓ pac code push でデプロイ
```

**承認の段数**
- 環境管理者: Power Platform Admin Center で Code Apps を有効化
- ライセンス: 利用者全員に Power Apps Premium 必須

### PCF と Code Apps の比較

| 観点 | PCF | Code Apps |
|---|---|---|
| 範囲 | 画面の一部だけ | アプリ全体 |
| 親アプリ | Canvas / Model-driven | 単独で動く |
| 用途 | 既存アプリの拡張 | フルコード開発 |
| 成熟度 | GA、安定 | 2026/2 GA、まだ若い |
| 学習コスト | 中 | 高（Reactに慣れてれば楽） |

### Power Pages

別軸だが触れておく。

- 外部公開 Web サイト用
- Liquid テンプレートと HTML/CSS/JS で構築
- Dataverse 連携、認証も内蔵
- 顧客ポータルや申請フォーム向け

---

## 8. Dataverse の通信手段 (網羅)

**WebSocket は存在しない**。Dataverse は業務データ台帳で、リアルタイム双方向通信は設計上想定外。

通信方向で 2 つに分けて整理。

### インバウンド (外部 → Dataverse)

| # | 手段 | 特徴 |
|---|---|---|
| 1 | **Web API** (REST/OData v4) | メイン、HTTPS + JSON、OAuth 2.0 |
| 2 | **Organization Service** (SOAP) | 古い、.NET SDK、トランザクション処理に強い |
| 3 | **Dataverse Connector** | Power Platform内、ノーコード、裏でWeb API |
| 4 | **Custom Connectors** | 外部APIをPower Platform内で登録、OpenAPI仕様 |
| 5 | **Dataflows / Data Integrator** | ETL、定期バルク取込、Power Query |
| 6 | **Dual-write** | Dynamics 365 F&O 専用双方向同期 |
| 7 | **Bulk Data Load** | `$batch`、Synapse Link、Data Import Wizard |
| 8 | **Virtual Tables** | 外部データソースを Dataverse として見せかける |

### アウトバウンド (Dataverse → 外部)

| # | 手段 | 特徴 |
|---|---|---|
| 9 | **Webhooks** | HTTP POSTで通知、軽い、Plug-in Registration Tool |
| 10 | **Azure Service Bus** | メッセージキュー、AMQP、大規模配信 |
| 11 | **Azure Event Hubs** | 大量イベントストリーミング、Kafka互換 |
| 12 | **Plug-ins** | .NETコードを内部に差し込み、同期/非同期 |
| 13 | **Power Automate** | ノーコード、Dataverse トリガー、数百のコネクタ |
| 14 | **Change Tracking API** | 差分ポーリング、pull型、Webhook補完 |
| 15 | **Dataverse Search Integration** | Azure Cognitive Search 連携 |

### 周辺・派生

| # | 手段 | 特徴 |
|---|---|---|
| 16 | **Power Pages** | 外部公開Webサイト経由 |
| 17 | **Dataverse for Teams** | Teams内軽量版、通常Dataverseとは別環境 |

### 認証手段 (全方式共通)

| フロー | 用途 |
|---|---|
| OAuth 2.0 Authorization Code Flow | ユーザー対話あり、ブラウザでログイン |
| Client Credentials Flow (Service Principal) | アプリ専用、バックエンド連携で標準 |
| Managed Identity | Azureリソース内から、Azure VM/Function等で楽 |
| Certificate-based | 証明書認証、Service Principalの強化版 |

### プロトコル・フォーマット階層

```
┌─────────────────────────────────────┐
│ Web API: REST + JSON + OData v4     │
│ Organization Service: SOAP + XML    │
│ Service Bus: AMQP / HTTPS           │
│ Event Hubs: AMQP / HTTPS / Kafka    │
│ Webhooks: HTTP POST + JSON          │
└─────────────────────────────────────┘
↓ すべて TLS 1.2+
┌─────────────────────────────────────┐
│        HTTPS / TCP                  │
└─────────────────────────────────────┘
```

### 用途別マトリクス

| やりたいこと | 推奨手段 |
|---|---|
| 自前バックエンドから読み書き | Web API |
| Power Platform内で完結 | Dataverse Connector |
| 大量バッチ取込 | Dataflows / `$batch` |
| 外部DBを見せかける | Virtual Tables |
| 即時通知（軽い） | Webhooks |
| 即時通知（確実性） | Plug-in (sync) |
| 大規模配信 | Service Bus |
| 大量ストリーム | Event Hubs |
| ノーコード連携 | Power Automate |
| 差分定期取得 | Change Tracking API |
| Dynamics F&O同期 | Dual-write |

### リアルタイム双方向の代替

双方向リアルタイムが要るなら、

```
Dataverse → Webhook → 自前WebSocketサーバ → ブラウザ
```

の中継パターンが定石。Microsoft 自身もこの設計を想定してる。

### Power Apps への通知

「Dataverse の変更を Power Apps 画面に反映」を実現する 4 つの粒度:

1. **Power Apps プッシュ通知** … Power Automate内で「Power Appsに通知を送信」アクション、スマホ/PC通知
2. **Teams/Outlook 通知経由** … 間接的だが運用上一番多い
3. **`Refresh()` ベース (ポーリング)** … タイマーコントロールで5〜30秒粒度
4. **Modelドリブンアプリの自動更新** … 弱いリアルタイム性

真のリアルタイム双方向には SignalR / Azure Web PubSub が要る。

---

## 9. Excel をフロントにする

`(7)` の 4 分類とは別軸。Office Add-ins の世界。

### 構成 3 パターン

**(A) ローカル完結・軽量**

```
Excel (.xlsm)
└ VBA UserForm + マクロ
└ ローカルファイル or 共有フォルダのデータ
```

- ネット不要、`.xlsm` 一発配布
- 1〜数人で回す内製ツール向け
- 外部連携要らないならこれで十分

**(B) Office Add-ins + 自前バックエンド**

```
Excel
└ Office Add-ins（タスクペイン or Content）
├ フロント配信: VPS (addin.n-sakana.com)
│   └ HTML/JS/CSS、nginx配信
└ API通信
↓
自前バックエンド (FastAPI等、VPS別ポート)
└ DB (PostgreSQL等)
```

- HTML/CSS/JS でリッチなUI
- Office.jsでセル操作
- 認証は自前OAuthか、社内ならEntra ID連携

**(C) Office Add-ins + Dataverse**

```
Excel
└ Office Add-ins
├ フロント配信: VPS or Static Web Apps
└ Dataverse Web API 直接 fetch
↓
Entra ID 認証 (OAuth Authorization Code Flow)
↓
Dataverse
```

- Microsoft 365 / Power Platform の世界に乗る
- Entra ID 認証で社員アカウント直接活用
- Premium ライセンス必須
- 認証は MSAL.js で処理

```javascript
// MSAL.jsでtoken取得
const msalApp = new PublicClientApplication({
    auth: {
        clientId: "<Client ID>",
        authority: "https://login.microsoftonline.com/<Tenant ID>"
    }
});

await msalApp.loginPopup();
const result = await msalApp.acquireTokenSilent({
    scopes: ["https://<org>.crm.dynamics.com/user_impersonation"]
});

// そのtokenでDataverse叩く
const accounts = await fetch(
    "https://<org>.crm.dynamics.com/api/data/v9.2/accounts",
    { headers: { Authorization: `Bearer ${result.accessToken}` } }
).then(r => r.json());

// Office.jsでExcelに書き込む
await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.getRange("A1").values = [[accounts.value[0].name]];
    await context.sync();
});
```

### 選択軸

| 軸 | (A) VBA | (B) Add-ins+自前 | (C) Add-ins+Dataverse |
|---|---|---|---|
| 配布の軽さ | ◎ | △ | △ |
| UI自由度 | △ | ◎ | ◎ |
| 外部連携 | △ | ◎ | ○ |
| 認証 | なし | 自前 | Entra ID標準 |
| ランニングコスト | ゼロ | VPS代だけ | Premiumライセンス |
| 既存資産活用 | VBAそのまま | Webスキル | Power Platform |

### Office Add-ins の種類

- **タスクペイン** … 右側に常駐するサイドパネル、HTML自由
- **Content Add-ins** … シート内に直接埋め込み（最近は非推奨気味）
- **ダイアログ** … モーダル
- **リボン拡張** … 自前のタブやボタン追加
- **カスタム関数** … `=MYFUNC(...)` で自前関数追加

### VBA × Web 通信

| 方向 | VBA |
|---|---|
| Excel→外部（リクエスト送信） | ◎ HTTP普通に叩ける（`MSXML2.XMLHTTP`、`WinHttp.WinHttpRequest`） |
| 外部→Excel（push受信） | ✕ 構造的に無理、ポーリングで代替 |
| WebSocket | ✕ 標準サポートなし、無理と思っていい |

VBA は Excel から能動的に取りに行く・送りに行く動きは強い。受け身（外から push される）は構造的に弱い。

### VBA の位置づけ

Office Add-ins が出ても、VBA UserForm の良さは残る。

- 配布が `.xlsm` 一個
- ネット不要、ローカル完結
- 既存マクロ資産そのまま
- COM の深いところまで触れる

**棲み分け**:
- ローカル業務マクロ → VBA
- 外部システム連携や凝った UI → Add-ins
- 軽い入力フォーム → VBA（Add-ins はオーバーキル）
- 複数人で同じ UI を本格運用 → Add-ins

### Office Add-ins の実体

実体は **Office 内で動く Web アプリ**。

- 実行環境: Office 内に埋め込まれたブラウザ (WebView) — WindowsはEdge WebView2、MacはWKWebView
- 開発: Node.js / npm でビルド (yo office, webpack, vite等)
- 言語: JavaScript / TypeScript + HTML / CSS
- ホスティング: HTML/JS/CSS の配信先が別途必要

「Node.js で開発」だが、最終成果物はブラウザで動くバンドル。サーバーサイドの Node.js が要るのは、Add-ins から叩く自前バックエンド API を書く時くらい（Python/Go でも可）。

### Office Scripts との違い

Office Scripts は TypeScript ベースだが、

- **外部HTTP通信が限定的** … fetch相当が制約あり
- 認証もブラウザ的なOAuthフローを回せない
- 設計思想として「外部連携は Power Automate 経由で」

Dataverse 直接通信は厳しい。実質「Office Scripts → Power Automate → Dataverse」の 3 段構成になる。

### SharePoint Excel / List 経由

技術的にできる。代表的な経路:

| 経路 | 実現性 | 推奨度 |
|---|---|---|
| Excel直 → Power Automate → Dataverse | ◎ | 業務データ集約に良い |
| SharePoint List → Power Automate → Dataverse | ◎ | リスト型業務に良い |
| Excel → SharePoint List → Dataverse | △ | 中継増えて壊れやすい |
| Office Scripts → Dataverse直接 | ✕ | 外部HTTP制約 |
| Office Scripts → Power Automate → Dataverse | ◎ | 現実解 |
| Office Add-ins (JS) → Dataverse直接 | ○ | 自由度高い、実装重い |

Power Platform の世界は「直接通信は基本 Power Automate に集約させる」設計思想。

---

## 10. ホスティングの考え方

Office Add-ins や Web 配信が必要な場面で、どこに置くか。

### 根本的な分類

```
ファイル置き場系（Azure Files / SharePoint / OneDrive）
└ 「ファイルを保管・共有する」のが目的
├ プロトコル: SMB/NFS、または認証付きWeb
├ 想定: 人間がダウンロード/編集する
└ 静的サイト配信機能: ない、または弱い

Web配信系（GitHub Pages / Static Web Apps / nginx / Cloudflare Pages）
└ 「ブラウザに配信する」のが目的
├ プロトコル: HTTPS素通り
├ 想定: ブラウザが直接読む
└ 静的サイト配信: 本職
```

### Azure Files がダメな理由

- プロトコルが **SMB/NFS**（ファイルシステム用）
- ブラウザは SMB 喋れない
- そもそも「Web サーバー」じゃなく「ネットワーク越しのファイル倉庫」

### SharePoint がダメな理由

- ファイルアクセスにログイン前提
- HTML ファイル叩いてもダウンロード扱いされたり Viewer が開いたりで JS が実行されない
- 静的 Web サイトとして配信する設計じゃない
- CORS 設定もユーザー側でいじれない

### Azure Blob Storage は使える

「静的 Web サイトホスティング機能」をオンにすれば、blob に置いたファイルが匿名 HTTPS でブラウザ配信される。Azure Files と名前は似てるが用途が違う。

### Office Add-ins の前提条件

- **HTTPS 必須**（HTTP は Office 側で弾く）
- **CORS 設定が要る場合あり**
- **マニフェスト XML にホスティング URL を記述**

### ホスティング選択肢

**個人開発・自分用**
- 手元の PC で `npm start` → localhost に HTTPS
- Yeoman generator が証明書を自動発行

**無料配布**
- GitHub Pages
- Cloudflare Pages
- Vercel / Netlify
- Azure Static Web Apps (Microsoft純正、Add-ins と相性◎)

**社内本格運用**
- Azure App Service (動的処理も可)
- 自前 IIS (オンプレ)
- VPS + nginx (フルコントロール、コスト追加ゼロで既存運用に乗れる)

### VPS ルートの利点

自前 VPS をすでに持っている場合:
- 既存運用（nginx / Cloudflare / systemd / git pull デプロイ）流用
- コスト追加ゼロ
- フルコントロール
- 既存サービスと隣で動かせる

### CORS の 2 層

Add-ins で CORS は 2 か所に出てくる:

1. **Add-ins 本体 → ホスティングサーバー** … 同一ドメインなら関係なし
2. **Add-ins から外部 API 叩く** … 相手側で許可リスト設定が必要、Dataverse Web API もここに該当

対策:
- 相手側の CORS 設定で Add-ins ドメインを許可
- バックエンドプロキシを挟む

---

## 11. 業務環境制約下での現実解

「自由にサーバを置ける VPS もない、ホスティングできるスペースもない」業務環境では、選択肢が絞られる。

| 状況 | 現実的な選択 |
|---|---|
| Excel をフロント + ホスティング不可 + Dataverse 必要 | (a) Power Apps で諦めてフロント作る |
| 同上、VBA 資産活かしたい | (b) VBA + Power Automate 経由で Dataverse |
| 会社が Azure 契約してくれそう | (c) Static Web Apps の無料枠で Office Add-ins ルート |

**VBA + Power Automate 構成**:
- フロント: VBA UserForm（ローカル完結、`.xlsm` 配布）
- 連携: VBA から Power Automate の HTTP トリガー叩く
- バックエンド: Power Automate フロー → Dataverse
- ホスティング不要、社内 Microsoft 365 アカウントで動く

業務寄りの現実解として無理がない構成。

---

## 用語クイックリファレンス

| 用語 | 意味 |
|---|---|
| HTTP | Webの基本通信プロトコル |
| HTTPS | HTTP + TLS暗号化 |
| REST | HTTPの使い方の流派 |
| WebSocket | 双方向リアルタイム通信プロトコル |
| TLS | 暗号化のお作法、TLS 1.2+が標準 |
| OAuth | パスワード渡さずに権限を貸す仕組み、現行は2.0 |
| Authorization Code Flow | 人間が同意してアプリが代理 |
| Client Credentials Flow | アプリ自身が動く、バックエンド向け |
| Entra ID | 旧 Azure AD、認証の世界 |
| App Registration | Entra ID にアプリを登録したもの |
| Service Principal | App Registration のテナント内実体 |
| Application User | Dataverse 側でアプリを認識する単位 |
| Security Role | Dataverse での権限束 |
| テナント | Microsoft 365 契約全体の単位 |
| 環境 | Power Platform の作業空間、テナント内に複数 |
| Dataverse | Power Platform の共通データ倉庫 |
| Power Apps | 画面を作るサービス、Canvas/Model-driven/Code Apps |
| Power Automate | 自動処理を組むサービス |
| PCF | Power Apps Component Framework、Canvas Apps に HTML部品 |
| Code Apps | フル React で書く Power Apps、2026/2 GA |
| Webhook | レコード変更時に HTTP POST で外に通知 |
| Plug-in | Dataverse 内部に .NET コードを差し込む |
| Office Add-ins | Office 内で動く Web アプリ、HTML/JS/CSS |
| Office Scripts | TypeScript の Excel スクリプト、外部通信は弱い |
