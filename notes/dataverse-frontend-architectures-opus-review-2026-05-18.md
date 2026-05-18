# Dataverseフロントエンドアーキテクチャ一覧 レビュー（Opus 4.7 max）

レビュー日: 2026-05-18
対象: `dataverse-frontend-architectures-2026-05-18.md`
レビュアー: Claude Opus 4.7 (1M context)

## Critical（重大な漏れ・誤り）

### C1. 列レベルセキュリティ (Field Security Profile / CLS) の「沈黙落とし」問題

A1〜A5, B2, C2, D1, D2 など Web API を直接叩く構成すべてに関わる。

- Field Security Profile が適用された列は、権限のないユーザーに対し **APIレスポンスから列ごと黙って削除されて返る**（エラーは出ない）。フロント側で「nullになる」と勘違いして調査時間を浪費しやすい。
- 機微列（個人ID、給与、評価等）が含まれるテーブルは、開発者アカウント (System Administrator) では見えるが、本番ユーザーには見えないケースが多い。
- これは Canvas Apps / Model-driven Apps でもフォーム上で同様の挙動。
- **調査方法**:
  - `make.powerapps.com` → テーブル → 列 → 「列セキュリティを有効にする」のフラグ立った列を一覧。
  - 各 Field Security Profile のメンバーと Read/Update/Create 権限を `Get-CrmFieldSecurityProfile` 系の PowerShell で出力。
  - **実エンドユーザーのアカウントで** Web API 叩いて、何が返ってこないか確認（System Administratorで開発検証するのは原則NG）。

### C2. Business Unit 境界 + 特権の深さ (Privilege Depth) の非対称

セキュリティロールの「特権の深さ」(None / User / BU / Parent:Child BU / Organization) は、**同一テーブル内でも操作（Read/Write/Create/Delete/Append/AppendTo/Share/Assign）ごとに独立**して設定される。

- 結果として **Read は Organization だが Write は User** のような非対称が頻発する。
- 開発者は通常 System Administrator (全部 Organization) で開発 → 本番ユーザーで「読めるが更新できない」「自分のレコードしか見えない」が発覚。
- 親子BU階層を組んでいる組織では、データが「自分のBU + 子BU」までしか見えない設定が一般的。フロントから他部署データを参照しようとして空が返るパターン。
- **Modern BU / Matrix Data Access** (2023〜) を有効化した環境では、従来のBUオーナーシップとは別軸でアクセス可能になっており、設計時の見立てが崩れる。
- **調査方法**:
  - 「Security Role Browser」(コミュニティアプリ) や `Test-CrmSecurityRole` で **想定エンドユーザー** を使った権限シミュレーション。
  - 環境の Business Unit 階層を `make.powerapps.com` → 環境 → 設定 → ユーザー + 権限 → ビジネスユニット で確認。
  - 「impersonation」を有効にしたサービスプリンシパルで他ユーザー名義のクエリを試す（`MSCRMCallerID` ヘッダー）。

### C3. API Service Protection Limits（スロットリング）

Web API を直接叩く全構成（A4 PCF, B2, C2, D1, D2, E系）に共通する隠れ地雷。明示記載が必要。

- 5分あたり **約 6,000 リクエスト/ユーザー/サーバー**、**累計実行時間 約 20分/ユーザー/サーバー**、**同時実行 52**。
- `429 Too Many Requests` + `Retry-After` ヘッダー（秒）で返る → クライアントは**指数バックオフ実装必須**。
- **サーバー単位** で計算されるが、クライアントはサーバー数を知れない → 観測上限は変動して見える。
- Application User（サービスプリンシパル）も同じ制限。「サーバーサイドだから無制限」は誤解。
- `$batch` を使う場合、内部ステップ数もカウント対象。プラグイン経由のクエリは追加で「Database execution time」制限あり。
- **調査方法**: Power Platform 管理センター → 環境 → 分析 → 「Dataverse」/「API要求」ダッシュボード。429のRetry-Afterログを集計する仕組みをフロントに組み込む。

### C4. ペイロード上限とファイル列・画像列の特殊扱い

- **単一 Web API リクエスト: 4MB**、**$batch: 16MB**。超過で 413/400。
- 「複数レコード一括 PATCH」を素朴に書くと、レコード件数次第で超過する → ChangeSet 分割設計が必要。
- ファイル列 (File column)・画像列 (Image column) は別エンドポイント（`InitializeFileBlocksUpload` → `UploadBlock` → `CommitFileBlocksUpload`）で 4MB チャンク分割が必須。Power Automate / Logic Apps の標準コネクタは裏でこれをやってくれるが、**自前 Web API では実装必要**。
- 画像列は複数サイズ（full/thumbnail）が自動生成 → ストレージ消費を見落としやすい。
- **調査方法**: Postman で 5MB の PATCH/POST を実際に投げて 413 を確認。File column は 1GB 上限を念のため確認。

### C5. ストレージ容量（DB / File / Log の3軸）と ライセンス連動

E1 (SharePoint List) の同期遅延より深刻な「Dataverse ネイティブ容量枯渇」が表に書かれていない。

- テナント基本容量: **DB 10GB + File 20GB + Log 2GB**。
- ライセンスごとの追加容量:
  - Power Apps Premium (per user): DB +250MB, File +400MB, Log +2GB
  - Dynamics 365 (種別による): DB +5GB, File +5GB, Log +2GB
- **3つの容量が独立**。File 多用で File だけ枯渇するパターンが本番障害の典型。
- 容量を食い潰す「裏方テーブル」: **PrincipalObjectAccess (POA)**, **AsyncOperationBase**, **PluginTraceLog**, **AuditBase**。
- Trial / Developer Plan 環境は容量が小さく、PoC で File アップロード試験すると一瞬で埋まる。
- **調査方法**: Power Platform 管理センター → リソース → 容量 → 環境別 DB/File/Log 使用量 + テーブル別ストレージレポート。「Top 5 容量消費テーブル」を必ず確認。

### C6. Microsoft Teams Tab App カテゴリの完全欠落

E2 (Dataverse for Teams) と別物で、独立カテゴリとして必要。

- **Teams Tab App (Teams Toolkit / TeamsFx)** で React/Vue/Blazor を Teams 内に埋め込み、**通常の本格 Dataverse** に SSO + Entra ID で接続する構成。
- ホスティングは Azure Static Web Apps か独自インフラ。
- 既に M365 + Teams を展開している会社で「ユーザー導線」が最短になるため、本番採用候補として有力。
- **本番で詰まりやすい**: カスタムアプリのアップロード許可（テナント設定 / 個別 / グローバル管理者承認）、Teams 管理センターのアプリ ポリシー、Manifest 署名。
- 「**B4. Teams Tab App + Dataverse Web API**」として追加すべき。

### C7. ホスティング先が D系で過小評価

D1 で「ホスティング先確保」と1行で済まされているが、ここが本番化最大の壁になる。アーキテクチャ表に「**ホスティング先**」と「**監査適合性**」のサブ列を入れた方がよい。

- Azure Static Web Apps（無料枠あり、Entra認証内蔵だが、カスタムドメイン・本番モードは法人ガバナンス審査）
- Azure App Service（有料、Vネット連携、Private Endpoint対応 → 情シスから見て統制しやすい）
- 社内オンプレ IIS（既存資産あれば最速、HTTPS/証明書整備が課題）
- SharePoint ページに HTML 配置（SPFx 経由でなくただ置く運用は監査困難で嫌がられる）
- GitHub Pages / Netlify / Vercel（外部SaaS → DLPでほぼ確実にブロック）

ホスティング選択は **「フロント技術」より組織内承認難易度が高い** ことが多い。

---

## Major（補足必須）

### M1. Connection References と ALM の前提整理

A1, A2, A4, A5, B2, C3, E1〜E4 の Solution-aware 系すべてに関わる。

- Power Automate フロー、Canvas App、PCF は Solution に入れて Dev → Test → Prod へ移送するのが本番運用必須。
- 移送先で **接続参照 (Connection Reference) の手動マッピング** が必要 → 「サービスアカウントは誰?」「そのライセンスは?」「Dataverse 接続はどのユーザーで?」を本番化前に固める必要がある。
- **Premium コネクタを使うフローは、接続オーナー（実行ユーザー）に Premium ライセンスが必要**。サービスアカウントへの Premium 付与は IT で渋られがち。
- 表に「ALM 経路」「サービスアカウント設計」列を足すと、本番化検討が前進する。
- **調査方法**: 「サービスアカウント用ライセンス確保」を IT に早期投げ。承認に数週間〜数ヶ月の組織が多い。

### M2. Conditional Access の影響範囲を具体化

横断チェックリスト3に1行あるが、本番障害の常連なので具体化が必要。

- **準拠デバイス必須ポリシー**: BYOD / 個人PC / 開発機からのアクセス全弾き。
- **MFA 要求**: VBA / デスクトップアプリのデバイスコードフローで「対話的補完が必要」となり、自動化に向かない。
- **サインイン頻度 (Sign-in Frequency)**: トークン有効期限が短くなる → 自動化ジョブが定期的に止まる。
- **継続的アクセス評価 (CAE)**: ユーザー状態変化が即時反映、長時間バッチで途中で切れる。
- **国・地域ベース**: 海外出張中アクセス不可。
- **アプリ単位制御**: Dataverse / Power Platform Apps グループが条件付きアクセスでどう扱われているか。
- **調査方法**: Entra ID 管理センター → 条件付きアクセス → 「What If」ツールで、自分のアカウント + ターゲットリソース (`https://*.crm.dynamics.com`) で評価。

### M3. Outlook Add-in カテゴリの追加

B3 の Office Add-ins と一括りにしているが、Outlook 連携は独自考慮点があるので分けるべき。

- **App for Outlook / Dynamics 365 App for Outlook** は Dataverse 標準機能、CRM レコードとメールを紐付ける UI を提供。
- カスタム Outlook Add-in (Office.js) + Dataverse Web API は MSAL.js + SSO で実装可能。
- 集中配信 (Centralized Deployment) は M365 グローバル管理者 or Exchange 管理者の承認必要。
- 個人配信は会社設定でほぼ全面禁止が多い。
- 「**B5. Outlook Add-in + Dataverse Web API**」として独立記載すべき。

### M4. 「進め方の型」の順序修正提案

現案: 自分で試す → 管理センター確認 → IT 経由で DLP/ロール → 候補絞り → 申請

修正案（ステップ0と2を追加）:

1. **データ分類の確定**（個人情報・機密区分・リテンション要件） — これが先に決まらないと後で「DLP で外部送信ブロック」と差し戻される。
2. **既存社内事例の探索**（CoE Toolkit があれば必ず存在チェック、他部署 Power Platform 利用者へのヒアリング、SharePoint で社内 Wiki / ナレッジ検索） — **未公開社内資産が答えのことがある**。
3. 自分のアカウントで触れる範囲を全部試す（**ただし** 監査ログに残ることを意識、IT に「PoC で試しています」と一声入れておくと摩擦回避）。
4. 管理センター確認 + **ライセンス棚卸し**（並行）。
5. **PoC 環境確保**: Trial (30日) / Developer Plan (無料・永続・個人開発専用) / Sandbox (組織所有・要申請) の使い分け。**Default 環境を PoC に使うのは避ける**（後述）。
6. DLP・環境ロール・条件付きアクセスを情シス経由で確認。
7. 本命2〜3個に絞り、**アーキテクチャ図 + データフロー図 + ライセンスコスト試算** を作成して正式申請。
8. PoC 実行 → セキュリティ審査 → 本番。

特に「**社内事例探索**」は組織政治的にも強力。「既に他部署で動いている」は情シスを動かす最大の説得材料。

### M5. Customer Lockbox / Information Barrier / Sensitivity Label の伝播

横断チェックリスト7を具体化。

- **Customer Lockbox**: Microsoft サポートが顧客データにアクセスする時の承認制御。有効化テナントではトラブル時の対応が遅れる → 本番運用 SLA に影響。
- **Information Barriers**: M365 全体で部署間の通信/データ共有を遮断する設定。Dataverse 直接対象外だが、Power Pages の B2B シナリオや Teams 連携で衝突。
- **Sensitivity Label**: Power BI Dataverse コネクタ経由のラベル伝播 (Inherit) が機能するか、Power Automate フロー出力に伝播するかは別物。フロー出力からラベルが落ちる "downgrade" 事故あり。
- **Power Platform DLP と M365 DLP は別エンジン**。両方確認必要。

### M6. Power Pages (A3) の落とし穴

- **認証プロバイダ**: Entra ID, B2C, Local Auth, Google, Microsoft Personal など複数。**B2C は別途 Azure サブスクリプション + B2C テナント必須**。
- **Web Role / Table Permissions**: Dataverse セキュリティロールと**別軸**のアクセスモデル。Dataverse 経験者ほど混乱する。
- **キャパシティ課金**: 認証ユーザー (Authenticated User) と匿名 (Anonymous) で別ライセンス、別課金。アクセス急増で**バースト課金**になりやすい → Capacity Add-on の手当を予算化。
- **コードのモード**: Liquid (ローコード) と Pro Code (Bootstrap + JS) の混在で複雑化、開発体制設計が要件。
- **カスタムドメイン**: 独自ドメイン + SSL は別途設定、組織の DNS 管理権限が必要。

### M7. PCF (A4) の具体制限

- **CRM iframe 内**で実行 → CSP / CORS の制約あり。任意の外部 API コール不可、Dataverse Web API 経由を強制される。
- React バージョンは**ホスト環境固定**（Virtual PCF はバージョン分離可能だが Preview）。
- フィールドバインド型 vs データセットバインド型で開発モデルが大きく異なる。
- `pac pcf push` には開発環境のシステムカスタマイザー権限 + Solution 書き込み権限が必要。**個人 Developer Plan で開発し、最終 Import を管理者経由でやる**運用が現実的。

### M8. ライセンスの罠（補足）

- **Power Apps per app は 1ライセンス = 1アプリ に紐付く**: バンドル不可。アプリ追加時に追加購入。
- **Power Apps Premium (per user)** は1ユーザーが無制限アプリを使える。アプリ数増えるなら早期切り替えがコスト的に有利。
- **Pay-As-You-Go (PAYG)**: Azure サブスクリプション紐付け、$10/月/アクティブユーザー程度。ライセンス購入手続きを回避できるが、**Azureコストセンターを持っていない部署には逆に取り入りにくい**。
- **Dynamics 365 ライセンス**にバンドルされた Power Apps は **Dynamicsエンティティに関連した用途** 限定。**カスタムテーブル中心のシナリオでは別途 Power Apps ライセンス必要**（"Use Rights" の罠）。
- **Power Automate Premium**: 接続オーナーに必要。「呼び出すユーザー」にも必要なケースあり（プロセス型フロー）。
- **Dataverse for Teams** は **1テーブル100万行・環境2GB上限**、本格 Dataverse へのIn-Placeアップグレード**不可**（再構築必要）。PoC で使って本番でハマる典型パターン。
- **Multiplexing ルール違反**: Dynamics 経由でデータを「中継」して別 UI から操作させても、エンドユーザーには別途ライセンスが必要。**監査で指摘されると過去分課金請求**されることがある。

### M9. Default 環境を PoC に使うリスク

- Default 環境はテナント全ユーザーが Maker ロールを持つ → **他人のリソースを誤って削除/編集できる事故**が起きやすい。
- IT 部門から見ても監査困難。
- **PoC は必ず別環境（Developer Plan 無料の個人環境、または Trial、Sandbox）で**。
- Developer Plan は無料・永続、開発専用、容量上限あるが PoC 用途には十分。

---

## Minor（補足あるとよい）

### m1. ネットワーク許可ドメイン補足

横断チェックリスト5に追記推奨:

- `*.api.crm{region}.dynamics.com` (Web API エンドポイント、リージョン別)
- `*.svc.dynamics.com`
- `*.powerapps.com`, `*.powerautomate.com`, `*.powerplatform.com`
- `*.flow.microsoft.com`
- `cdn.office.net`, `*.officeapps.live.com` (Office Add-ins/Scripts)

リージョン: Japan East (`jpn`), US (`usa`), EU (`eur`) など、テナントの Home region 次第。Proxy / SSL Inspection で「Microsoft 証明書ピンニング失敗」する組織もある（ネット監視装置が MITM 中のとき）。

### m2. PowerShell モジュールでの調査効率化

調査をスクリプト化すべきもの:

- `Microsoft.PowerApps.Administration.PowerShell` — 環境/DLP/ロール一覧
- `Microsoft.Xrm.Tooling.PowerShell` — Dataverse 接続・FetchXML 実行
- `Microsoft.Graph.PowerShell` — Entra ID 側の権限・条件付きアクセス確認
- `Microsoft.PowerApps.Checker.PowerShell` — Solution チェッカー

CLI:
- `pac` (Power Platform CLI) — 環境作成、Solution Import、PCF 開発
- `m365` (CLI for Microsoft 365) — テナント全体の設定確認

### m3. CoE Toolkit / Managed Environments の認識

中〜大企業の Power Platform 成熟組織には CoE Toolkit / Managed Environments が入っていることが多い。

- **Managed Environments**: 環境ガバナンス機能が有効化（ソリューションチェッカー強制、メーカー制限、データポリシー継承）。
- 入っていれば、ガバナンスは整っているので使いやすい。
- 入っていない環境では、自分で安全策を講じる必要。
- **調査方法**: 管理センター → 環境一覧 → 「Managed」フラグ確認。

### m4. テナント間ゲスト（B2B）の罠

複数テナントのゲストになっている場合、`Switch directory` でアカウント切り替えが必要。

- Dataverse Web API は**テナントごとに別エンドポイント** → トークン取得時に正しい tenant_id を指定。
- Multi-tenant のクライアントアプリは Entra ID 登録時に "Multi-tenant" 有効化必要、管理者同意必要。

### m5. Power BI (D3) の書き戻し可能オプション補足

- Power BI **+ Power Apps Visual** で書き戻し可能（Canvas Apps を埋め込むので結局 A1 系の制限を引き継ぐ）。
- **Power BI Datamart** は別物、内部的に SQL DB で Dataverse とは別。
- **Power BI Direct Query for Dataverse** はパフォーマンス課題あり（複雑クエリで遅い）。
- 機密ラベルの伝播は Power BI Pro/PPU + 設定が必要。

### m6. 監査ログの取得粒度

別物として、本番障害調査時に「どこに何が残るか」を事前マッピング:

- M365 統一監査ログ（Compliance Center）
- Dataverse Activity Logs (Application Insights 連携可能、別途設定)
- Power Platform Activity Logs (`make.powerapps.com` → 環境 → 監査ログ、デフォルト30日)
- Plug-in Trace Log（プラグイン障害追跡用、別保持期間）

### m7. Excel 系（B1, B2, C1, C3）共通の運用課題

- **共有 Excel ファイル**として配布する場合、ファイル更新時の再配布の煩雑さ。
- 個人 OneDrive 保存 vs SharePoint Doclib 保存で**外部接続 (Power Query) の挙動が変わる**。
- **機密ラベル付き Excel** からの外部接続を禁止する設定が増えている（M365 E5）。
- マクロブロックポリシー（ファイルがインターネット由来とマークされていると VBA 実行不可）。

### m8. Plug-ins / Custom APIs / Webhooks (バックエンド拡張)

これらはバックエンド拡張だが、フロントエンドアーキテクチャの「裏側のロジック分離」として重要:

- フロント側で複雑な処理を書く代わりに、Dataverse **Custom API**（サーバー側 C#）で抽象化 → フロントは1回の API コール、で複雑な業務処理を実装。
- **Plug-in sandbox**: 2分タイムアウト、ネットワーク egress 制限（許可ドメインのみ）、サードパーティパッケージ取り込みに制約、System.Net 系の一部 API 不可。
- **Webhook**: Dataverse から外部 HTTP エンドポイントへのコールバック、外部システム連携に使える。

### m9. テスト用エンドポイント

「Web API 接続できるか」の最小テスト:

```
GET https://{org}.crm{region}.dynamics.com/api/data/v9.2/WhoAmI
GET https://{org}.crm{region}.dynamics.com/api/data/v9.2/RetrieveCurrentOrganization(AccessType=@p1)?@p1=Microsoft.Dynamics.CRM.EndpointAccessType'Default'
```

これが Postman / curl で通れば、認証・ネットワーク・権限の基本が揃っていることが確認できる。各アーキ調査の最初のステップに組み込む価値あり。

---

## 追加すべきアーキテクチャ候補

### B4. Microsoft Teams Tab App + Dataverse Web API

- React/Vue/Blazor SPA を Teams Toolkit で開発、Teams 内タブとして提供。
- Teams SSO で Entra ID トークン取得 → Dataverse Web API 直叩き。
- **ライセンス**: M365 (Teams) + Dataverse 接続権限 + Entra ID アプリ登録。
- **詰まりやすい**: Teams カスタムアプリのアップロード許可、組織内配布承認、Manifest 署名。
- **調査**: Teams 管理センター → Teams apps → Permission policies、アプリ アップロード設定。

### B5. Outlook Add-in + Dataverse Web API

- Office.js + MSAL.js でメール画面に Dataverse 連携 UI。
- **ライセンス**: M365 + アプリ登録 + 集中配信。
- **詰まりやすい**: 集中配信承認、サイドロード禁止、Exchange Online 連携必須。
- **調査**: Exchange 管理センター → 組織アプリ → 集中配信状況。

### B6. SharePoint Framework (SPFx) Web Part + Dataverse Web API

- SharePoint ページに React Web Part を埋め込み、Dataverse Web API。
- **ライセンス**: M365 (SharePoint) + アプリ登録 + テナント アプリカタログ。
- **詰まりやすい**: SPFx App Catalog のテナント有効化、開発者署名証明書、テナント拡張機能の許可設定。
- **調査**: SharePoint 管理センター → 詳細 → API アクセス、アプリカタログ。

### D5. Azure Static Web Apps + Azure Functions (BFF) + Dataverse Web API

- フロント SPA を Static Web Apps、サーバーサイドプロキシを Functions に、認証は Built-in Auth (Entra ID)。
- **メリット**: CORS 問題回避、トークン管理サーバー側に閉じれる、コスト低、Azure 内で完結。
- **ライセンス**: Azure サブスクリプション (Static Web Apps Free/Standard, Functions Consumption)。
- **詰まりやすい**: Azure サブスクリプション確保、リソースグループ作成権限、Static Web Apps 認証連携設定、Functions の Application User 登録。
- **調査**: 自身の Azure サブスクリプション権限 (Reader / Contributor)、IT 部門の Azure 統制ポリシー。

### D6. Blazor WebAssembly + MSAL + Dataverse Web API

- C# フルスタック開発者向け、SPA だが JavaScript ベースではない。
- Dataverse SDK もブラウザでは使えないので結局 OData/Web API。
- **詰まりやすい**: 初期ロードサイズ、AOT コンパイル設定、ホスティング先。
- **調査**: 開発機での `dotnet new blazorwasm` 動作、ホスティング先選定。

### D7. Mobile Native App (MAUI / React Native / iOS/Android) + Dataverse Web API

- Intune マネージドモバイル、または Power Apps mobile 経由。
- **ライセンス**: Intune または社内 MDM 統合、アプリ署名証明書。
- **詰まりやすい**: 社内アプリストア配信、企業署名証明書、Intune App Protection Policy、業務利用に必要な BYOD ポリシー。
- **調査**: モバイルデバイス管理ポリシー、MAM 設定。

### D8. Azure App Service (Server-side .NET) + Dataverse SDK (Server-side)

- フロントは別途、バックエンドが App Service 上の ASP.NET Core で **Dataverse Service Client SDK** 使用。
- **メリット**: フル SDK 機能（Web API にない高度な操作 - Execute Multiple, ImportSolution など - も可能）、サーバー側でトークンキャッシュ、Application Insights 連携、Private Endpoint。
- **詰まりやすい**: Azure サブスクリプション、App Service Plan、Vネット連携、Application User 登録、SDK バージョン差異。

### E5. Azure Logic Apps + Dataverse Connector

- Power Automate の Azure 版、エンタープライズ向け。
- **メリット**: Azure DevOps 連携、IP フィルタ、Vネット統合、ARM Template での IaC。
- **デメリット**: Azure サブスクリプション必要、コスト体系が Actions 実行ベース（予測しづらい）。
- **詰まりやすい**: Azure ガバナンスが Power Platform より厳しい組織あり、Connector の機能差。

### E6. Virtual Tables / Virtual Entities

- Dataverse から見ると通常テーブルだが、実体は外部 SQL / REST API。
- **用途**: Dataverse をフロント・ロジックに使いつつ、データソースは外部維持（既存 DB 活用、Migration 不要）。
- **詰まりやすい**: パフォーマンス（リアルタイム参照）、認証 (Virtual Connector App Setting)、書き込み対応の制約、CRUD 制限。

### E7. Synapse Link for Dataverse → Azure Data Lake → カスタム BI / 分析基盤

- 読み取り専用の分析パイプライン。
- **用途**: 大量データ分析、長期保存、Power BI 直結より柔軟。
- **詰まりやすい**: Synapse Workspace / Azure Storage Account の権限、CDM 形式の理解、リアルタイム性は望めない（最低15分遅延）。

### E8. Microsoft Loop コンポーネント + Power Automate + Dataverse

- 2024〜の新機能、Outlook/Teams/Word/Loop アプリ内で生きるコンポーネント。
- **詰まりやすい**: Loop 自体がテナントで有効化されているか、ガバナンス未成熟、配信モデルが未確立。

### E9. Outsystems / Mendix / Salesforce Lightning + Dataverse Connector

- 既に契約・利用がある企業向けの選択肢として記載すべき。
- **詰まりやすい**: 各製品の Dataverse コネクタの成熟度、保守スキル集中、二重ライセンス費用。

### C4. Microsoft Access + Dataverse Connector (ODBC/OLE DB)

- レガシーだが現役の業種あり（特に古参の業務システム持ち込み案件）。
- **詰まりやすい**: ODBC ドライバ配布、Access Runtime ライセンス、移行性ゼロ。

### F1. Plug-in / Custom API / Webhook によるバックエンド分離 (アーキテクチャ補助)

- フロントエンド「単独」ではないが、上記いずれと組み合わせる「ロジック層」として明示すべき。
- Dataverse 側に Custom API を作り、フロントからはそれを1回呼ぶ → トランザクション保証・複雑業務ロジックを集約。
- **詰まりやすい**: Plug-in Sandbox の 2分タイムアウト、egress 制限、デプロイのリードタイム。

---

## 観点別の追加コメント

### 観点6. Dataverse 特有のハマりどころ（追加表）

| 領域 | 典型的ハマり | 対策 |
|---|---|---|
| POA テーブル肥大化 | レコード共有 (Share) を多用すると POA が数百万行に → DB 容量 + パフォーマンス劣化 | 共有を Access Team に切り替え、定期的な POA クリーンアップ |
| AsyncOperationBase | 非同期プラグイン/フロー大量実行で肥大化 | Job Retention Policy 設定（自動削除） |
| Plugin Trace Log | デフォルト OFF だが ON にすると爆増 | 一時的 ON のみ、定期削除 |
| Audit Log | 全項目監査 ON はストレージ枯渇の元 | 必要列のみ ON、Retention Policy |
| Solution の依存ループ | A → B → A の循環で Import 失敗 | Layer 分割、Solution Checker |
| Choice (旧 Option Set) のグローバル/ローカル混在 | グローバルにすべきものをローカルで作ると後で統合困難 | 命名規則、グローバル前提 |
| Lookup の Polymorphic (Customer/Owner/Regarding) | Web API で `@odata.bind` 形式間違いやすい | Type 明記 `"customer_account@odata.bind": "/accounts(...)"` |
| Calculated / Rollup Column の評価タイミング | 即時 vs 遅延、null 処理 | 仕様確認後 UI 設計 |
| Business Rule の実行範囲 | フォーム / サーバー / 両方 | Webhook やプラグインで補完 |
| Dataverse Search 未有効化 | search / quickFind の挙動が貧弱 | 環境設定 → Dataverse Search 有効化、再インデックス待ち |
| 楽観ロック (ETag) 未実装 | 同時更新で後勝ち発生 | `If-Match` で ETag 検証、競合時のUX設計 |
| ChangeTracking | 全テーブルでデフォルト有効ではない | 必要テーブルで有効化（API による変更通知） |
| Currency 列 | Base Currency への自動換算挙動 | 換算レート、Currency Field の2重持ち（Amount + Amount_Base） |
| DateTime 列の TimeZone | User Local / Time-Zone Independent / Date Only の3種 | 設計時に確定、Web API のフォーマット注意 |
| 古い Web API バージョン残存 | v8.x のコードがネット上に多い | 必ず v9.2 を使う |

### 観点7. 大企業 IT 政治・運用ステップ補足

- **「PoC 許可」と「本番許可」は別審査**: PoC が通っても本番で同じシステムが使える保証はない。両方の審査要件を事前確認。
- **データオーナーシップの明確化**: 触るデータの管掌部門が情シスや別部署にある場合、別途その部署の承認が必要。Dataverse では「環境オーナー」と「データオーナー」が分かれることが多い。
- **DPIA (Data Protection Impact Assessment) / プライバシー影響評価**: 個人情報を扱う場合、EU/UK 関係なら必須。日本企業でも個情法改正で社内手続きとして整備されつつある。
- **第三者契約レビュー**: Microsoft の利用規約・データ処理契約 (DPA)・サブプロセッサーリストを法務がレビュー済みか。M365 導入時に通っていても、Power Platform 部分が別レビューになる組織あり。
- **退職時引き継ぎ要件**: 個人開発した Power Apps / Power Automate は所有者 (Owner) が退職すると孤児化 → サービスアカウントまたは Co-owner で保護。Solution 化も孤児化対策の一手。
- **CAB (Change Advisory Board) / 変更管理プロセス**: 本番 Solution Import が変更管理対象になる組織あり、リードタイム数日〜数週間。
- **IT Champion / Power Platform CoE Lead の特定**: 組織にいれば必ず接触。いない場合は自分が(暗に)その役を担うことになる覚悟。
- **「失敗した時の戻し」プラン**: PoC 失敗時にデータをどう保全・廃棄するか、最初に決めておく。
- **承認会議の周期**: 月次 / 四半期で開かれる審査会のスケジュールに合わせて成果物を準備。これが本番化スピードの最大ボトルネック。

### 観点8. プライシング・ライセンスの罠（追加）

- **Trial → 本番への昇格は環境ごと別**: Trial 環境を本番に昇格できない、必ず本番環境を別途作成 → 移送。
- **ライセンス購入のリードタイム**: Volume License 経由だと購入承認に数週間。CSP (Cloud Solution Provider) 経由ならクレカ即購入だが、企業によっては禁止。
- **Power Pages 課金の予測困難性**: 認証ユーザー (Authenticated User) は「アクティブユーザー数 × 月」課金、突発的アクセス増で予算超過。Capacity Add-on の手当を予算化。
- **AI Builder クレジット**: Power Automate の AI アクション、Copilot で個別消費。プレフィックスとしてのテナント割当 + 追加購入。
- **Dataverse Capacity Add-on**: DB/File/Log 別に購入可能だが、購入決裁が頻繁発生しやすい。月次 / 四半期で使用量レビューを業務に組み込む。
- **Multiplexing ルール違反**: Dynamics 経由でデータを中継して別 UI から操作させても、エンドユーザーには別途ライセンス必要。監査で指摘されると過去分課金請求される。
- **Co-pilot 関連**: 2025〜26 で Copilot Studio や Power Platform Copilot のライセンスモデルが頻繁変更。**契約交渉時の有効期限**を確認、自動更新で爆増事故あり。
- **「Use Rights」の罠**: Dynamics 365 / Power Apps の "Use Rights" 文言を細部まで読まないと、想定利用が違反になっていることがある。法務にレビューを通すか、Microsoft アカウントマネージャーに書面回答を依頼。

---

## 全体感

このリストは「使える機能の網羅」より「**組織で実装可能かのチェック**」に重心を置いており、その方向性は正しい。特に DLP / Entra ID 権限 / 環境ロールを横断軸に据えた構成は、本番化の壁を整理する道具として優秀。

その上で、本番化フェーズまで見据えるなら以下3つの強化を勧める。

1. **Dataverse ネイティブの非機能制限を明示化**: API throttling, ペイロード 4MB, ストレージ 3軸, Field Security, BU 境界 — これらは「アーキ選定」より下のレイヤーだが、選んだ後に必ず引っかかる。冒頭の「横断チェックリスト」に **Dataverse 固有のサブセクション**を追加するのがよい。

2. **ホスティング先という第2軸の独立化**: 特に D系（独自 Web アプリ）は「フロント技術」より「**どこにホストするか**」の方が組織内承認難易度が高い。表の「アーキテクチャ」列とは別にホスティング選択を分離して並べると、議論が回りやすい。Azure Static Web Apps / App Service / オンプレ IIS / SharePoint の4択を表現できると IT 部門との会話が楽になる。

3. **「進め方の型」のステップ0 (データ分類) と ステップ2 (社内事例探索)**: この2つが入ると、後工程の手戻りが大幅に減る。特に「社内事例探索」は **「既に他部署で動いている」という事実が情シスを動かす最大の説得材料** であり、政治的に最強。

最後に — このリストは個人の調査用としては非常に良くできている。本番化フェーズで IT 部門に提示する場合は、「アーキテクチャ表」を**ロングリスト**、絞り込んだ2〜3個を**ショートリスト**として別資料に分け、ショートリスト側に「セキュリティ要件マッピング」「運用負荷」「コスト試算」「失敗時の戻し計画」を加える二段構えにするのが、組織提案として通りやすい型である。
