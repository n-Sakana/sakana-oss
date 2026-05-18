# Dataverseフロントエンドのアーキテクチャ候補 一覧

更新日: 2026-05-18  
改訂: Codexレビュー + Opusレビュー統合版  
対象: Power Platform Dataverseをバックエンドにしたフロントエンド構築方法の網羅整理  
作成: hrgn + リュカ (Discord DM 壁打ち)

## 背景と目的

会社PC本番環境では「技術的に可能/不可能」と「本番環境で通るか」は別レイヤー。
セキュリティポリシー・Entra ID権限・DLP・デバイス制御・ライセンス管理・監査運用の制約が強く、
"技術的には確認済みだが、本番環境で通るか要調査" のグレーゾーンが厚い。

この文書は、Dataverseをバックエンドとするフロントエンドを構築する際に
現実的に候補に挙がるアーキテクチャをロングリスト化し、それぞれについて以下を一覧化する。

- ID/認証方式
- 利用者ライセンス
- 作成者/管理者権限
- ALM可否
- 本番で詰まりやすいポイント
- 情シス例外申請の必要性
- 本番通過確率
- 最小検証手順

設計判断の前に「作れるか」ではなく「本番でブロックされないか」「誰の権限で動くか」「誰が運用責任を持つか」を絞り込むためのチェックシートとして使う。

## この文書の前提

- Dataverseは通常のMicrosoft 365付属機能ではなく、Power Apps/Power Automate/Dynamics 365/Power Pages等のライセンスと利用文脈に強く依存する。
- Dataverse for Teamsと本格Dataverseは別物として扱う。
- Microsoft Lists/SharePoint ListsとDataverse for Teamsは別物として扱う。
- DLPで問題になるのは「Dataverseコネクタ単体がブロックされるか」より、「同一アプリ/フロー内で別分類コネクタと混在するか」「複数DLPポリシーの合成でブロックされるか」「Custom Connector/HTTP/外部SaaSが許可されるか」である。
- Web API直叩き、PCF、Custom API、BFF/API中間層、Power Automate/Logic AppsはいずれもDataverseの権限・API制限・容量制限を回避しない。
- PoC許可と本番許可は別審査。本番環境ではALM、監査、所有者、ライセンス、戻し手順まで必要。

## アーキテクチャの大分類

- A. Power Platform純正系
- B. Office / M365製品をフロント化する系
- C. ファイル / マクロ / レガシークライアント系
- D. 独自Webアプリ / デスクトップ / モバイル系
- E. 仲介・ノーコード・データ連携系
- F. Dataverseバックエンド補助系

## ショートリスト: 会社環境前提で最初に検討する候補

ロングリストは網羅目的だが、厳しめの会社PC本番環境では最初から全候補を同列に検討しない。
まずは次の5本を本命候補として比較するのが現実的。

| 優先 | 候補 | 向いている条件 | 選定根拠 | 注意点 |
|---|---|---|---|---|
| 1 | A2 Model-driven Apps | Dataverse中心の業務データCRUD、権限/監査/フォーム重視 | Dataverseのセキュリティモデル、ビュー、フォーム、ALMに最も素直に乗る。本番説明がしやすい | UI自由度は低い。複雑UXはCustom Page/PCF併用を検討 |
| 2 | A1 Canvas Apps | 画面を素早く作りたい、小〜中規模の業務入力 | Power Apps管理下でDLP/共有/ライセンスを説明しやすい。PoCから本番への距離が短い | 複雑ロジック、委任、性能、接続参照、利用者ライセンスに注意 |
| 3 | D5/D8 BFF/API中間層 + Dataverse | 独自UIが必須、認証/監査/レート制御を中央集約したい | SPA直結より情シスに説明しやすい。Application User、OBO、証明書、監査をサーバー側で管理できる | Azure/ホスティング承認、アプリ登録、運用責任、追加コストが必要 |
| 4 | D6 Teams Tab App + Dataverse | Teamsが標準入口、利用者導線をTeams内に閉じたい | 既存M365導線に乗るため現場展開しやすい。Teams SSOと組み合わせやすい | Teamsカスタムアプリ許可、Manifest/配布承認、ホスティング先が壁 |
| 5 | B4 SPFx Webパーツ + Dataverse/BFF | SharePointポータルが業務入口、社内サイトに業務UIを置きたい | SharePointが承認済み入口なら導入しやすい。ポータル/社内Wikiと親和性が高い | SPFx App Catalog、API permission承認、フロントホスティングとCSPに注意 |

補足:

- D1の独自SPA + Dataverse Web API直結は技術的には可能だが、アプリ登録・リダイレクトURI・同意・ホスティング・条件付きアクセス・トークン管理の審査が重い。まずD5/D8の中間層ありと比較する。
- Excel/VBA/Office Scripts系は社内端末制御とファイル運用で詰まりやすく、本命ではなく暫定・周辺業務・読み取り用途として扱う。
- Power Pagesは外部/社外/匿名/認証済みポータルの要件がある場合に強いが、Table Permissionsと容量課金の設計が必須。

## ロングリスト: アーキテクチャ一覧

| ID | アーキテクチャ | ID/認証方式 | 利用者ライセンス | 作成者/管理者権限 | ALM可否 | 本番で詰まりやすいポイント | 情シス例外申請 | 通過確率 | 最小検証手順 |
|---|---|---|---|---|---|---|---|---|---|
| A1 | Canvas Apps + Dataverse | 委任認証。各ユーザーのDataverse権限で実行 | Power Apps Premium/per app等。Dataverse利用権が必要 | Environment Maker、必要に応じてSystem Customizer。共有先にはDataverse security role | 可。Solution、connection reference、environment variable | DLP混在、委任制限、複雑式の性能、共有してもテーブル権限不足、接続参照の本番差し替え | 中 | 高 | 許可環境で1テーブルCRUD、一般ユーザーで実行、DLP違反有無を確認 |
| A2 | Model-driven Apps | 委任認証。Dataverse security roleでアプリ/データ制御 | Power Apps Premium/per app、または適切なDynamics 365文脈 | System Customizer以上が実質必要。アプリ共有はsecurity role経由 | 可。Managed solutionが本番基本 | フォーム/ビュー/サイトマップ権限、BU境界、列レベルセキュリティ、コマンドバー、BPF、モバイル/オフライン | 低〜中 | 高 | 最小テーブル、フォーム、ビュー、アプリモジュールをSolution化し一般ユーザーでCRUD |
| A3 | Power Pages | Power Pages認証。Entra ID/B2C/外部ID/ローカル等 | Power Pages authenticated/anonymous capacity、Dataverse容量 | Power Pages管理、Dataverse table permissions/web role設定 | 可。Power Pages/Dataverse solution。ただしサイト設定の移送注意 | 外部公開審査、Table Permissions漏れ、Web Role、カスタムドメイン/証明書、容量課金、匿名アクセス | 高 | 中 | 非公開/認証必須サイトで1テーブルをTable Permissions付きで公開し一般ユーザーで確認 |
| A4 | Custom Pages + PCF | Power Apps文脈の委任認証。PCFはホストアプリの文脈 | Power Apps Premium/per app等 | System Customizer、PCF import権限、pac/npm/node利用可否 | 可。PCFはSolution前提。本番はmanaged solution | PCF持ち込み審査、npm/pac利用不可、solution checker、CSP/CORS、React/ホスト制約、unmanaged禁止 | 中〜高 | 中 | Developer/SandboxでPCFをmanaged solution importし、一般ユーザーで表示/CRUD確認 |
| A5 | Power Apps Code Apps | Entra ID認証。Power Platform管理下のWebアプリ | Power Apps Premiumが前提になりやすい | 環境でCode Apps有効化、VS Code/git/node/dotnet/pac/npm CLI | 可。ただし機能制限とGit統合可否を都度確認 | 機能成熟度、会社PCの開発ツール制限、preview/GA状況、テナント設定無効化 | 中〜高 | 中 | 管理者にCode Apps有効化可否を確認し、サンプルを許可環境で発行 |
| A6 | Model-driven内 Web resources / JavaScript / Command bar / iframe | Model-drivenアプリ文脈。Xrm.WebApi等の委任認証 | A2と同じ | System Customizer。Web resource/command bar編集権限 | 可。Solution管理 | unsupported DOM操作、保守性、フォームイベント複雑化、iframe先のCSP/認証、権限検証不足 | 中 | 中〜高 | 最小Web resourceからWhoAmI/1件取得を実行し、一般ユーザーで権限差を確認 |
| A7 | Power Apps Teams/SharePoint埋め込み、mobile/wrap配布 | 埋め込み先はM365/Power Apps認証。実データは委任認証 | Power Apps実行ライセンス、モバイル/配布要件に応じてIntune等 | Power Apps作成権限、Teams/SharePoint埋め込み権限、wrapは証明書/配布権限 | 可。ただし配布設定は別管理が残りやすい | Teams/SharePointの埋め込み許可、Power Apps mobile、Intune MAM、オフラインプロファイル、wrap配布 | 中 | 中 | Canvas/Model-drivenをTeams/SharePointに埋め込み、モバイル/社内端末で起動確認 |
| B1 | Excel + Power Query (Dataverseコネクタ) | ExcelのM365サインイン。読み取り中心の委任認証 | Excel/M365 + Dataverse閲覧権限 | Excel利用権限。Power Query/外部接続許可 | 弱い。ファイル配布/クエリ管理中心 | 書き戻し不可に近い、ラベル/IRM/DLP、外部接続ブロック、ローカルキャッシュ、再配布 | 低〜中 | 中 | 機密ラベル付き/なしExcelでDataverse接続、更新、保存先差による挙動を確認 |
| B2 | Office Scripts + Power Automate + Dataverse | Office Scripts単体の外部fetchは弱い。Power Automate経由は接続所有者/実行者文脈 | M365 + Power Automate/Power Apps文脈に応じたPremium | Office Scripts有効化、Power Automate作成権限、Dataverse接続権限 | 一部可。フローはSolution化、スクリプト保管場所に注意 | Power Automate経由では外部fetch不可、OAuth資格情報保管不可、Scripts無効化、Excel Online connector DLP | 中 | 中 | Excel for Webの自動化タブ、Run script、Dataverseコネクタ付きフローを許可環境で確認 |
| B3 | Office Add-ins (Word/Excel) + Dataverse/BFF | Office.js + MSAL/SSOまたはBFF経由 | M365 + Dataverse/Power Apps等の利用権 | 集中配信、AppSource/サイドロード許可、ホスティング承認 | 中。manifest/ソース管理/配信管理。Power Platform Solutionとは別 | サイドロード禁止、集中配信承認、アドイン取得ボタン無効、社内CDN/ホスティング、SSO同意 | 高 | 低〜中 | 組織アドイン配布可否、manifest配信、最小SSO/WhoAmIを確認 |
| B4 | SPFx Webパーツ + Dataverse/BFF | SharePointログイン + MSAL/Graph/Dataverse、またはBFF | M365 + Dataverse利用権 | SharePoint App Catalog、API access承認、SPFx開発/配布権限 | 中。SPFxパッケージはALM可、Dataverse側は別Solution | App Catalog未整備、API permission承認、CSP、社内ポータルの変更管理、ホスティング | 中〜高 | 中 | App Catalog有無、SPFx hello world配布、BFFまたはDataverse WhoAmI確認 |
| B5 | Outlook Add-in / Dynamics 365 App for Outlook | Office.js SSO/MSAL、または標準App for Outlook | M365/Exchange + Dataverse/Dynamics/Power Apps文脈 | Exchange/M365集中配信、アドイン承認、D365 App for Outlook設定 | 中。manifest/集中配信。Dataverse構成はSolution | サイドロード禁止、Exchange Online連携、集中配信承認、メールデータ取り扱い、監査 | 高 | 低〜中 | 組織アプリ配信状況、Outlookでアドイン起動、メール文脈から最小Dataverse参照 |
| C1 | Excel VBA + REST | 実装次第。OAuthが難しく、資格情報管理が課題 | Excel + Dataverse利用権 | マクロ実行許可、署名済みVBA、信頼済み場所 | 弱い。ファイル配布/署名/バージョン管理 | マクロ全面ブロック、Defender ASR、Trust Center、トークン保管、64bit Office差、監査困難 | 高 | 低 | 署名済みxlsmを信頼済み場所で実行し、会社端末でマクロ/HTTP可否を確認 |
| C2 | Excel VBA + Dataverse Web API | デバイスコード/認可コード等。条件付きアクセスとMFAに弱い | Excel + Dataverse利用権 | Entraアプリ登録、管理者同意、マクロ許可 | 弱い | デバイスコードフロー禁止、MFA/サインイン頻度、トークン保管、アプリ登録不可、監査/保守 | 高 | 低 | Entraアプリ登録可否、デバイスコード/認可コードでWhoAmIが通るか確認 |
| C3 | Excel + Office Scripts + Power Automate | Power Automate接続所有者/実行者文脈 | M365 + Power Automate/Power Apps Premium文脈 | Scripts/Power Automate有効化、Dataverse connector利用権 | 中。フローはSolution化可 | Premium割当、Run script制限、Excelファイル保存場所、DLP、同期/タイムアウト | 中 | 中 | Excel Online + Run script + Dataverseコネクタで1件更新、一般ユーザー起動確認 |
| C4 | Microsoft Access + Dataverseリンクテーブル | Office/Access認証、ODBC/OLE DB/Dataverse connector | Access/Runtime + Dataverse利用権 | Access利用許可、ドライバ配布、リンクテーブル設定 | 弱い | レガシー運用、端末配布、ドライバ、ローカルキャッシュ、移行性、マクロ/VBA | 高 | 低 | Access起動/リンクテーブル作成/一般ユーザー更新、ドライバ配布可否を確認 |
| D1 | 独自SPA (React/Vue) + Dataverse Web API直結 | SPAアプリ登録 + MSAL + delegated `user_impersonation` | Dataverse利用権。必要ならPower Apps等 | Entraアプリ登録、redirect URI、管理者同意、ホスティング | 可。ただしPower Platform Solution外。CI/CD別 | アプリ登録不可、ユーザー同意禁止、条件付きアクセス、ホスティング、トークン処理、API制限 | 高 | 低〜中 | SPA登録、localhost redirect、WhoAmI、一般ユーザーでCRUD/429処理を確認 |
| D2 | デスクトップアプリ (.NET/WPF/Electron) + MSAL + Web API | Public client/interactive/device code等 | Dataverse利用権 | アプリ登録、ローカルインストール、コード署名、配布経路 | 中。アプリ配布/更新管理は別 | EXE/MSI禁止、SmartScreen/Defender、署名、MFA/CAE、端末更改、監査 | 高 | 低 | 署名済み最小アプリを会社端末で起動し、MSALログイン/WhoAmI確認 |
| D3 | Power BI (フロント代用) + Dataverse | Power BI connectorの委任認証。更新はPower Apps visual等併用 | Power BI Pro/PPU/Capacity + 必要に応じてPower Apps | Workspace作成/共有権限、Dataverse閲覧権限 | 可。PBIX/Deployment pipeline。Power Apps側は別 | 基本は参照中心、書き戻しは別アプリ、機密ラベル、DirectQuery性能、共有先ライセンス | 低〜中 | 中〜高 | Dataverse接続、共有先閲覧、Power Apps visualで最小更新導線を確認 |
| D4 | Dynamics 365標準フォーム流用 | Dynamics 365/Dataverse標準認証 | Dynamics 365対象アプリのUse Rights | D365環境/フォーム編集/セキュリティロール | 可。Solution/managed solution | 既存D365導入有無、Use Rights制約、標準フォーム改修影響、業務オーナー承認 | 中 | 中 | 既存D365アプリ/テーブル/フォーム編集権限とライセンス文脈を確認 |
| D5 | Azure Static Web Apps + Functions/API Management BFF + Dataverse | フロントはEntra認証。BFFはOBOまたはApplication User/S2S | Dataverse利用権 + Azureコスト。OBOなら利用者側権限も必要 | Azure subscription、App registration、Dataverse application user、Key Vault等 | 可。Azure CI/CD + Dataverse Solutionを分離管理 | Azure利用申請、APIM/Functions権限、証明書/シークレット、Private Endpoint、監査/運用責任 | 高 | 中 | Azureリソース作成可否、BFFからWhoAmI/1件取得、Retry-Afterログ確認 |
| D6 | Teams Tab App + Dataverse/BFF (React/Vue/Blazor WASM含む) | Teams SSO + delegated/OBO、またはBFF | M365 Teams + Dataverse利用権 | Teamsカスタムアプリ許可、アプリ登録、ホスティング | 中〜高。Teams manifest/CI/CD + Dataverse Solution | Teamsアプリポリシー、カスタムアプリ配布、Manifest、SSO同意、ホスティング先 | 高 | 中 | Teams管理センターのアプリ許可、最小Tab配布、Teams内SSO/WhoAmI確認 |
| D7 | Mobile Native (MAUI/React Native/iOS/Android) + Dataverse/BFF | MSAL mobile、またはBFF/Intune管理 | Dataverse利用権 + Intune/MDM/MAM要件 | モバイルアプリ配布、署名証明書、Intune policy | 中。モバイルCI/CDとPower Platform ALMは別 | 社内アプリストア、企業署名、BYOD、Intune App Protection、オフライン/端末紛失 | 高 | 低 | MDM/Intuneポリシー確認、最小アプリ配布、会社端末でログイン/失効確認 |
| D8 | Azure App Service / Container Apps server-side .NET + Dataverse SDK | Server-side confidential client。Application User/S2SまたはOBO | Dataverse利用権 + Azureコスト | Azure hosting、アプリ登録、Dataverse application user、SDK/Key Vault | 可。Azure CI/CD + Solution | ホスティング承認、VNet/Private Endpoint、証明書/secret rotation、SDK差異、監査 | 高 | 中 | App ServiceからDataverse SDKでWhoAmI/CRUD、Application Insightsログ確認 |
| E1 | SharePoint List連携 → Dataverse同期 | SharePoint/Power Automate接続所有者文脈 | M365 + Power Automate/Power Apps文脈 | SharePointサイト/List作成、Power Automate作成、Dataverse権限 | 中。フローはSolution化可、List構成は別 | 二重管理、同期遅延、リストビューしきい値、委任、権限モデル二重化、容量 | 中 | 中 | 1リスト/1テーブル同期、競合/削除/権限差/5000件近辺の挙動確認 |
| E2 | Microsoft Lists / SharePoint Lists + Power Apps | SharePoint/Power Apps委任認証 | M365 + Power Apps標準/必要に応じてPremium | List作成、Power Apps作成/共有 | 中。Power AppsはSolution化可、Listは別管理 | Dataverseではない。複雑リレーション/監査/権限/委任で限界、後でDataverse移行が重い | 低〜中 | 中 | Lists + CanvasでCRUD、件数/権限/委任警告/ラベル挙動を確認 |
| E3 | Dataverse for Teams + Teams内アプリ | Teamsメンバー/所有者モデル + Dataverse for Teams | 一部M365/Office 365に含まれる場合あり。Teams Premium前提ではない | Teams内Power Apps、Teams所有者/メンバー管理 | 限定的。本格Dataverseとは差がある | Teams内に閉じる、容量/機能制限、本格Dataverseとの差、PoC後の本番移行で詰まる | 低〜中 | 中 | Teams内でテーブル/アプリ作成、容量/権限/本格Dataverse移行手順を確認 |
| E4 | Power Automate Desktop + UIフロー | 実行端末/実行ユーザー文脈 | PAD/Power Automate Premium、無人実行は追加 | PADインストール、実行マシン、資格情報管理 | 弱〜中。フロー管理は可能だが端末依存 | RPA監査、端末ロック、無人/有人、VDI/Citrix、資格情報、端末更改で壊れる | 高 | 低〜中 | PADインストール可否、1操作自動化、端末ロック/再起動/監査ログ確認 |
| E5 | Copilot Studio | Copilot Studio認証/チャネル認証。Dataverse接続はコネクタ権限 | Copilot Studioライセンス/メッセージ課金 + Dataverse利用権 | Copilot Studio有効化、生成AI利用許可、Dataverse接続 | 中。Solution化可 | チャットUI限定、生成AI制限、会話ログ、データ越境、DLP、誤回答監査 | 高 | 中 | Copilot Studio有効化、Dataverse参照アクション、会話ログ/生成AIポリシー確認 |
| E6 | Custom Connector / Logic Apps + Dataverse | Custom Connector/OAuth、Logic Apps managed identity/connection | Power Apps/Automate Premium、Logic AppsはAzure課金 | Custom Connector作成、Azure/Logic Apps権限、DLP許可 | 中〜高。OpenAPI/ARM/IaCとSolutionを分離管理 | Custom Connector禁止、HTTP/外部endpoint DLP、Azure統制、コスト予測、接続所有者 | 高 | 中 | Custom Connector保存可否、Logic AppsからDataverse 1件取得、DLP違反確認 |
| E7 | Virtual Tables (Dataverse表面 + 外部DB/API実体) | Dataverse認証 + 外部接続認証 | Dataverse利用権 + 外部DB/APIライセンス | Virtual table/connector/provider設定、外部DB権限 | 可。ただし外部スキーマ/接続設定は別管理 | 性能、リアルタイム依存、書き込み制限、外部認証、検索/集計制約、障害切り分け | 中〜高 | 中 | 外部データ1テーブルをVirtual Table化し、Model-driven/CanvasでCRUD可否確認 |
| E8 | Synapse Link / Fabric / Dataflows / Power BI分析基盤 | サービス間連携。読み取り/分析中心 | Fabric/Power BI/Azure Storage等 + Dataverse容量 | Synapse Link/Fabric/Storage権限、データ所有者承認 | 高。分析基盤ALM/IaC次第 | 参照専用、遅延、データ複製、データ所在地、容量、Purview/ラベル、コスト | 高 | 中 | Synapse Link有効可否、1テーブル複製、Power BI/Fabric側の権限と遅延確認 |
| E9 | Microsoft Loop + Power Automate/Dataverse連携 | M365/Loop認証 + Flow接続所有者文脈 | M365 Loop利用権 + Power Automate/Dataverse利用権 | Loop有効化、Power Automate作成、Connector許可 | 弱〜中。LoopコンポーネントとFlow/Solutionの管理が分かれる | Loop自体のテナント有効化、ガバナンス未成熟、共有範囲、監査、DLP、配布モデル | 中〜高 | 低〜中 | Loop利用可否、共有範囲、Flow連携、Dataverse 1件参照/更新を確認 |
| E10 | OutSystems / Mendix等 iPaaS・ローコード基盤 + Dataverse | 製品ごとのSSO/Connector | 各製品ライセンス + Dataverse利用権 | 外部SaaS契約、Connector承認、製品管理者 | 製品依存 | Connector成熟度、二重ライセンス、外部SaaS DLP、保守スキル集中、ベンダーロックイン | 高 | 低〜中 | 既存契約/社内利用実績を確認し、公式Connectorで1件取得/更新を検証 |
| F1 | Custom API / Plug-in / Webhookによるバックエンド分離 | Dataverse内実行。呼び出し元は各フロントの認証。Application User/OBOも可 | Dataverse/Power Platform利用権 | Plug-in登録、C#開発、solution import、管理者承認 | 可。Solution必須 | Plug-in sandbox 2分、egress制限、デプロイ審査、同期処理のUX影響、トレース肥大化 | 中〜高 | 中 | Custom APIを1つ作り、Canvas/Model-driven/Web APIから呼び出し、権限/タイムアウト確認 |

## ホスティング先の独立評価

独自Webアプリ系では、フロント技術より「どこに置くか」の承認が重いことが多い。
D1/D5/D6/D8/B3/B4/B5は、アーキテクチャとは別にホスティング先を評価する。

| ホスティング先 | 通しやすさ | 監査適合性 | 詰まりやすい点 |
|---|---|---|---|
| Azure Static Web Apps | 中 | 中〜高 | Azure subscription、カスタムドメイン、Built-in Auth、Private network要件 |
| Azure App Service | 中 | 高 | App Service Plan、VNet/Private Endpoint、Key Vault、運用責任 |
| Azure Functions + APIM | 中 | 高 | APIM費用、API設計、証明書/secret、レート制御、監視 |
| 社内オンプレIIS | 中 | 中〜高 | 証明書、パッチ、サーバー運用、Dataverse outbound許可 |
| SharePoint/SPFx | 中〜高 | 中 | App Catalog、API permission、CSP、ページ運用 |
| GitHub Pages / Netlify / Vercel等外部SaaS | 低 | 低〜中 | 外部SaaS、DLP、データ所在地、契約/監査、社内プロキシ |
| ローカルPC配布 | 低 | 低 | 端末制御、署名、更新、監査、退職/端末更改 |

## 本番環境で「通るか」を見る横断チェックリスト

全アーキテクチャに共通する確認軸。

1. **データ分類**: 個人情報、機密区分、営業秘密、金融/人事/評価データ、保存期間、外部共有有無を先に確定する。
2. **ライセンス棚卸し**: M365、Power Apps、Power Automate、Power BI、Dynamics 365、Power Pages、Copilot Studio、Azure課金のどれが必要か。
3. **DLPポリシー**:
   - Dataverse、SharePoint、Excel Online、Office 365 Outlook、HTTP、Custom Connector、Azure、外部SaaSの分類。
   - 同一アプリ/フロー内でBusiness/Non-Business混在にならないか。
   - 複数DLPポリシーの合成でBlockedにならないか。
   - Advanced connector policiesやCustom Connector endpoint制限がないか。
4. **Entra ID権限**:
   - アプリ登録できるか。
   - API Permission付与に管理者同意が要るか。
   - ユーザー同意が禁止されていないか。
   - SPA redirect URI、mobile redirect URI、confidential client secret/certの扱い。
   - Application User/Service PrincipalをDataverse側に作れるか。
5. **条件付きアクセス**:
   - 準拠デバイス必須、MFA、サインイン頻度、継続的アクセス評価、地域制限。
   - デバイスコードフロー禁止、レガシー認証禁止。
   - EntraのWhat Ifで対象ユーザー + Dataverse/Power Platformリソースを評価する。
6. **環境戦略**:
   - Default環境をPoC/本番に使わない。
   - Developer、Trial、Sandbox、Production、Managed Environmentのどれを使うか。
   - 環境security group、リージョン、データ所在地、バックアップ/リストア方針。
7. **環境ロール / Dataverseロール**:
   - Environment Maker、System Customizer、System Administrator、Basic User、カスタムsecurity role。
   - アプリ共有とDataverseデータ権限は別。
   - 一般ユーザーでCRUD検証する。
8. **ALM**:
   - Solution-aware化、managed solution、connection references、environment variables、Power Platform pipelines。
   - Dev/Test/Prod移送、承認ステージ、rollback、solution layer、依存関係。
9. **API Throttling / 性能**:
   - Service protection: 目安として5分約6000 req、累計実行時間約20分、同時実行約52。
   - 429 `Retry-After`、指数バックオフ、クライアント側キュー。
   - `$batch`内部ステップも計上される前提で設計する。
   - Application Userも無制限ではない。
10. **ペイロード / ファイル**:
   - 通常操作、`$batch`、connector、File/Image columnで上限が違う。
   - 標準JSON操作は4MB、`$batch`は16MBを超えない前提で検証する。
   - File/Image columnは専用エンドポイントとchunk upload/downloadを使う。
11. **ストレージ容量**:
   - DB / File / Logの3軸、Dataverse Search index、監査ログ、添付、ファイル、フロー履歴。
   - POA、AsyncOperationBase、PluginTraceLog、AuditBaseの肥大化。
12. **ネットワーク**:
   - `*.dynamics.com`, `*.crm.dynamics.com`, `*.api.crm*.dynamics.com`
   - `login.microsoftonline.com`, `graph.microsoft.com`
   - `*.powerapps.com`, `*.powerautomate.com`, `*.powerplatform.com`, `*.flow.microsoft.com`
   - `cdn.office.net`, `*.officeapps.live.com`
   - 社内プロキシ、TLS inspection、証明書ピンニング、リージョン別URL。
13. **デバイス制御**:
   - マクロ実行、Office Add-in、任意EXE/MSI、Microsoft Store、VS Code、node/npm、dotnet、pac、nuget、ODBCドライバ。
   - Defender/SmartScreen/ASR、コード署名、Intune準拠端末。
14. **データガバナンス**:
   - Purview/Sensitivity Label、IRM、M365 DLP、Power Platform DLP、Information Barriers。
   - Power Automate出力やExcelエクスポートでラベルが落ちないか。
15. **監査ログ / 運用**:
   - M365統一監査ログ、Dataverse Audit、Power Platform Activity Logs、Application Insights、Plug-in Trace Log。
   - 障害時の問い合わせ先、アラート、SLA、所有者退職時の引き継ぎ。
16. **政治・承認**:
   - データオーナー、業務オーナー、Power Platform管理者、Entra管理者、端末管理、ネットワーク、セキュリティ、監査、SAM/ライセンス管理、運用保守のRACI。
   - PoC許可と本番許可の差、CAB/変更管理、承認会議周期。

## Dataverse特有のハマりどころ

| ID | 領域 | 典型的なハマり | 影響する構成 | 対策 / 調査方法 |
|---|---|---|---|---|
| DV-01 | Field Security Profile / 列レベルセキュリティ | 権限のない列がAPIレスポンスから黙って消える。nullではなく列自体が返らない | Canvas、Model-driven、PCF、Web API、BFF | System Adminではなく一般ユーザーで検証。列セキュリティ有効列とProfileメンバーを一覧化 |
| DV-02 | BU境界 + Privilege Depth | ReadはOrganizationだがWriteはUserなど操作ごとに深さが非対称 | 全構成 | Read/Write/Create/Delete/Append/Assign/Shareを操作別に検証。BU階層とModern BU設定確認 |
| DV-03 | API Service Protection | 429、Retry-After、処理時間上限、同時実行上限。Application Userも対象 | Web API、PCF、Power Automate、Logic Apps、BFF | Retry-Afterログ集計、指数バックオフ、API要求ダッシュボード確認 |
| DV-04 | ペイロード上限 | 4MB/16MB前提で設計しないと一括更新や巨大JSONで413/400 | Web API、Power Automate、Logic Apps | 5MB/20MBテストを実施。大きい処理は分割、File columnは専用API |
| DV-05 | File/Image column | 通常列と別扱い。chunk upload/download、thumbnail生成、容量消費 | 添付/画像/ファイルを扱う全構成 | `InitializeFileBlocksUpload`等を使う。ファイルサイズ、保存先、File容量を監視 |
| DV-06 | DB/File/Log容量 | 3軸が独立して枯渇する。Fileだけ、Logだけ枯れる | 全構成 | Power Platform管理センターで容量とテーブル別消費を定期確認 |
| DV-07 | POA肥大化 | レコード共有を多用するとPrincipalObjectAccessが肥大化 | 共有/Access Team/手動Share | Share多用を避け、Access Team/Owner Team設計。POA上位消費確認 |
| DV-08 | AsyncOperationBase肥大化 | 非同期処理、ワークフロー、フロー履歴でDB容量消費 | Power Automate、classic workflow、非同期処理 | Job retention、自動削除、失敗ジョブ監視 |
| DV-09 | PluginTraceLog肥大化 | TraceをONにしたまま本番運用して容量が増える | Plug-in/Custom API | 一時ON、保持期間、定期削除。障害時だけ詳細化 |
| DV-10 | AuditBase肥大化 | 全テーブル/全列監査でLog/DB容量を圧迫 | 全構成 | 必要列だけ監査、retention policy、監査要件と容量見積もり |
| DV-11 | Lookup Polymorphic | Customer/Owner/Regardingの`@odata.bind`書式ミス | Web API、BFF、Power Automate HTTP | entity set名と型を明示。例: `customerid_account@odata.bind` |
| DV-12 | Calculated/Rollup列 | 評価タイミングが即時ではない。Rollupは遅延/再計算待ち | 全UI/API | UI上で即時値として使わない。必要ならサーバー側ロジックで補完 |
| DV-13 | ETag未実装 | 同時更新で後勝ちになり、上書き事故が起きる | Web API、BFF、デスクトップ | `If-Match`、楽観ロック、競合時UXを設計 |
| DV-14 | DateTime 3種別 | User Local / Time-Zone Independent / Date Onlyの混同 | 全UI/API | 列設計時に確定。APIのUTC/ローカル変換をテスト |
| DV-15 | Business Rule実行範囲 | フォームだけ、サーバー側、両方の違いでAPI更新時に抜ける | Model-driven、Canvas、Web API | 重要ルールはPlug-in/Custom APIでサーバー側に寄せる |
| DV-16 | Solution依存ループ | Import失敗、layer競合、削除不可 | Power Platform純正/PCF/Custom API | Solution分割、依存関係確認、managed upgradeで検証 |
| DV-17 | Choice設計 | ローカルChoice乱立で統合困難 | 全構成 | グローバルChoice命名規則、値の固定、翻訳/ラベル管理 |
| DV-18 | Dataverse Search | 未有効化、再インデックス待ち、検索対象列不足 | Model-driven、検索UI | 環境設定、検索対象テーブル/列、インデックス完了を確認 |
| DV-19 | Change Tracking | 既定で全テーブル有効ではない | 同期/差分取得 | 必要テーブルで有効化、差分APIの検証 |
| DV-20 | Currency列 | AmountとBase Amount、為替レート、通貨変更の扱い | 金額を扱う全構成 | 通貨/換算要件を設計で確定し、APIレスポンスを確認 |
| DV-21 | Plug-in/Custom API制限 | Sandbox 2分、外部通信制限、パッケージ制約、同期処理でUX劣化 | F1、A6、D5/D8補助 | 長時間処理は非同期化、Trace/監視、外部通信要件を事前確認 |
| DV-22 | 古いWeb API例 | v8.xや非推奨サンプルを流用して失敗 | Web API全般 | v9.2前提で実装。公式ドキュメントの日付を確認 |

## ライセンス・プライシングの罠

- **Multiplexing違反**: 中間API、Dynamics、RPA、共有アカウントでDataverse操作を中継しても、実際のエンドユーザーに必要なライセンスは免除されない。監査で指摘されると過去分課金請求リスクがある。
- **Dynamics 365 Use Rights**: Dynamics 365に含まれるPower Apps/Dataverse利用権は、該当Dynamicsアプリ文脈に制限されることがある。カスタムテーブル中心の別業務アプリは別途Power Appsライセンスが必要になりやすい。
- **Power Apps per app vs Premium**: per appはアプリ単位・ユーザー単位の割当管理が必要。アプリ数が増えるとPremium per userのほうが安くなる可能性がある。
- **Managed Environment**: Managed EnvironmentでPower Platform資産を実行する利用者/フロー実行者にはPremium系ライセンスや容量が要求される場合がある。
- **Power Automate文脈**: アプリ文脈内のフローは含まれる範囲があるが、スケジュール、バックグラウンド、HTTP、Custom Connector、RPA、無人実行は別ライセンスになりやすい。
- **サービスアカウント**: 接続所有者/実行所有者としてサービスアカウントを置く場合、そのアカウントのライセンス、退職しない所有者、PIM/資格情報管理が必要。
- **Dataverse for Teams**: 一部M365に含まれるが、容量・Teams内利用・ライフサイクル・本格Dataverseとの差がある。PoCで使って本番で再構築になる典型リスク。本格DataverseへのIn-Placeアップグレード不可前提で、移行/再構築コストを見積もる。
- **Default環境**: 全社ユーザーがMakerになりやすく、手動バックアップ不可、本番用途非推奨。PoCに使わない。
- **Dataverse Capacity**: DB/File/Logが別枠。添付、File/Image、監査、検索インデックス、POA、AsyncOperationBase、PluginTraceLogが積み上がる。
- **Power Pages**: authenticated/anonymous capacity、Dataverse容量、外部ユーザー、トラフィック増で費用が跳ねる。バースト想定が必要。
- **Power BI**: Pro/PPU/Capacity、共有先閲覧権限、Power Apps visualを使う場合のPower Appsライセンスが別。
- **Azure中間層**: Functions、App Service、APIM、Key Vault、Application Insights、Private Endpoint、BandwidthのAzureコストと管理責任が別途発生する。
- **Trial/Developer/PoC**: Trial環境やDeveloper Planで動いた構成をそのまま本番環境に昇格できるとは限らない。移送/再作成/managed solution前提で考える。
- **購入リードタイム**: Volume License/CSP/社内購買プロセスで数週間〜数ヶ月かかることがある。PoC完了後に初めて見積もると遅い。
- **AI Builder/Copilot**: AI Builder credit、Copilot Studioメッセージ、Power Platform Copilot系は契約・地域・データ利用条件が変わりやすい。

## 調査の進め方の型

実務的に詰まらない順序。

0. **データ分類の確定**  
   個人情報、機密区分、リテンション、外部共有、利用者範囲、データオーナー、SLAを先に確定する。

1. **業務目的と最小スコープの確定**  
   1テーブル、1画面、1ロール、1更新処理まで絞る。最初から全機能を承認に出さない。

2. **社内事例探索**  
   CoE Toolkit、Managed Environments、社内Wiki、SharePoint、Teams、Power Platform利用者コミュニティ、他部署事例を探す。既に動いている社内事例は情シス説得の最大材料。

3. **情シス/Power Platform管理者に調査用ガードレールを確認**  
   「どの環境ならPoCしてよいか」「Default環境は禁止か」「実データ禁止か」「外部API禁止か」を先に確認する。監査ログに残る前提で動く。

4. **自分のアカウントで触れる範囲を確認**  
   make.powerapps.com、make.powerautomate.com、Power Platform管理センター、Entra管理センター、Microsoft 365管理センター、Teams管理センター、SharePoint管理センターへのアクセス可否を確認。

5. **PoC環境を確保**  
   Developer Plan、Trial、Sandbox、専用開発環境のどれを使うかを決める。Default環境は使わない。

6. **DLP・環境ロール・条件付きアクセス・ライセンスを確認**  
   管理者経由でDLP一覧、環境security group、Managed Environment、容量、Dataverse security role、Entraアプリ登録/同意ワークフローを確認する。

7. **最小プローブを実施**  
   `WhoAmI`、1件取得、作成、更新、削除、Field Security列、BU違い、429/Retry-After、5MB/25MBファイル、一般ユーザー実行を確認する。

8. **本命2〜3個に絞る**  
   アーキテクチャ図、データフロー図、ID方式、ライセンス費、ホスティング先、運用責任、失敗時の戻しを比較表にする。

9. **正式申請**  
   アプリ登録、ホスティング、接続所有者、サービスアカウント、ALM、監査、サポート体制、CAB/変更管理を通す。

10. **小規模パイロット後に本番化判定**  
    実ユーザー、実データに近いダミーデータ、負荷、監査、運用手順、戻し手順を確認してから本番化する。

## 最小検証テンプレート

各候補で少なくとも次を埋める。

| 項目 | 確認内容 |
|---|---|
| 認証 | 誰のIDで実行されるか。ユーザー委任かApplication Userか。OBOか。 |
| 権限 | 一般ユーザーでRead/Create/Update/Delete/Append/AppendToが通るか。 |
| セキュリティ | Field Security、BU境界、チーム権限、共有、ビューを確認したか。 |
| DLP | 同一アプリ/フロー内の全コネクタ分類を確認したか。 |
| CA | 条件付きアクセスWhat Ifを実行したか。 |
| API制限 | 429/Retry-After時の挙動を確認したか。 |
| ペイロード | 5MB JSON、25MB File column、`$batch`分割を試したか。 |
| 容量 | DB/File/Log、Top容量テーブル、監査/Traceの保持を確認したか。 |
| ALM | Dev/Test/Prod移送、connection reference、managed solution、rollbackを確認したか。 |
| 監査 | どのログに誰の操作として残るか確認したか。 |
| 所有者 | 業務オーナー、環境オーナー、接続オーナー、運用担当を決めたか。 |

## 管理者に依頼する確認項目

自分で触れる範囲の確認だけでは本番判断に足りない。
以下はPower Platform管理者、Entra管理者、M365管理者、端末/ネットワーク管理者に分担して確認してもらう。

| 管理領域 | 確認項目 | 依頼先 | 判断に効くポイント |
|---|---|---|---|
| Power Platform環境 | 利用可能なDeveloper/Sandbox/Production環境、Default環境利用禁止ルール | Power Platform管理者 | PoC場所と本番移送経路 |
| Managed Environment | 対象環境がManagedか、共有制限/パイプライン/データポリシーが有効か | Power Platform管理者 | ライセンスとALM要件 |
| DLP | Dataverse、SharePoint、Excel Online、HTTP、Custom Connector、Azure、外部SaaSの分類 | Power Platform管理者 | 同一アプリ/フロー内の混在可否 |
| Advanced connector policies | 許可リスト、endpoint制限、Custom Connector host制限 | Power Platform管理者 | BFF/API/Custom Connectorの可否 |
| 環境security group | 誰が環境に入れるか、JIT追加か、部署別に環境分離されているか | Power Platform管理者 | 利用者展開とBU設計 |
| Dataverse security role | 一般ユーザー用role、作成者role、Application User用role | Dataverse管理者 | CRUDと最小権限 |
| Field Security Profile | 列セキュリティ対象列、Profileメンバー、Read/Update/Create権限 | Dataverse管理者 | 沈黙落としの予防 |
| Business Unit | BU階層、Modern BU、チーム/所有者設計 | Dataverse管理者 | 部署横断参照/更新の可否 |
| 容量 | DB/File/Log使用量、Top消費テーブル、Audit/POA/AsyncOperationBase | Power Platform管理者 | 本番容量見積もり |
| アプリ登録 | ユーザーがアプリ登録可能か、申請ルート、命名/所有者ルール | Entra管理者 | D1/D5/D6/D8の成立可否 |
| 管理者同意 | `user_impersonation`、Graph、SharePoint、Teams SSO等の承認手順 | Entra/M365管理者 | SPA/Teams/SPFx/Add-inの可否 |
| 条件付きアクセス | 準拠デバイス、MFA、サインイン頻度、CAE、地域制限、デバイスコード禁止 | Entra管理者 | VBA/デスクトップ/自動化の停止リスク |
| サービスプリンシパル | Application User作成可否、証明書/secret保管、ローテーション方針 | Entra/Dataverse管理者 | BFF/S2S/Logic Appsの運用 |
| Teams app policy | カスタムアプリ upload、組織配布、Manifest、アプリ許可ポリシー | Teams管理者 | Teams Tab Appの可否 |
| SharePoint App Catalog | テナント/サイトApp Catalog、SPFx配布、API access承認 | SharePoint管理者 | SPFx Webパーツの可否 |
| Office Add-ins | 集中配信、サイドロード、AppSource、組織アドイン | M365/Exchange管理者 | Word/Excel/Outlook Add-inの可否 |
| Office Scripts | Office Scripts有効化、Power Automate統合、外部fetch制限 | M365管理者 | B2/C3の可否 |
| 端末制御 | VBA、EXE/MSI、Store、VS Code、node/npm、dotnet、pac、ODBC | 端末/Intune管理者 | 開発と配布の可否 |
| ネットワーク | Dynamics/Power Platform/Graph/Azure/Office CDNの許可、TLS inspection | ネットワーク管理者 | 接続失敗の切り分け |
| 監査/ログ | 統一監査ログ、Dataverse audit、Power Platform Activity、App Insights | セキュリティ/監査 | 本番障害調査と監査適合 |
| ライセンス | Power Apps/Automate/Pages/BI/Copilot/Azure課金の購入ルートとリードタイム | SAM/購買 | 本番化スケジュール |
| 社内事例 | 既存Power Apps/Dataverse/Teams Tab/SPFx/Power Pages利用例 | CoE/IT Champion | 申請の説得材料 |

## ショートリスト比較観点

本命候補を2〜3個に絞るときは、機能比較ではなく本番通過性で比較する。

| 比較軸 | 見ること |
|---|---|
| ID方式 | 委任認証、OBO、Application User、接続所有者、共有接続のどれか |
| 権限境界 | Dataverse role、BU、Field Security、チーム、共有で要件を満たせるか |
| ホスティング | Power Platform内、Teams/SharePoint、Azure、オンプレ、外部SaaSのどれか |
| デバイス制御 | 会社PCで開発/実行/更新できるか。追加ソフトが必要か |
| DLP適合 | 必要コネクタが同一分類か。HTTP/Custom Connector/外部SaaSが許可されるか |
| ALM | Dev/Test/Prod、managed solution、CI/CD、rollbackがあるか |
| 監査 | 誰の操作としてどこにログが残るか。Application Insights等に流せるか |
| 性能 | API制限、ページング、バッチ、ファイル、検索、同時実行を吸収できるか |
| ライセンス | 利用者数増加時の単価、サービスアカウント、Power Automate/Power BI/Azure費用 |
| 運用体制 | 障害時の一次対応、所有者退職時の引き継ぎ、変更管理のリードタイム |
| 失敗時の戻し | アプリ停止、Solution rollback、データ削除/保全、接続無効化ができるか |

## PoCで作る最小成果物

PoCは「画面の見た目」ではなく、本番で詰まる論点を潰すために作る。

- 1つのDataverseテーブル。
- 1つの一般ユーザー用security role。
- Field Security対象列を1つ含める。
- BUまたはチーム所有の差が出るレコードを2件作る。
- 一般ユーザーでRead/Create/Update/Delete/Append/AppendToを試す。
- `WhoAmI`と1件取得のAPIログを残す。
- 5MB相当の通常payloadと25MB相当のFile columnを試す。
- 429/Retry-Afterを受けた場合のログ出力だけは先に実装する。
- DLP違反が起きる組み合わせを1つ意図的に試し、エラーメッセージを保存する。
- DevからTest相当へmanaged solutionまたはパッケージを移送する。
- 一般ユーザーで実行し、System Administratorでは再現しない問題を記録する。
- 監査ログに誰の操作として残るかを確認する。

## 調査に使う代表ツール

- Power Platform管理センター: 環境、DLP、容量、Managed Environment、分析。
- Power Apps maker portal: Solution、テーブル、列セキュリティ、アプリ共有、Code Apps/PCF。
- Entra ID管理センター: アプリ登録、API permission、管理者同意、条件付きアクセスWhat If。
- Microsoft 365管理センター: ライセンス、Office Scripts、Office Add-ins、監査/コンプライアンス導線。
- Teams管理センター: カスタムアプリ、Teams app permission policy、アプリ配布。
- SharePoint管理センター: App Catalog、API access、サイト/外部共有。
- PowerShell:
  - `Microsoft.PowerApps.Administration.PowerShell`
  - `Microsoft.Xrm.Tooling.PowerShell`
  - `Microsoft.Graph.PowerShell`
  - `Microsoft.PowerApps.Checker.PowerShell`
- CLI:
  - `pac` (Power Platform CLI)
  - `m365` (CLI for Microsoft 365)
- API最小テスト:
  - `GET https://{org}.crm{region}.dynamics.com/api/data/v9.2/WhoAmI`
  - `GET https://{org}.crm{region}.dynamics.com/api/data/v9.2/RetrieveCurrentOrganization(AccessType=@p1)?@p1=Microsoft.Dynamics.CRM.EndpointAccessType'Default'`

## IT部門・大企業文化を踏まえた運用ステップ

- **RACIを先に置く**: データオーナー、業務オーナー、Power Platform管理者、Entra管理者、端末管理、ネットワーク、セキュリティ、監査、ライセンス管理、運用保守。
- **PoC許可と本番許可を分ける**: PoCで許された外部通信、Trial環境、個人所有接続は本番で通らないことが多い。
- **社内事例を探す**: 既に他部署でPower Apps/Dataverse/Teams Tab/SPFxが動いていれば、例外申請ではなく既存パターン踏襲として説明できる。
- **データオーナー承認**: 環境オーナーとデータオーナーは別。Dataverse環境を触れても業務データ利用権があるとは限らない。
- **DPIA/プライバシー影響評価**: 個人情報、評価、人事、顧客情報、EU/UK関連データは法務/プライバシー手続きが必要になりやすい。
- **第三者契約レビュー**: Microsoft DPA、サブプロセッサー、Azure/外部SaaS、iPaaS、AI利用条件を法務が確認済みか。
- **所有者退職対策**: 個人所有Power Apps/Power Automate/接続は孤児化する。Co-owner、サービスアカウント、Solution化、環境管理者を設計する。
- **CAB/変更管理**: 本番Solution import、Azureデプロイ、Teams app配布、SharePoint app配布が変更管理対象か確認する。
- **戻し手順**: Solution rollback、データ削除/保全、接続無効化、アプリ非公開、Azure停止手順を最初から書く。
- **承認会議の周期**: 月次/四半期の審査会がある場合、PoC完了タイミングではなく審査日から逆算する。

## 次の段階で詰めるべきこと

- 会社のセキュリティ特性を整理して、各アーキテクチャの「本番通過確率」を実環境の値で更新する。
- ショートリスト3〜5本について、1ページずつアーキテクチャ図、ID方式、データフロー、ライセンス、運用責任を作る。
- 最小プロトタイプの段取りを切る。1テーブル、1画面、1権限、1更新処理、1ファイル操作で十分。
- 管理者に依頼する情報を整理する。DLP、環境ロール、容量、条件付きアクセス、アプリ登録、ライセンス、既存社内事例。
- 本番化判断用に、コスト試算、リスク、戻し手順、監査ログ、サポート体制をまとめる。

## 参考確認先

最新値・ライセンス・制限は変わるため、実施時点で公式情報を再確認する。

- Dataverse API limits overview: https://learn.microsoft.com/en-us/power-apps/maker/data-platform/api-limits-overview
- Dataverse Service Protection API limits: https://learn.microsoft.com/en-us/power-apps/developer/data-platform/api-limits
- Dataverse connector throttling: https://learn.microsoft.com/en-us/connectors/commondataserviceforapps/
- Dataverse file column data / chunking: https://learn.microsoft.com/en-us/power-apps/developer/data-platform/file-column-data
- Dataverse storage capacity: https://learn.microsoft.com/en-us/power-platform/admin/capacity-storage
- Dataverse security roles: https://learn.microsoft.com/en-us/power-platform/admin/database-security
- Column-level security: https://learn.microsoft.com/en-us/power-platform/admin/field-level-security
- Power Platform DLP connector classification: https://learn.microsoft.com/en-us/power-platform/admin/dlp-connector-classification
- Power Platform environments overview: https://learn.microsoft.com/en-us/power-platform/admin/environments-overview
- Power Platform ALM basics: https://learn.microsoft.com/en-us/power-platform/alm/basics-alm
- Power Apps code apps overview: https://learn.microsoft.com/en-us/power-apps/developer/code-apps/overview
- Dataverse for Teams overview: https://learn.microsoft.com/en-us/power-platform/admin/about-teams-environment
- Comparing Lists, Dataverse for Teams, and Dataverse: https://learn.microsoft.com/en-us/power-apps/teams/compare-data-sources
- Office Scripts external calls: https://learn.microsoft.com/en-us/office/dev/scripts/develop/external-calls
- Power Platform licensing FAQ: https://learn.microsoft.com/en-us/power-platform/admin/powerapps-flow-licensing-faq

## 全体感

この一覧は「使える機能の百科」ではなく、「厳しめの会社PC本番環境で通せるか」を判断するための資料として使う。

最初に見るべきはフロント技術ではなく、データ分類、ID方式、ライセンス、DLP、ALM、ホスティング、運用責任である。
そのうえで、Dataverse中心の業務アプリならModel-driven/Canvas、独自UIが必須ならBFF/API中間層、ユーザー導線重視ならTeams Tab/SPFxを優先して比較する。

Excel/VBA/Office Scripts/RPA/iPaaSは、局所用途では有効だが、本番の中核UIにするには監査・配布・権限・ライセンス・保守の説明コストが高い。
Power Pages、Virtual Tables、Synapse/Fabricは用途が明確な場合に強いが、汎用フロントとして安易に選ぶと別種のガバナンス課題を抱える。

最終判断は「技術的に作れるか」ではなく、「誰の権限で動き、誰が費用を払い、誰が監査され、誰が障害対応し、誰が退職後も維持するか」で決める。
