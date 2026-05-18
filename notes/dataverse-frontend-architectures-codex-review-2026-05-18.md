## Critical（重大な漏れ・誤り）

- **E2 は分類とライセンスが誤解を生みます。** 「Microsoft Lists (Dataverse for Teams)」は一つの方式ではなく、Microsoft Lists/SharePoint Lists と Dataverse for Teams は別物です。Dataverse for Teams は一部の Microsoft 365 / Office 365 サブスクリプションに含まれる Teams 向け Dataverse 環境であり、Teams Premium 前提とは書かないほうがよいです。候補は「Microsoft Lists / SharePoint Lists + Power Apps」と「Dataverse for Teams + Teams アプリ」に分離してください。
- **DLP の記述は補正が必要です。** Dataverse コネクタは Power Platform の中核コネクタで、従来型 DLP ではブロック不可に分類される扱いがあります。ただし Business / Non-Business の混在、複数 DLP ポリシーの合成、Advanced connector policies、カスタムコネクタや HTTP との併用で保存・実行が止まることはあります。A1 の「Dataverse コネクタが Business / Non-Business で分かれてるとブロック」は、「Dataverse 単体」ではなく「同一アプリ/フロー内で別分類コネクタと混在するとブロック」に修正すべきです。
- **B2 の Office Scripts + Dataverse Web API は、技術的成立条件がかなり厳しいです。** Office Scripts の外部 `fetch` は Excel アプリ上では可能ですが、OAuth2 の対話的サインインや安全な資格情報保管の仕組みがなく、Power Automate 経由実行では外部 API 呼び出しが失敗する制約があります。Dataverse Web API へ直接 OAuth 付きで安定接続する候補としては弱く、「Office Scripts + Power Automate / HTTP with Entra ID / Dataverse コネクタ」に寄せるほうが現実的です。
- **本番判定に必須の中間 API / BFF パターンが抜けています。** 会社 PC・Entra・DLP・デバイス制御が厳しい前提では、React/Vue から Dataverse Web API へ直接接続する D1 だけではなく、「フロントエンド + Azure Functions/App Service/API Management/Logic Apps などの中間層 + Dataverse S2S/OBO」が重要候補です。アプリ登録、管理者同意、サービスプリンシパル、Dataverse アプリケーションユーザー、証明書/シークレット管理、監査、レート制御を中間層に集約できます。

## Major（補足必須）

- **D1 の CORS は主ブロッカーとして書きすぎです。** Dataverse Web API は SPA からの CORS 利用をサポートしています。実際に詰まるのは、SPA アプリ登録の可否、リダイレクト URI 登録、Dataverse `user_impersonation` 権限、ユーザー同意禁止、条件付きアクセス、社内プロキシ、ホスティング許可です。
- **Power BI は「完全に書き戻し不可」と断定しないほうがよいです。** Power BI レポート自体は分析・参照中心ですが、Power Apps visual や Power Automate button を組み込めば更新導線は作れます。ただし Power BI + Power Apps/Automate の複合ライセンス、DLP、埋め込み先、利用者権限が増えるので、本命候補にするなら別行に分けて評価してください。
- **Power Pages は「外部公開系」だけでなく、内部/認証済みポータルとしての利用も評価対象です。** 本番で詰まるのは外部公開の印象だけではなく、サイト可視性、カスタムドメイン/証明書、WAF/プロキシ、Web roles、Table permissions、匿名アクセス禁止、Portal Web API の範囲、容量課金、外部ユーザー認証方式です。
- **A4 の PCF 調査方法は `pac pcf push` だけでは足りません。** 本番観点では、managed solution として import できるか、solution checker を通せるか、PCF の npm/node/pac CLI が会社 PC で使えるか、社内プロキシ越しに npm registry へ出られるか、ALM パイプラインに乗るか、unmanaged 持ち込み禁止に抵触しないかを確認してください。
- **A5 Code Apps はステータスと制約を都度確認する列が必要です。** 2026-05-18 時点では公式ドキュメントが整備されていますが、機能制限・前提ツール・環境設定・Premium ライセンス要件が変わりやすい領域です。VS Code、git、Node.js、dotnet、Power Apps CLI/npm CLI の利用可否が会社 PC で大きな壁になります。
- **Dataverse セキュリティモデルの行が不足しています。** アプリ共有、環境アクセス、Dataverse security role、テーブル権限、列レベルセキュリティ、所有者/チーム/BU/共有によるレコードアクセスは別レイヤーです。「アプリが開けるがデータが見えない」「API は 403」「System Administrator だけ再現しない」が頻発します。
- **ビュー権限とビューはセキュリティ境界ではない点を明記してください。** System view / personal view / app module に含める view / FetchXML の結果は、最終的に Dataverse のレコード権限で絞られます。ビューで非表示にしただけでは API や別ビューから見える可能性があります。
- **API 制限は横断チェックリストに昇格すべきです。** Dataverse には entitlement と service protection の両方があり、外部 API では 429 と `Retry-After` 前提の設計が必要です。既定値の目安として service protection は 5 分スライディングウィンドウ、要求数・実行時間・同時要求数で制限されます。Dataverse コネクタにも 300 秒あたり 6000 calls/connection の目安があります。
- **4MB 上限は「Dataverse 全 API の一律 payload 上限」として書かないでください。** ファイル/画像列では Web API の chunk upload/download があり、推奨 chunk size として 4,194,304 bytes が返るケースがあります。一方で単一リクエストで 128MB 未満のファイルアップロードもあります。対象がコネクタ、Web API、File column、添付、Power Automate アクションのどれかで制限が違います。
- **ストレージ容量は database/file/log と Dataverse Search を分けて見る必要があります。** 添付、Note、File/Image columns、監査ログ、プラグイントレース、フロー履歴、検索インデックスが容量に効きます。容量不足は単なる課金問題ではなく、運用・管理操作の停止リスクになります。
- **Default environment を試作場所にするリスクを明記してください。** Default environment は全社ユーザーが maker になりやすく、手動バックアップ不可、本番用途非推奨です。調査は許可された Developer/Sandbox/専用環境で行うべきです。

## Minor（補足あるとよい）

- A1/A2 は「利用者ライセンス」と「作成者/管理者権限」を列で分けると読みやすくなります。Environment Maker、System Customizer、System Administrator、アプリ共有先ユーザーの必要ロールは別物です。
- A2 Model-driven apps は sitemap/theme だけでなく、フォーム/ビュー/コマンドバー/Business Process Flow/モバイル対応/オフライン可否/アプリモジュール共有も詰まりどころです。
- A3 Power Pages は Table permissions の設定漏れが情報漏えいの主要リスクです。匿名アクセスを許すテーブル、Web API 有効化、列単位の露出を確認項目に入れてください。
- B1 Excel + Power Query は「更新は可能だが書き戻しは基本別経路」「Excel ファイルの持ち出し/ラベル/DLP/IRM/ローカルキャッシュ」が本番リスクです。
- B3 Office Add-ins は集中展開だけでなく、Office アドインの manifest、AppSource 禁止、アドイン取得ボタン無効、共有ランタイム、社内 CDN/ホスティング承認も確認対象です。
- C1/C2 VBA 系は、OAuth トークン保管、MSAL 相当の実装難度、64bit Office、Trust Center、署名済みマクロ、信頼済み場所、Defender ASR ルールを補足してください。原則として本命候補ではなく、暫定/個人運用リスク枠に寄せるのが妥当です。
- E1 SharePoint List 連携は「5000件制限」ではなく「リストビューしきい値・インデックス・委任・同期整合性・二重権限モデル」を書くほうが正確です。Microsoft Lists/SharePoint は大量データの物理上限より、実運用の検索・ビュー・権限継承で詰まります。
- E3 PAD は有人/無人、端末ロック、実行アカウント、Windows セッション、監査、資格情報管理、Citrix/VDI、端末更改時の壊れやすさを分けて評価してください。
- E4 Copilot Studio はチャット UI 限定だけでなく、生成 AI 利用可否、データ越境、会話ログ、DLP、Dataverse への接続、Copilot Studio ライセンス/メッセージ課金を確認してください。

## 追加すべきアーキテクチャ候補（あれば）

- **D5 独自 Web アプリ + BFF/API 中間層 + Dataverse**: Azure Functions/App Service/API Management/Logic Apps/Container Apps など。S2S または OBO で Dataverse に接続。企業本番では直接 SPA より説明しやすいことが多いです。
- **D6 Teams タブアプリ + Dataverse**: Teams Toolkit/React/SPFx などで Teams 内 UI として提供。Teams 管理センターのカスタムアプリ許可、アプリカタログ、アプリ権限ポリシー、Entra 同意が詰まりどころです。
- **B4 SharePoint Framework (SPFx) Web パーツ + Dataverse/API 中間層**: SharePoint が標準入口の企業では通しやすい候補。テナントアプリカタログ、管理者承認、Graph/API 権限、ホスティング、CSP が論点です。
- **A6 Model-driven app 内の Web resources / JavaScript / Command bar / iframe 埋め込み**: 独自 UI を最小限にし、Dataverse の認証・フォーム・セキュリティ文脈に乗せる方式。unsupported DOM 操作や保守性がリスクです。
- **A7 Power Apps の Teams/SharePoint 埋め込み、Power Apps mobile/offline/wrap**: フロント配布方式として別評価に値します。Intune/MAM、モバイル利用許可、オフラインプロファイル、アプリラップ配布の承認が論点です。
- **B5 Microsoft Access + Dataverse リンクテーブル**: Access が許可されている組織では候補。ただし accdb 配布、マクロ/VBA、ローカルキャッシュ、権限、長期保守の観点で本番には重いリスクがあります。
- **E5 Custom Connector + Canvas Apps / Power Automate**: Dataverse 直接ではなく社内 API をカスタムコネクタ化して Power Apps から使う方式。DLP、カスタムコネクタ作成権限、OpenAPI 管理、Premium ライセンスが論点です。
- **E6 読み取り専用/分析特化: Synapse Link/Fabric/Dataflows/Power BI + Power Apps visual**: 業務更新ではなく参照・分析中心なら候補。リアルタイム性、容量、Fabric/Power BI ライセンス、データ複製ガバナンスを評価します。

## 観点別の追加コメント

### 1. 網羅性
現状は Power Platform 純正、Office、VBA、独自アプリ、連携系が一通りあります。ただし「企業本番で通しやすい入口」である Teams、SharePoint/SPFx、中間 API、Power Apps 埋め込み/モバイル配布が抜けています。直接 Dataverse に行く方式と、社内承認済みの入口に寄せる方式を分けると判断しやすくなります。

### 2. 本番で詰まりやすいポイント
本番で最も詰まるのは、技術可否よりも「誰の ID でデータアクセスするか」です。ユーザー委任、サービスプリンシパル、アプリケーションユーザー、偽装、共有接続、Power Automate 実行所有者を候補ごとに明記してください。特にサービスプリンシパルは Dataverse 側にアプリケーションユーザーを作り、専用 security role を割り当てる必要があります。

### 3. 調査方法
「触ってみる」だけでなく、以下の証跡が取れる形にしてください。

- 自分・一般利用者・管理者相当の 3 アカウントで、作成/実行/共有/CRUD の結果を表にする。
- Power Platform 管理者に、DLP ポリシー、環境 security group、Managed Environment、環境種別、容量、テナント設定、許可済みコネクタを確認依頼する。
- Entra 管理者に、アプリ登録可否、ユーザー同意可否、管理者同意ワークフロー、条件付きアクセス、デバイスコードフロー禁止、SPA redirect URI 登録ルールを確認依頼する。
- 最小プロトタイプはダミーデータ、許可された開発環境、1 テーブル、1 画面、1 権限ロールで実施する。
- API 系は `WhoAmI`、1件取得、作成、更新、削除、429/`Retry-After` 処理、5MB/25MB ファイル操作をチェック項目にする。

### 4. 横断チェックリスト
追加推奨の観点は以下です。

- 環境戦略: Default/Developer/Sandbox/Production、Managed Environment、環境 security group、リージョン/データ所在地。
- ALM: solution-aware 化、managed solution、connection references、environment variables、Power Platform pipelines、承認ステージ、rollback、solution layer 管理。
- 認証/認可: delegated vs application permission、OBO、アプリケーションユーザー、証明書/シークレット保管とローテーション、PIM/JIT、break-glass。
- 運用: 監視、アラート、監査ログ、Application Insights、Dataverse analytics、CoE Starter Kit、問い合わせ窓口、障害時復旧、バックアップ/リストア。
- データ保護: Purview labels、DLP、IRM、保持/削除、eDiscovery、外部共有、ダウンロード/Excel エクスポート制御。
- パフォーマンス: API request entitlement、service protection、バッチ上限、ページング、delegation、検索、インデックス、添付/ファイル容量。
- 端末/ネットワーク: Intune 準拠端末、Defender/ASR、SmartScreen、コード署名、社内プロキシ/TLS inspection、npm/nuget/pac/VS Code の利用可否。

### 5. 進め方の型
現在の「IT部門に聞く前に物理的に何が動くか把握」は、厳しめ企業では shadow IT や監査アラートに見える可能性があります。順序は次のほうが安全です。

1. 業務目的、データ分類、利用者数、外部共有有無、更新頻度、SLA、データオーナーを先に確定。
2. 情シス/Power Platform 管理者に「調査用に許可された環境と禁止事項」だけ先に確認。
3. 許可された環境で最小プローブを実施し、権限・DLP・ライセンス・端末制約を記録。
4. 候補を 2〜3 個に絞り、ライセンス費・例外申請・運用負荷・監査性で比較。
5. 本命候補だけ正式申請し、アプリ登録/ホスティング/ALM/サポート体制を通す。
6. 小規模パイロット後に本番化判定を行う。

### 6. Dataverse 特有のハマりどころ
追加すべき項目です。

- Security role は累積評価。複数ロール/チームロールで意図せず権限が広がる。
- BU 境界、owner team、access team、共有、階層セキュリティ、modernized business units の組み合わせで見え方が変わる。
- Column-level security は API/アプリ全体に効くが、System Administrator では隠れないため検証アカウントを分ける必要がある。
- App sharing と Dataverse data access は別。アプリを共有してもテーブル権限がなければ実行時に失敗する。
- Business rules、plug-ins、classic workflow、Power Automate、duplicate detection、required fields、status reason が API 経由でも発火/制約になる。
- `$select` を使わない全列取得、深い `$expand`、大量 FetchXML、N+1 API はすぐ性能問題になる。
- `$batch` は最大 1000 個の個別要求を含められるが、entitlement/service protection を回避する手段ではない。
- File/Image/Attachment/Note は容量・chunking・DLP・ダウンロード制御を別途評価する。

### 7. 情シス文化・政治的ステップ
不足しているのは、承認者の整理です。データオーナー、業務オーナー、Power Platform 管理者、Entra 管理者、端末管理/Intune、ネットワーク、セキュリティ、監査、SAM/ライセンス管理、運用保守の RACI を早めに置くべきです。また「例外申請が必要な候補」と「既存ガードレール内で実施できる候補」を分けると、情シス側の心理的負担が下がります。

### 8. プライシング・ライセンスの罠
見落としやすい罠は以下です。

- Dataverse は基本的に Premium 扱い。Microsoft 365 付属の Power Apps 権利だけで通常 Dataverse 本番アプリを動かせるとは限らない。
- Dataverse for Teams は M365 に含まれる場合があるが、容量・Teams 内利用・ライフサイクル・本格 Dataverse との差がある。
- Dynamics 365 ライセンスの Power Apps 利用権は「該当 Dynamics 365 アプリ文脈」に制限されることがある。
- Power Apps per app は「アプリ単位で利用者に割り当てる」発想で、環境/アプリ/ユーザーの割当管理が必要。
- Managed Environment では利用者・フロー実行者に Premium 系ライセンス/容量が要求される場面がある。
- Power Automate はアプリ文脈内なら含まれる範囲がある一方、スケジュール/バックグラウンド/HTTP/カスタムコネクタ/RPA/無人実行は別ライセンスになりやすい。
- Power Pages は authenticated/anonymous capacity、Dataverse 容量、外部ユーザー、トラフィック増で費用が跳ねる。
- Power BI は Pro/PPU/Capacity、Power Apps visual は Power Apps ライセンス、共有先の閲覧権限が別に必要。
- Dataverse database/file/log/request capacity は小さな追加容量の積み上げで不足しがち。監査ログ、検索インデックス、添付、フロー履歴も効く。
- Azure Functions/App Service/API Management/Key Vault/Application Insights/Private Endpoint など中間層の Azure コストも比較表に入れる。

### 照合した主な公式資料

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

一覧の骨格はかなり良く、候補の広がりも実務的です。ただし、現状は「フロントエンドの形」中心で、「企業本番で承認されるための ID、環境、DLP、ALM、ライセンス、運用責任」の列がまだ弱いです。

特に大企業では、技術的に最短の D1 直接 SPA よりも、A1/A2 の純正 Power Apps、A6 の model-driven 内拡張、D5 の中間 API、D6/B4 の Teams/SharePoint 入口のほうが通しやすいことがあります。最終判断は「作れるか」ではなく「誰が所有し、誰の権限で動き、誰が監査・障害対応・費用負担するか」で決まります。

次の改訂では、各行に「ID方式」「利用者ライセンス」「作成者/管理者権限」「ALM可否」「情シス例外申請」「本番通過確率」「最小検証手順」を追加すると、設計判断にそのまま使えるチェックシートになります。
