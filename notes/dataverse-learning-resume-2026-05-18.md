# Dataverse学習レジュメ 2026-05-18

対象読者: Dataverseを実務で触っているが、公式用語で説明しようとすると詰まりやすい人。特に、VBAやExcel業務ツールの経験があり、WebシステムやPower Platformの語彙を体系的に整理したい人。

このレジュメは、次の2資料を土台にして再構成した。

- `notes/dataverse-stack-overview.md`
- `notes/dataverse-frontend-architectures-2026-05-18.md`

用語はできるだけ現在の公式表記に寄せる。ただし実務では古い名前が残っているため、必要に応じて `Table(旧Entity)`、`Column(旧Field/Attribute)`、`Row(旧Record)` のように併記する。ライセンス、制限値、製品提供状態は変わりやすいので、実案件では必ず実施時点のMicrosoft公式情報を確認する。

---

# 第1部 Webシステムの一般論

## 1. はじめに

### ふんわり入口

Dataverseを学ぶときに最初に困るのは、Dataverseそのものの難しさだけではない。むしろ、周辺の言葉が一気に押し寄せることが難しい。

たとえば、次のような言葉が同じ会話に出てくる。

- Table
- Column
- Row
- Web API
- OData
- OAuth
- Entra ID
- Security Role
- Environment
- Solution
- DLP
- Canvas Apps
- Model-driven Apps
- PCF
- Application User

一つひとつは別の分野の言葉だが、Dataverseでは全部がつながっている。Excelで言えば、ブック、シート、セル、マクロ、ユーザーフォーム、共有フォルダ、パスワード保護、VBA参照設定が一つの業務ファイルにまとまっているようなものだ。

このレジュメでは、いきなりDataverseの設定画面に飛び込まない。まずWebシステム一般の言葉をそろえる。次にMicrosoftの世界で同じ概念がどう呼ばれるかを見る。そして最後に、実装パターンとハマりどころを整理する。

### 正確に言うと

このレジュメで扱うDataverseは、Microsoft Power Platformの共通データ基盤としてのMicrosoft Dataverseである。Dataverseは単なるデータベースではなく、データ保存、権限、API、監査、業務ロジック、アプリ連携、ALMをまとめて扱うプラットフォームである。

本書の読み方は次の通り。

1. 第1部でWebシステムの土台を作る。
2. 第2部でその土台をMicrosoft Power PlatformとDataverseの公式用語に対応させる。
3. 第3部で現実の会社環境で使えるアーキテクチャ候補を比較する。
4. 第4部でDataverse特有の罠を事典として引けるようにする。

「サルでもわかる」系の入口にするが、用語を雑に丸めない。難しい言葉は避けず、先に日常の例えを置いてから、公式寄りの定義に進む。

### VBA的にいうと

VBA経験者は、次の対応で考えると入りやすい。

| VBA/Excelの感覚 | Web/Dataverse側の言葉 |
|---|---|
| `.xlsm` ファイル | アプリ、Solution、環境内の資産 |
| Worksheet | Table、画面、ビューの土台 |
| `Cells(1, 1)` | RowとColumnの交点 |
| UserForm | フロントエンド |
| Buttonの `_Click` | JavaScriptやPower Fxのイベント |
| ADO接続 | Web API呼び出し、Connector |
| Recordset | APIから返ったJSON配列 |
| 標準モジュール | JavaScript module、Power Fx式、Plug-in |
| 参照設定 | npm package、Connector、API permission |

ただし完全に同じではない。Excelは1つのファイルの中で画面、データ、コードが密結合になりやすい。Dataverseでは、データはDataverse、画面はPower AppsやWebアプリ、認証はEntra ID、権限はSecurity Role、移送はSolutionという具合に、責任が分かれる。

### 図

```mermaid
graph TD
    A[第1部 Web一般論] --> B[第2部 MicrosoftとDataverse]
    B --> C[第3部 実装パターン]
    C --> D[第4部 ハマりどころ事典]
    A --> A1[DB/API/認証/フロント/セキュリティ]
    B --> B1[Table/Column/Web API/Entra ID/Solution]
    C --> C1[34アーキテクチャと本番ブロッカー]
    D --> D1[DV-01からDV-22]
```

### 混同しやすい近接概念

「Dataverseを学ぶ」と「Power Appsを学ぶ」は近いが同じではない。Power Appsは画面を作るサービスで、Dataverseはデータと権限とAPIの土台である。

「Dataverseを学ぶ」と「SQL Serverを学ぶ」も近いが同じではない。Dataverseの背後にはデータベース的な仕組みがあるが、利用者が直接SQLで自由にテーブルを作ってJOINする製品ではない。DataverseではTable、Relationship、Security Role、Solution、Web APIなどの抽象化を通して扱う。

「ローコードだから簡単」と「本番運用が簡単」も同じではない。画面を作るだけなら簡単でも、DLP、環境、ライセンス、監査、ALM、所有者退職対策まで含めると、普通の業務システムと同じだけ設計が必要になる。

### ここを押さえれば次に進める

Dataverseは「データベースの代わり」だけではない。DB、API、認証連携、権限、監査、イベント、運用のまとまりとして見ると理解しやすい。VBAで一体化していたものが、WebとMicrosoft 365の世界では部品に分かれている。この部品の名前を順番にそろえるのが、このレジュメの目的である。

---

## 2. システムってなんだろう

### ふんわり入口

システムとは、ざっくり言えば「人が毎回手でやると面倒な仕事を、決まった手順で回す仕組み」である。

身近な例で考える。飲食店の注文を想像する。

1. お客さんがメニューを見る。
2. 店員さんに注文を伝える。
3. 厨房に注文が届く。
4. 料理人が作る。
5. 会計で金額が計算される。
6. 売上が記録される。

これを全部紙でやっても業務は回る。ただし、忙しくなると聞き間違い、書き間違い、二重入力、集計漏れが起きる。そこで、注文端末、厨房モニター、レジ、売上データベースをつないで「注文システム」にする。

会社の業務システムも同じである。申請、承認、顧客管理、在庫管理、案件管理、日報、問い合わせ管理などを、画面、データ、権限、通知、集計でつないだものがシステムである。

### 正確に言うと

システムは、目的を達成するために複数の要素が相互に関係して動く仕組みである。ITシステムでは、主に次の要素がある。

| 要素 | 役割 |
|---|---|
| 利用者 | 画面を操作する人 |
| フロントエンド | 利用者が触る画面 |
| バックエンド | 業務処理やAPIを担う裏側 |
| データベース | データを保存する場所 |
| 認証 | 利用者が誰かを確認する仕組み |
| 認可 | その人が何をしてよいかを決める仕組み |
| 通知 | メール、Teams、プッシュ通知など |
| ログ/監査 | 誰が何をしたかを残す仕組み |
| 運用 | 障害対応、変更管理、バックアップ、権限棚卸し |

Dataverseを中心に見ると、Dataverseはこのうちデータベース、認可、API、監査、イベントフックの多くをまとめて引き受ける。

### VBA的にいうと

VBAの業務ツールでは、1つの`.xlsm`に多くの役割が入る。

| `.xlsm`内の要素 | システム一般論 |
|---|---|
| シート | データ保存、簡易DB |
| UserForm | フロントエンド |
| 標準モジュール | 業務ロジック |
| ボタンイベント | ユーザー操作の入口 |
| 非表示シート | 設定値やマスタ |
| 保護、パスワード | 簡易的な権限制御 |
| 共有フォルダ配布 | 簡易的なデプロイ |

Excelではこの一体感が強みである。一方、本番で複数人が同時に使うと、ファイル破損、バージョン違い、権限管理、監査、同時編集、配布が苦しくなる。Dataverseを含む業務システムは、ここを分離して管理する発想である。

### 図

```mermaid
graph LR
    User[利用者] --> UI[フロントエンド]
    UI --> API[バックエンド/API]
    API --> DB[データベース]
    API --> Auth[認証/認可]
    API --> Log[ログ/監査]
    API --> Notify[通知]
```

### 混同しやすい近接概念

「アプリ」と「システム」は混同しやすい。アプリは利用者から見える入口を指すことが多い。システムは、アプリ、データ、権限、通知、運用まで含む全体を指す。

「自動化」と「システム化」も同じではない。自動化は一部の作業を自動で走らせること。システム化は、入力、処理、保存、権限、監査、運用まで含めて仕組みにすること。

「画面ができた」と「業務に使える」も同じではない。業務利用には、誰が使えるか、誰が承認するか、データが消えたらどうするか、退職者の所有物をどう引き継ぐか、監査で説明できるかが必要になる。

### ここを押さえれば次に進める

システムは画面だけではない。データ、処理、権限、監査、運用を含む。Dataverseを学ぶときも、テーブル設計だけでなく「誰が、どの画面から、どの権限で、どのAPIを通って、どのログを残すか」まで一続きで考える。

---

## 3. Webアプリの三層

### ふんわり入口

レストランで考えると、Webアプリの三層はわかりやすい。

- 客席: お客さんが見るメニューや注文端末
- 厨房: 注文を受けて調理する場所
- 倉庫: 食材や在庫を保管する場所

Webアプリでは、客席がフロントエンド、厨房がバックエンド、倉庫がデータベースである。

利用者は画面を見る。画面は裏側に「この顧客一覧をください」「この案件を登録してください」と依頼する。裏側はデータベースを読んだり書いたりして、結果を返す。

### 正確に言うと

Webアプリの代表的な構成は三層アーキテクチャである。

| 層 | 英語 | 役割 |
|---|---|---|
| プレゼンテーション層 | Frontend / UI | 利用者に画面を見せ、入力を受け取る |
| アプリケーション層 | Backend / Application | 業務ルール、API、認証連携、外部連携を処理する |
| データ層 | Database / Storage | データを保存し、検索や更新を行う |

Dataverseを使う場合、この三層の対応は構成によって変わる。

| 構成 | フロントエンド | バックエンド | データ層 |
|---|---|---|---|
| Canvas Apps | Canvas Apps | Power Platform/Dataverse Connector | Dataverse |
| Model-driven Apps | Model-driven Apps | Dataverse標準機能 | Dataverse |
| React + Web API直結 | React SPA | Dataverse Web API | Dataverse |
| React + BFF | React SPA | Azure Functions/App Service等 | Dataverse |
| Excel + Power Automate | Excel/VBA/Office Scripts | Power Automate | Dataverse |

### VBA的にいうと

VBAでは三層が1ファイルに寄りがちである。

```vb
' VBAでは画面、処理、保存が同じブック内に集まりやすい
Private Sub btnSave_Click()
    Worksheets("Data").Cells(nextRow, 1).Value = txtCustomerName.Value
    Worksheets("Data").Cells(nextRow, 2).Value = Now
End Sub
```

Webアプリでは、同じ処理が分かれる。

```javascript
// フロントエンド: ボタン押下でAPIに送る
async function saveCustomer() {
  const response = await fetch("/api/customers", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ name: document.querySelector("#name").value })
  });

  if (!response.ok) {
    throw new Error("登録に失敗しました");
  }
}
```

```sql
-- データ層: 実際にはDBに行として保存される
INSERT INTO Customers (CustomerName, CreatedOn)
VALUES ('株式会社サンプル', CURRENT_TIMESTAMP);
```

Dataverseでは直接このSQLを利用者が書くわけではない。代わりにTableに対してWeb APIやConnector経由で行を作る。

### 図

```mermaid
sequenceDiagram
    participant U as 利用者
    participant F as フロントエンド
    participant B as バックエンド/API
    participant D as データベース

    U->>F: 保存ボタンを押す
    F->>B: POST /customers
    B->>D: INSERT相当の処理
    D-->>B: 保存結果
    B-->>F: 201 Created
    F-->>U: 保存完了を表示
```

### 混同しやすい近接概念

「バックエンド」と「データベース」は別である。バックエンドは処理をする場所、データベースは保存する場所である。ただしDataverseはWeb API、権限、Plug-in、Custom APIなどを持つため、単なるデータベースよりバックエンド寄りの機能も持つ。

「フロントエンド」と「アプリ」も文脈で揺れる。Power Appsで「アプリ」と言うとCanvas AppやModel-driven Appを指すことが多い。Web開発で「フロント」と言うとReactやVueなどの画面部分を指すことが多い。

「サーバーレス」と「バックエンドなし」も違う。Azure Functionsのようなサーバーレスは、サーバー管理をクラウドに任せるという意味であり、裏側の処理が消えるわけではない。

### ここを押さえれば次に進める

Webアプリは、画面、処理、データの三層で見ると整理しやすい。Dataverseはデータ層でありながら、API、権限、イベント処理も持つ。そのため「DataverseをDBとしてだけ見る」と見落としが出る。

---

## 4. データベース入門

### ふんわり入口

データベースは、きれいに整理された台帳である。Excelの表を想像すると入りやすい。

たとえば顧客一覧がある。

| 顧客ID | 顧客名 | 電話番号 |
|---|---|---|
| C001 | 株式会社サンプル | 03-0000-0000 |
| C002 | 鈴木商店 | 06-0000-0000 |

この表で、横1行が1人の顧客、縦1列が顧客名や電話番号などの項目である。データベースでも基本は同じで、表、行、列で考える。

ただし業務データが大きくなると、Excelの表だけでは苦しくなる。顧客、案件、見積、請求、担当者を全部1枚のシートに入れると、同じ顧客名が何度も出たり、修正漏れが起きたり、どの行が何を表しているのかわからなくなる。そこで、テーブルを分け、IDでつなぐ。

### 正確に言うと

リレーショナルデータベースでは、データをテーブルとして管理する。

| 一般DB用語 | Dataverse用語 | 旧名/近い言葉 |
|---|---|---|
| Table | Table | Entity |
| Column | Column | Field / Attribute |
| Row | Row | Record |
| Primary Key | Primary Key | 主キー |
| Foreign Key | Lookup / Relationship | 外部キー的な関係 |
| Relationship | Relationship | リレーション |

主キーは、行を一意に識別する値である。社員番号、顧客ID、注文IDのようなものだ。外部キーは、別のテーブルの行を指す値である。DataverseではLookup列とRelationshipで表現される。

正規化は、データの重複や不整合を減らすためにテーブルを適切に分ける考え方である。

例として、悪い表を考える。

| 注文ID | 顧客名 | 顧客電話 | 商品名 | 単価 |
|---|---|---|---|---|
| O001 | 株式会社サンプル | 03-0000-0000 | マウス | 2000 |
| O002 | 株式会社サンプル | 03-0000-0000 | キーボード | 5000 |

顧客電話が変わると、同じ顧客の全行を直す必要がある。そこで顧客テーブルと注文テーブルに分ける。

```sql
CREATE TABLE Customers (
  CustomerId varchar(20) PRIMARY KEY,
  CustomerName varchar(100) NOT NULL,
  Phone varchar(30)
);

CREATE TABLE Orders (
  OrderId varchar(20) PRIMARY KEY,
  CustomerId varchar(20) NOT NULL,
  ProductName varchar(100) NOT NULL,
  UnitPrice int NOT NULL,
  FOREIGN KEY (CustomerId) REFERENCES Customers(CustomerId)
);
```

トランザクションは、複数の更新を「全部成功」または「全部失敗」にする単位である。銀行振込で、Aさんの口座から1万円を引いたのにBさんの口座に足されない、という途中失敗を防ぐ考え方である。

```sql
BEGIN TRANSACTION;

UPDATE Accounts
SET Balance = Balance - 10000
WHERE AccountId = 'A';

UPDATE Accounts
SET Balance = Balance + 10000
WHERE AccountId = 'B';

COMMIT;
```

Dataverseでは利用者が自由にSQLトランザクションを書くわけではないが、Web APIの`$batch`、Plug-in、Custom API、Organization Serviceなどで複数処理の整合性を考える場面がある。

### VBA的にいうと

`Worksheet.Cells(1, 1)`は、表の1行目1列目を指す。DBでは、通常「1行目」のような物理的な順番に意味を持たせない。代わりに主キーで行を特定する。

```vb
' Excel的な発想
Worksheets("Customers").Cells(2, 1).Value = "C001"
Worksheets("Customers").Cells(2, 2).Value = "株式会社サンプル"
```

```sql
-- DB的な発想
SELECT CustomerName
FROM Customers
WHERE CustomerId = 'C001';
```

VBAのRecordsetは、DBから取ってきた行のまとまりである。JavaScriptでAPIを呼んだ場合は、JSON配列がRecordsetに近い。

```javascript
const result = await fetch("/api/customers").then(r => r.json());
for (const row of result.value) {
  console.log(row.name);
}
```

### 図

```mermaid
erDiagram
    CUSTOMER ||--o{ ORDER : places
    CUSTOMER {
        string customer_id PK
        string name
        string phone
    }
    ORDER {
        string order_id PK
        string customer_id FK
        string product_name
        int unit_price
    }
```

### 混同しやすい近接概念

「表」と「テーブル」は近いが、Excelの表とDBのテーブルは同じではない。Excelでは見た目、計算式、メモ、結合セルなどが混ざりやすい。DBのテーブルは型、制約、主キー、関係を持つ。

「外部キー」と「Lookup」も近いが、DataverseではLookupはUIや権限、Relationship、参照先のEntity Set名などと絡む。単なるID列ではない。

「削除」と「非アクティブ化」も違う。業務システムでは、行を物理削除せず、状態列で無効化する方が監査や参照整合性に向く場合が多い。Dataverseの多くの標準Tableにも状態を表す列がある。

### ここを押さえれば次に進める

DBの基本はTable、Row、Column、主キー、外部キー、Relationshipである。Dataverseではこれらが公式用語としてTable、Column、Row、Lookup、Relationshipに対応する。Excelの表に似ているが、主キー、権限、関係、監査、APIが乗る点が大きく違う。

---

## 5. APIと通信

### ふんわり入口

APIは、システム同士の注文窓口である。

レストランで言えば、客が厨房に直接入って冷蔵庫を開けるのではなく、店員さんに「カレーをください」と頼む。店員さんは注文を受け、厨房に伝え、料理を客席へ戻す。

Webシステムでも、画面がデータベースを直接触るのではなく、APIに「顧客一覧をください」「案件を登録してください」と頼むことが多い。APIは頼み方のルールを持った窓口である。

### 正確に言うと

HTTPはWebの基本通信プロトコルである。HTTPSはHTTPをTLSで暗号化したもの。現代のWeb APIは基本的にHTTPSで通信する。

HTTPにはメソッドがある。

| メソッド | 意味 | 典型用途 |
|---|---|---|
| GET | 取得する | 一覧取得、詳細取得 |
| POST | 作成する、処理を依頼する | 新規作成、Action呼び出し |
| PATCH | 部分更新する | 一部列の更新 |
| PUT | 置き換える | 全体更新 |
| DELETE | 削除する | 行の削除 |

リクエストは依頼、レスポンスは返事である。

```http
GET /api/data/v9.2/accounts?$select=name,telephone1 HTTP/1.1
Host: org.crm.dynamics.com
Authorization: Bearer <access_token>
Accept: application/json
```

返事にはステータスコードが付く。

| ステータス | 意味 |
|---|---|
| 200 OK | 成功 |
| 201 Created | 作成成功 |
| 204 No Content | 成功、返す本文なし |
| 400 Bad Request | 依頼の書き方が悪い |
| 401 Unauthorized | 認証できていない |
| 403 Forbidden | 権限がない |
| 404 Not Found | 見つからない |
| 409 Conflict | 競合 |
| 412 Precondition Failed | ETagなど前提条件に失敗 |
| 429 Too Many Requests | リクエスト過多 |
| 500 Internal Server Error | サーバー側エラー |

RESTはHTTPの使い方の作法である。プロトコルそのものではない。「リソースにURLを割り当て、HTTPメソッドで操作する」という考え方である。

ODataは、データを問い合わせるための規格で、`$select`、`$filter`、`$expand`、`$top`などのクエリを使う。Dataverse Web APIはOData v4系の作法で操作する。

```javascript
const url = "https://org.crm.dynamics.com/api/data/v9.2/accounts"
  + "?$select=name,telephone1"
  + "&$filter=startswith(name,'A')";

const response = await fetch(url, {
  headers: {
    Authorization: `Bearer ${accessToken}`,
    Accept: "application/json"
  }
});

if (!response.ok) {
  throw new Error(`${response.status} ${response.statusText}`);
}

const data = await response.json();
console.log(data.value);
```

### VBA的にいうと

VBAからもHTTPは呼べる。ADOでDBに接続する感覚に近いが、接続先はSQL ServerではなくWeb APIである。

```vb
Dim http As Object
Set http = CreateObject("MSXML2.XMLHTTP")

http.Open "GET", "https://example.com/api/customers", False
http.setRequestHeader "Accept", "application/json"
http.Send

If http.Status = 200 Then
    Debug.Print http.responseText
End If
```

ただし、Dataverse Web APIをVBAから直接安全に呼ぶにはOAuth、MFA、条件付きアクセス、トークン保管が絡む。会社本番ではここが非常に重い。VBAから直接Dataverseへ行くより、Power AutomateやBFFを挟む方が現実的な場合が多い。

### 図

```mermaid
sequenceDiagram
    participant FE as 画面
    participant API as API窓口
    participant DV as Dataverse

    FE->>API: GET /accounts
    API->>DV: account行を検索
    DV-->>API: JSON
    API-->>FE: 200 OK + JSON
```

### 混同しやすい近接概念

「HTTP」と「REST」は違う。HTTPは通信のルール、RESTはHTTPをどう使うかの設計作法である。

「REST」と「OData」も違う。RESTは広い設計スタイル、ODataはデータ問い合わせの具体的な規格である。Dataverse Web APIはREST風に見えるが、ODataの書式を多く使う。

「API」と「Connector」も違う。APIはシステムが公開する窓口。ConnectorはPower PlatformからAPIを使いやすくする接続部品である。Dataverse Connectorの裏側ではDataverseのAPIが使われるが、利用者は直接HTTPを書かずに扱える。

### ここを押さえれば次に進める

APIはシステム同士の窓口である。HTTPメソッド、リクエスト、レスポンス、ステータスコードを理解すると、Dataverse Web APIやConnectorのエラーが読めるようになる。DataverseではODataの語彙がよく出るため、`$select`、`$filter`、`$expand`は早めに慣れる。

---

## 6. 認証と認可

### ふんわり入口

認証と認可は、会社の受付と入館証で考えるとわかりやすい。

受付で「あなたは誰ですか」と確認するのが認証である。社員証、免許証、顔認証、パスワード、MFAがこれに当たる。

その後、「あなたは3階の経理部屋に入れますか」「金庫を開けられますか」と判断するのが認可である。同じ社員でも、全員が全室に入れるわけではない。

Dataverseでは、Entra IDが主に「あなたは誰か」を担当し、DataverseのSecurity RoleやBusiness Unitなどが「何ができるか」を担当する。

### 正確に言うと

認証(Authentication)は、主体が誰であるかを確認すること。認可(Authorization)は、認証済みの主体に対して、どの操作を許可するかを決めること。

OAuth 2.0は、パスワードを直接渡さず、限定された権限を表すトークンでAPIを使うための仕組みである。OpenID ConnectはOAuth 2.0の上にID情報を扱う層を加えたもので、ログインした人のIDを扱う。

トークンは、ホテルのカードキーに近い。マスターキーではなく、期限や入れる部屋が決まったカードを渡す。

| 用語 | 意味 |
|---|---|
| Access Token | APIを呼ぶための短命な通行証 |
| Refresh Token | Access Tokenを再取得するための通行証 |
| Scope | 何をしてよいかの範囲 |
| Client ID | アプリを識別するID |
| Client Secret | アプリが秘密に持つ合言葉 |
| Tenant ID | Microsoft 365テナントを識別するID |
| Authorization Code Flow | 人間がログインしてアプリに代理権限を渡す流れ |
| Client Credentials Flow | アプリ自身がサーバー間で動く流れ |
| On-Behalf-Of | 受け取ったユーザートークンをもとに、バックエンドがユーザーの代わりに下流APIを呼ぶ流れ |

Authorization Code Flowの大まかな流れは次の通り。

```mermaid
sequenceDiagram
    participant U as 利用者
    participant A as アプリ
    participant E as Entra ID
    participant API as API

    U->>A: アプリを開く
    A->>E: ログインへリダイレクト
    U->>E: ID/パスワード/MFA
    E-->>A: 認可コード
    A->>E: 認可コードをトークンに交換
    E-->>A: Access Token
    A->>API: Bearer Token付きで呼び出し
```

Client Credentials Flowは、人間が画面でログインしないバッチやサーバー間連携で使う。

```javascript
// Node.jsでClient Credentialsのトークンを取る例
const body = new URLSearchParams({
  grant_type: "client_credentials",
  client_id: process.env.CLIENT_ID,
  client_secret: process.env.CLIENT_SECRET,
  scope: "https://org.crm.dynamics.com/.default"
});

const tokenResponse = await fetch(
  `https://login.microsoftonline.com/${process.env.TENANT_ID}/oauth2/v2.0/token`,
  {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body
  }
);

const { access_token } = await tokenResponse.json();
```

### VBA的にいうと

VBAの小さなツールでは、Windowsログイン済みの前提で動いたり、ブックを開ける人なら使える、という形になりやすい。これは「認証と認可をExcelファイル運用に寄せている」状態である。

業務システムでは、認証と認可を明示する。

| VBA運用 | Web/Dataverse運用 |
|---|---|
| 共有フォルダに置いた人だけ開ける | Entra IDでログインする |
| シート保護パスワード | Security RoleやColumn-level security |
| マクロ内に接続文字列 | App Registration、Secret、Key Vault |
| 担当者別ファイル配布 | Business Unit、Team、共有 |

特に、VBAにClient Secretを埋め込む設計は避ける。ブックを配ることはSecretを配ることに近くなる。

### 混同しやすい近接概念

「ログインできる」と「データを読める」は違う。Entra IDでログインできても、Dataverse側にSystem Userがあり、Security Roleがあり、対象TableへのRead権限がなければ読めない。

「401」と「403」も違う。401は認証ができていない。403は認証はできたが権限がない。DataverseではApplication Userを作り忘れた場合やSecurity Role不足で、この違いが重要になる。

「Application User」と「Service Principal」も違う。Service PrincipalはEntra ID側のアプリ実体。Application UserはDataverse側でそのアプリをユーザーとして扱うための登録である。

### ここを押さえれば次に進める

認証は「あなた誰」、認可は「何してよい」。Dataverse連携ではEntra IDとDataverseの2世界をまたぐ。App RegistrationだけではDataverse権限は得られない。Application UserとSecurity Roleまでそろって初めて、サーバー間連携が動く。

---

## 7. フロントエンドの形

### ふんわり入口

フロントエンドは、利用者が触る入口である。店舗で言えば売り場、窓口、注文端末、案内板である。

同じ商品を売るにも、店舗、電話注文、自動販売機、ECサイト、スマホアプリでは入口が違う。裏側の在庫や会計は共通でも、利用者が触る形は変えられる。

Dataverseも同じで、同じTableをCanvas Apps、Model-driven Apps、Power Pages、Teams Tab、Excel、Reactアプリなど、さまざまな入口から扱える。

### 正確に言うと

フロントエンドの代表的な形には、SPA、SSR、SSGなどがある。

| 方式 | 意味 | 向く場面 |
|---|---|---|
| SPA | Single Page Application。ブラウザ内で画面遷移を処理する | 業務アプリ、管理画面 |
| SSR | Server-Side Rendering。サーバーでHTMLを生成して返す | 初期表示速度、SEO、認証付きサイト |
| SSG | Static Site Generation。事前にHTMLを生成する | ドキュメント、ブログ、静的サイト |

React、Vue、Angular、Svelteなどはフロントエンドフレームワークまたはライブラリである。UI部品、状態管理、画面遷移、ビルドなどを助ける。

Dataverse周辺では、フロントエンドの選択肢をローコード度合いで見ると整理しやすい。

| 層 | 例 | 特徴 |
|---|---|---|
| ローコード | Canvas Apps、Model-driven Apps | Power Platform内、認証/権限/共有が標準寄り |
| 部分コード | PCF、Web resource、Command bar JS | 標準アプリの一部を拡張 |
| フルコードだがPower Platform内 | Code Apps | React等で書き、Power Platform管理下で動かす |
| 完全独立 | React/Vue + Dataverse Web API/BFF | 自由度が高いが認証、ホスティング、審査が重い |

### VBA的にいうと

VBAのUserFormはフロントエンドである。

```vb
Private Sub btnSearch_Click()
    Call SearchCustomer(txtKeyword.Value)
End Sub
```

JavaScriptではイベントリスナーで似たことをする。

```javascript
document.querySelector("#searchButton").addEventListener("click", async () => {
  const keyword = document.querySelector("#keyword").value;
  const result = await fetch(`/api/customers?keyword=${encodeURIComponent(keyword)}`)
    .then(r => r.json());
  renderCustomers(result.value);
});
```

Power Fxでは、ボタンの`OnSelect`に式を書く。

```powerfx
ClearCollect(
    colAccounts,
    Filter(Accounts, StartsWith('Account Name', txtKeyword.Text))
)
```

### 図

```mermaid
graph TD
    DV[Dataverse]
    Canvas[Canvas Apps]
    Model[Model-driven Apps]
    Pages[Power Pages]
    PCF[PCF]
    Code[Code Apps]
    SPA[独自React SPA]
    Excel[Excel/Office Add-ins]

    Canvas --> DV
    Model --> DV
    Pages --> DV
    PCF --> Model
    PCF --> Canvas
    Code --> DV
    SPA --> DV
    Excel --> DV
```

### 混同しやすい近接概念

「画面自由度」と「本番通過性」は別である。Reactで独自UIを作れば自由度は高いが、Entraアプリ登録、ホスティング、条件付きアクセス、DLP、監査、ライセンスの説明が必要になる。

「Canvas Apps」と「Model-driven Apps」は同じPower Appsでも発想が違う。Canvasは自由配置の画面を作る。Model-drivenはDataverseのデータモデル、フォーム、ビュー、サイトマップを中心に自動生成寄りで作る。

「PCF」と「Code Apps」も違う。PCFはCanvas/Model-drivenに埋め込む部品。Code Appsはアプリ全体をコードで作る方式である。

### ここを押さえれば次に進める

フロントエンドは入口である。Dataverseを中心にする場合、入口は一つではない。最初に「誰が、どの端末で、どの業務導線から使うか」を決めると、Canvas、Model-driven、Teams、SharePoint、独自Web、Excelのどれが妥当か判断しやすくなる。

---

## 8. セキュリティの基本

### ふんわり入口

セキュリティは、玄関の鍵だけではない。会社なら、受付、社員証、部屋の鍵、金庫、監視カメラ、入退室ログ、持ち出し禁止ルール、廃棄ルールまで含む。

Webアプリでも同じで、ログイン画面だけ作れば安全というわけではない。入力値、ブラウザ、API、権限、外部サイト、通信、ログ、端末、運用まで見る必要がある。

Dataverseは標準で強いセキュリティモデルを持つが、フロントエンドや連携方法を間違えると、別の穴を作る。特に独自Web、Office Add-ins、VBA、Power Automate HTTP、Custom Connectorを使う場合は注意が必要である。

### 正確に言うと

Webセキュリティでよく出る基本用語を整理する。

| 用語 | 日常の例え | 意味 |
|---|---|---|
| CSRF | 本人の印鑑を勝手に押される | ログイン済みブラウザに意図しないリクエストを送らせる攻撃 |
| XSS | 掲示板に悪意ある貼り紙を貼る | サイト内に悪意あるJavaScriptを実行させる攻撃 |
| SQL Injection | 申請欄に命令文を書き込む | 入力値を悪用してSQLを改ざんする攻撃 |
| CORS | 入館できる取引先ドメインの許可リスト | ブラウザが別オリジンAPIを呼ぶ制御 |
| CSP | 店内で実行してよいスクリプトの規則 | 読み込めるスクリプトや画像の制限 |

SQL Injectionの危険な例。

```javascript
// 悪い例: 入力値をSQLに直接連結する
const sql = "SELECT * FROM Users WHERE name = '" + userInput + "'";
```

安全な例。

```javascript
// 良い例: パラメータ化する
const result = await db.query(
  "SELECT * FROM Users WHERE name = ?",
  [userInput]
);
```

Dataverse Web APIを使う場合、通常は自分でSQLを組み立てないため、典型的なSQL Injectionとは形が違う。ただし、`$filter`文字列を入力値から組み立てる場合、エスケープや許可リスト設計が必要になる。

```javascript
// 悪い例: 入力値をOData filterにそのまま連結する
const url = `/api/data/v9.2/accounts?$filter=name eq '${keyword}'`;

// ましな例: シングルクォートをエスケープし、許可する検索条件を限定する
const escaped = keyword.replaceAll("'", "''");
const safeUrl = `/api/data/v9.2/accounts?$filter=startswith(name,'${escaped}')`;
```

CORSはブラウザの制御である。サーバー間通信では同じ制約はない。独自SPAからDataverseを直接叩く場合、ブラウザ、リダイレクトURI、許可オリジン、認証フローが絡むため、本番審査が重くなりやすい。

### VBA的にいうと

VBAツールでは、次のようなセキュリティ問題が起きやすい。

| VBAでありがちなこと | Web/Dataverseでの危険 |
|---|---|
| 接続文字列やパスワードをコードに直書き | Secret漏洩 |
| 共有フォルダに最新版を置くだけ | 改ざん、版数不明 |
| マクロ有効化を手作業で依頼 | 端末制御、監査、署名問題 |
| 入力値をSQL文字列に連結 | SQL Injection |
| 全員が同じ共有アカウントで実行 | 監査不能、Multiplexing問題 |

VBAが悪いわけではない。ローカル少人数の業務には強い。ただしDataverseのような中央データ基盤につなぐなら、認証、監査、トークン、ライセンス、DLPを含めて設計する必要がある。

### 図

```mermaid
graph TD
    Browser[ブラウザ/Office WebView] -->|HTTPS| API[API/Dataverse]
    Browser -->|悪意あるScript混入| XSS[XSSリスク]
    OtherSite[別サイト] -->|ログイン済み状態を悪用| CSRF[CSRFリスク]
    API --> Auth[認証/認可]
    API --> Log[監査ログ]
    API --> DLP[DLP/データポリシー]
```

### 混同しやすい近接概念

「暗号化」と「安全」は違う。HTTPSで通信していても、権限設計が間違っていれば見えてはいけないデータが見える。

「認証」と「入力検証」も別である。ログイン済みユーザーでも、入力値は信用しない。業務アプリでは、悪意だけでなく誤入力も事故の原因になる。

「DataverseのSecurity Role」と「画面上の非表示」も違う。画面でボタンや列を隠しても、API権限が残っていれば別経路で更新できる可能性がある。重要な制約はサーバー側、つまりDataverseの権限やPlug-in/Custom API側で守る。

### ここを押さえれば次に進める

セキュリティはログインだけではない。入力、API、権限、通信、ブラウザ制御、監査、DLP、端末管理まで含む。Dataverseを使う場合も、標準機能に任せる部分と、自分で守る部分を分けて考える。

---

