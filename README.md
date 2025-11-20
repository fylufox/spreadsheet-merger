<!--
Copyright 2025 Google LLC

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

      http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
-->
# スプレッドシート統合ツール

外部キーで各シートに分割されているデータを一つのシートに統合するGoogle Apps Scriptツールです。

## 機能

- 複数のシートから外部キーに基づいてデータを統合
- 設定シートで統合するデータを柔軟に指定
- 自動的にカラム名にシート名プレフィックスを付与

## 使い方

### 1. デプロイ

```bash
npm install
npm run deploy
```

### 2. 設定シートの作成

スプレッドシートを開くと、メニューバーに「シート統合ツール」メニューが表示されます。

1. 「シート統合ツール」→「設定テンプレート作成」を選択
2. 「設定」シートが作成されます

### 3. 設定シートの編集

#### 設定シートの構造

| 項目 | 内容 |
|------|------|
| **A1セル** | "外部カラム" (固定) |
| **B1セル** | 外部キーとして使用するカラム名 (例: "id") |
| **A2セル** | "探索範囲(行)" (固定) |
| **B2セル** | 外部キーカラムを探索する範囲(行数)、デフォルト: 10 |
| | |
| **A3セル** | "シート名" (ヘッダー) |
| **B3セル** | "カラム名" (ヘッダー) |
| **A4セル以降** | 統合したいシート名(同じシート名を複数行に記入可能) |
| **B4セル以降** | 各シートから取得したいカラム名(1行に1カラム) |

#### 設定例

```
外部カラム           | id
探索範囲(行)       | 10
                   |
シート名            | カラム名
ユーザー            | name
ユーザー            | email
ユーザー            | department
注文                | order_date
注文                | amount
注文                | status
配送                | shipping_address
配送                | delivery_date
```

この設定により、`id`をキーとして以下のように統合されます:

- 各シートの先頭10行以内で`id`カラムを探索
- `ユーザー`シートから: `name`, `email`, `department`
- `注文`シートから: `order_date`, `amount`, `status`
- `配送`シートから: `shipping_address`, `delivery_date`

### 4. データ統合の実行

1. 「シート統合ツール」→「データ統合実行」を選択
2. 「統合データ」シートが作成され、統合結果が出力されます

### 出力形式

出力シートのカラム名は `シート名.カラム名` の形式になります:

| id | ユーザー.name | ユーザー.email | ユーザー.department | 注文.order_date | 注文.amount | 注文.status | 配送.shipping_address | 配送.delivery_date |
|----|--------------|----------------|-------------------|----------------|------------|------------|---------------------|-------------------|
| 1  | 山田太郎      | yamada@...     | 営業部             | 2025-01-15     | 10000      | 完了        | 東京都...            | 2025-01-20        |
| 2  | 佐藤花子      | sato@...       | 開発部             | 2025-01-16     | 25000      | 処理中      | 大阪府...            | 2025-01-22        |

## データ要件

### 各シートの形式

- ヘッダー行: 外部キーカラム名を含む行(探索範囲内に必須)
- データ行: ヘッダー行の次の行から
- 探索範囲: 各シートの先頭から指定した行数までを探索
- データが途中の行から始まる場合も対応可能

### データ統合のルール

- 外部キーが一致する行を横に結合
- 該当データがない場合は空欄
- 外部キーが重複している場合は、最後の行のデータが使用されます

## トラブルシューティング

### エラー: 設定シートが見つかりません

→ 「設定テンプレート作成」を実行してください

### エラー: 外部キーカラム名が設定されていません

→ 設定シートのB1セルに外部キーカラム名を入力してください

### エラー: データソースが設定されていません

→ 設定シートの4行目以降にシート名とカラム名を入力してください(1行に1カラム)

### 警告: シート "XXX" の探索範囲内に外部キーカラム "YYY" が見つかりません

→ 探索範囲を増やすか、シート内の外部キーカラム名を確認してください

### 警告: シート "XXX" が見つかりません

→ 設定シートで指定したシート名が存在するか確認してください

### 警告: カラム "XXX" が見つかりません

→ 指定したカラム名がシートに存在するか確認してください

## ライセンス

Apache License 2.0
