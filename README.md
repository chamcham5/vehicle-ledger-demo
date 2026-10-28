# 車両台帳・PDF照合から月次CSVまでの架空デモ

車両Excel、購入契約PDF、割賦PDFの3ファイルを照合し、支払表Excelと指定月CSVを作る小さなサンプルです。入力と金額はすべて架空です。実顧客の帳票や受注実績を示すものではありません。

生成AIの支援を受けて制作しました。実行時は固定形式の文字列を読み取る処理で、外部AIサービスやクラウドへデータを送信しません。

## 閲覧だけで確認する

- [車両台帳](vehicles.xlsx)：DEMO-A／B／Cの3台。番号はTOKYO-500-TEST-01〜03です。
- [購入契約の例](contracts.pdf)：各ページに架空データであることを明記。台帳の車両番号・登録番号・型式と照合します。
- [割賦明細の例](installments.pdf)：3台・計6回分。支払回数・日付・金額・契約合計を照合します。

2027年1月を指定した場合、2明細・合計225,001円のCSVを作ります。全月の支払表には6明細・合計545,000円を保存します。

```csv
payment_date,vehicle_id,registration,amount_jpy
2027-01-15,DEMO-C,TOKYO-500-TEST-03,125000
2027-01-31,DEMO-A,TOKYO-500-TEST-01,100001
```

登録番号を食い違わせた入力では、`Vehicle/contract mismatch`で停止し、出力先フォルダーを作りません。入力を手で修正した後に、新しい実行で再確認できます。生成したCSVは、新しく保存した支払表Excelを再読した内容から作ります。

## 自分で再現する

Python 3.12.14で確認しています。必要なライブラリのインストールにはネット接続が必要です。専用の仮想環境で、リポジトリ直下から次を実行します。

```text
python -m venv .venv
# Windows: .venv\Scripts\activate
# macOS/Linux: source .venv/bin/activate
python -m pip install -r requirements.txt
python -B verify_demo.py
```

`verify_demo.py`は、正常入力・不一致での停止・手修正後の復帰・同じ保存Excelからの4か月分を一時フォルダーで確認し、結果を表示します。一時ファイルは終了時に削除します。既存の出力や別の私的ファイルは不要です。

結果を残す場合は、まだ存在しない出力先を指定します。

```text
python -B pipeline.py --vehicles vehicles.xlsx --contracts contracts.pdf --installments installments.pdf --month 2027-01 --out demo-output
```

`demo-output`に支払表、CSV、入力照合レポートを保存します。既存の出力先には上書きしません。ファイルI/Oの途中で失敗した場合は一部出力が残る可能性があるため、内容を確かめず自動再試行しないでください。

## 確認できる範囲

このサンプルは、同梱した`CONTRACT_V1`／`SCHEDULE_V1`のテキストPDFと固定CSV形式に対応します。任意の帳票・スキャン画像のOCR、実際の会計ソフトへの取込、GUI、現場への導入は未対応・未確認です。支払表の集計は生成時の保存値で、手編集後の自動再計算は提供しません。

別形式へ適用する場合は、入力様式、例外時の運用、CSV列、受入条件、実行環境を別途確認する必要があります。実データをこの公開リポジトリへ投稿しないでください。

## 実業務への適用を検討する方へ

このデモと同じ「車両台帳と契約・支払予定を照合し、月次CSVへまとめる作業」について、有料の個別対応を検討する場合は、まず次の条件を確認します。

- 月ごとの処理件数、帳票の種類数、文字入りPDFかスキャン画像か。
- 必要なCSVの列名・文字コード・対象月の決め方。列の例は架空の値に置き換えてください。
- 不一致を誰が確認・修正するか、利用するOSと実行環境、納品物の受入条件。

[公開Issueで、仕様の概要を質問できます](https://github.com/chamcham5/vehicle-ledger-demo/issues/new/choose)。投稿は第三者も閲覧できます。会社名・氏名・登録番号・実際の帳票や顧客データを載せず、架空の項目名と件数だけで説明してください。

この公開デモの利用は無料です。実帳票への対応、導入、保守はデモに含まれません。個別対応の可否・費用・納期は仕様を確認して別途判断するもので、このページは受注や動作を保証するものではありません。

## ソースと依存物

本リポジトリの自作ソースと架空サンプルは[MIT License](LICENSE)です。ライブラリ本体、Python実行環境、インストーラーは同梱しません。依存物にはそれぞれの著作権・許諾が適用されます。

- [pypdf 6.10.0の公式許諾](https://github.com/py-pdf/pypdf/blob/6.10.0/LICENSE)
- [openpyxl 3.1.5の公式パッケージ](https://pypi.org/project/openpyxl/3.1.5/)
- [et-xmlfile 2.0.0の公式パッケージ](https://pypi.org/project/et-xmlfile/2.0.0/)
- [Python 3.12の公式許諾](https://docs.python.org/3.12/license.html)

依存作者による推薦・動作保証を示すものではありません。
