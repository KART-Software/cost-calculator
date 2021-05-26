# cost-calculator

>このパッケージは開発中で、安定しておりません。バグを見つけた方は報告をお願いします。

cost-calculator は、学生フォーミュラ(FSAEJ)におけるコスト審査のファイル作成を支援するツールです。
以下の機能があります。

* コストテーブル(.xlsx)からFCAファイル(.xlsx)への"Unit Cost"の書き込み。（開発中）
* FCAファイルからBOMファイル(.xlsx)への"Quantity"、"Material Cost"、"Process Cost"、"Fastener Cost"、"Toolin Cost"、"Link to FCA Sheet"の書き込み。
* FCAファイルへの、裏付け資料(.pdf)へのリンクの書き込み。

## 互換性
* Pythonのバージョンは3.7.0以上を使用してください。

## インストール
最初に、バージョン3.7.0以上のpython環境、pipを使える環境を用意してください。

その後、
```
$ pip install git+https://github.com/KART-Software/cost-calculator
```
を実行してください。

## 使用方法
使用する前に、まずバックアップを取ってください。バックアップを取らずに実行して、ファイルが破損もしくは修復不能になった場合、一切の責任を負いません。

次に、ディレクトリ構造が以下の構造になっていることを確認してください。(詳しくは配布されている"2021_fsaej_localrules02_e.pdf"を参照してください。)

![GORILLA](https://i.gzn.jp/img/2018/01/15/google-gorilla-ban/00.jpg)

使用方法は２通りあります。

### コマンドラインアプリとして使う場合(現在は仕様不可)

### pythonライブラリとして使う方法
まずコマンドラインで、
```
$ python
```
を実行してpythonのshellに入ってください。抜けるには
```python
>>> quit()
```
を実行します。

以下、`$`と`>>>`で、実行環境を区別します。

* コストテーブルからFCAファイルへの"Unit Cost"の書き込み。（開発中です。使わないでください。）

使用例
```python
$ python
>>> from cost_calculator import costTableToFca
>>> costTableToFca("example/cost_table_files", "example/fca_files_empty")
```

`costTableToFca(,)`は２つの引数を取る関数です。

第１引数には、"Material、Process、ProcessMultiplier、Fastener、Toolingの全５種類のコストテーブルが入ったフォルダのパス"、第２引数には"複数のFCAファイルが入ったフォルダのパス"を指定してください。

"コストテーブルが入ったフォルダのパス"の中には他のファイルを含めても構いませんが、必ず全５種類のコストテーブルのファイル(.xlsx)を含めてください。

"FCAファイルが入ったフォルダのパス"は、その中のフォルダにFCAが入っていても構いません。３階層下まで、FCAファイルを探索して、FCAファイルに書き込みます。また、他のファイルがフォルダ内に入っていても、自動で除外するので、気にしなくて大丈夫です。

* FCAファイルからBOMファイルへのデータの書き込み

使用例
```python
>>> from cost_calculator import fcaToBom
>>> fcaToBom("example/fca_files", "example/BrakeSystem_BOM.xlsx")
```
`fcaToBom(,)`は２つの引数を取る関数です。

の第１引数には、"複数のFCAファイルが入ったフォルダのパス"、第２引数には"BOMファイルのパス"を指定してください。

"FCAファイルが入ったフォルダのパス"は、その中のフォルダにFCAが入っていても構いません。３階層下まで、FCAファイルを探索して、FCAファイルに書き込みます。また、他のファイルがフォルダ内に入っていても、自動で除外するので、気にしなくて大丈夫です。

"BOMファイルのパス"は、フォルダのパスではなく、単一ファイルのパスをしてください。

* FCAファイルへ、裏付け資料へのリンクの書き込み

使用例
```python
>>> from cost_calculator import supplToFca
>>> supplToFca("example/fca_files")
```
`supplToFca()`は１つだけ引数を取る関数です。

引数には、"複数のFCAファイルが入ったフォルダのパス"を指定してください。

"FCAファイルが入ったフォルダのパス"は、その中のフォルダにFCAが入っていても構いません。３階層下まで、FCAファイルを探索して、FCAファイルに書き込みます。また、他のファイルがフォルダ内に入っていても、自動で除外するので、気にしなくて大丈夫です。

注）PDFファイルの中身を探索するのに、pythonライブラリのpdfminer.sixを使用しています。探索に少し時間がかかることがあります。
___
## 開発環境

poetryでパッケージ管理をしています。

参考

https://org-technology.com/posts/python-poetry.html
https://cocoatomo.github.io/poetry-ja/
https://kk6.hateblo.jp/entry/2018/12/20/124151

### 推奨

環境構築前に以下のコマンドで、ワークスペースディレク鳥に仮想環境を作成できます。
```
$ poetry config virtualenvs.in-project true
```

### poetry環境構築
```
$ poetry install
```

## それぞれの機能お試し
```
$ make example-ctf
```

```
$ make example-ftb
```

```
$ make example-stf
```