# DBからSelectのSQLだけ実行するプログラム

## 環境
 - altgraph 0.17.4
 - packaging 24.1
 - pefile 2023.2.7
 - pyinstaller 6.11.0
 - pyinstaller-hooks-contrib 2024.9
 - pyodbc 5.1.0
 - pywin32-ctypes 0.2.3
 - sqlparse 0.5.1

## 使用方法
1. `config.ini`に接続情報を入力
2. `***.sql`に実行したいSQLを記載
3. `dbSelect.py`を実行し、`***.sql`を指定
3. `SQL_RESULT_yyyymmdd_hhmm_ss`に以下が出力される
    - `[連番]_input.sql`
    - `[連番]_output.csv` もしくは `[連番]_[コメント].csv`

## ファイル仕様
### `config.ini`
接続時に使用する設定を記載する。  
Accessの場合、[外部データ] > [リンクテーブルマネージャー] > 接続しているデータソース名 > [編集] > [接続文字列] で確認可能。
```
DSN=DSN_SAMPLE
DATABASE=sample
DB=sample
DFLT_BIGINT_BIND_STR=1
DRIVER=MySQL ODBC 9.1 Unicode Driver
NO_SCHEMA=1
PASSWORD=password
PORT=3306
PWD=password
SERVER=192.168.56.103
UID=root
USER=root
```

### `***.sql`
実行するSQLを記載する。  
セミコロンの直後に記載したコメントは実行結果のファイル名となる。
#### OK
```
SELECT * FROM `sample`.`data1`;
SELECT * FROM `sample`.`data2`; -- data2output
SELECT
 * 
FROM
  `sample`.`data3`; -- data3output
```
#### NG
```
-- Accessのみのクエリを含むSQLをMySQLで実行している(Switch関数など)
SELECT sample.data1.uid, Switch( sample.data1.uid In ("900001","900002","900003"),"テストユーザ", sample.data1.uid In ("000001","000002","000003"),"管理者" ) AS ユーザ区分 INTO ユーザ一覧 FROM sample.data1;
```

### 出力結果
#### `[連番]_input.sql`
実行したSQL文

#### `[連番]_output.csv` もしくは `[連番]_[コメント].csv`
SQLで指定した出力結果のカンマ区切り形式（csv）
