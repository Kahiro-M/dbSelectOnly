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
2. `***.sql`に実行したいSQLを入力（1行）
3. `dbSelect.py`を実行し、`***.sql`を指定
3. `SQL_RESULT_yyyymmdd_hhmm_ss`に以下が出力される
    - input_[連番].sql
    - output_[連番].csv
    - output_[連番].txt

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
SQLは複数行実行できるが、1つのSQLは1行に収める必要がある。
#### OK
```
SELECT * FROM `sample`.`data1`;
SELECT * FROM `sample`.`data2`;
```
#### NG
```
-- 1つのSQLが複数行に渡っている。
SELECT
 * 
FROM
  `sample`.`data1`;
```
```
-- Accessのみのクエリを含むSQLをMySQLで実行している(Switch関数など)
SELECT sample.data1.uid, Switch( sample.data1.uid In ("900001","900002","900003"),"テストユーザ", sample.data1.uid In ("000001","000002","000003"),"管理者" ) AS ユーザ区分 INTO ユーザ一覧 FROM sample.data1;
```

#### input_[連番].sql
実行したSQL文

#### output_[連番].csv
SQLで指定した出力結果のカンマ区切り形式（csv）

#### output_[連番].txt
SQLで指定した出力結果のJSON形式
