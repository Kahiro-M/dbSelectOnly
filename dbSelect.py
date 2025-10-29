import re
import pyodbc
import sys
import os
from mkdir_datetime import mkdir_datetime
dir_path = './'


# iniファイルから設定を読み込む
def readConfigIni(filePath='config.ini'):
    import configparser

    #ConfigParserオブジェクトを生成
    config = configparser.ConfigParser()

    #設定ファイル読み込み
    config.read(filePath,encoding='utf8')

    configData = {}
    configString = ''

    #設定情報取得
    if(config.has_option('DB','DSN')):
        configData['DSN'] = config.get('DB','DSN')
        configString += repr('DSN='+configData['DSN'])[1:-1]+r';'
    else:
        configData['DSN'] = ''

    if(config.has_option('DB','UID')):
        configData['UID'] = config.get('DB','UID')
        configString += repr('UID='+configData['UID'])[1:-1]+r';'
    else:
        configData['UID'] = ''

    if(config.has_option('DB','PWD')):
        configData['PWD'] = config.get('DB','PWD')
        configString += repr('PWD='+configData['PWD'])[1:-1]+r';'
    else:
        configData['PWD'] = ''

    if(config.has_option('DB','DATABASE')):
        configData['DATABASE'] = config.get('DB','DATABASE')
        configString += repr('DATABASE='+configData['DATABASE'])[1:-1]+r';'
    else:
        configData['DATABASE'] = ''

    if(config.has_option('DB','DB')):
        configData['DB'] = config.get('DB','DB')
        configString += repr('DB='+configData['DB'])[1:-1]+r';'
    else:
        configData['DB'] = ''

    if(config.has_option('DB','DFLT_BIGINT_BIND_STR')):
        configData['DFLT_BIGINT_BIND_STR'] = config.get('DB','DFLT_BIGINT_BIND_STR')
        configString += repr('DFLT_BIGINT_BIND_STR='+configData['DFLT_BIGINT_BIND_STR'])[1:-1]+r';'
    else:
        configData['DFLT_BIGINT_BIND_STR'] = ''

    if(config.has_option('DB','DRIVER')):
        configData['DRIVER'] = config.get('DB','DRIVER')
        configString += repr('DRIVER='+configData['DRIVER'])[1:-1]+r';'
    else:
        configData['DRIVER'] = ''

    if(config.has_option('DB','NO_SCHEMA')):
        configData['NO_SCHEMA'] = config.get('DB','NO_SCHEMA')
        configString += repr('NO_SCHEMA='+configData['NO_SCHEMA'])[1:-1]+r';'
    else:
        configData['NO_SCHEMA'] = ''

    if(config.has_option('DB','PASSWORD')):
        configData['PASSWORD'] = config.get('DB','PASSWORD')
        configString += repr('PASSWORD='+configData['PASSWORD'])[1:-1]+r';'
    else:
        configData['PASSWORD'] = ''

    if(config.has_option('DB','PORT')):
        configData['PORT'] = config.get('DB','PORT')
        configString += repr('PORT='+configData['PORT'])[1:-1]+r';'
    else:
        configData['PORT'] = ''

    if(config.has_option('DB','PWD')):
        configData['PWD'] = config.get('DB','PWD')
        configString += repr('PWD='+configData['PWD'])[1:-1]+r';'
    else:
        configData['PWD'] = ''


    if(config.has_option('DB','SERVER')):
        configData['SERVER'] = config.get('DB','SERVER')
        configString += repr('SERVER='+configData['SERVER'])[1:-1]+r';'
    else:
        configData['SERVER'] = ''


    if(config.has_option('DB','UID')):
        configData['UID'] = config.get('DB','UID')
        configString += repr('UID='+configData['UID'])[1:-1]+r';'
    else:
        configData['UID'] = ''


    if(config.has_option('DB','USER')):
        configData['USER'] = config.get('DB','USER')
        configString += repr('USER='+configData['USER'])[1:-1]+r';'
    else:
        configData['USER'] = ''

    configData['CONNECTION_STRING'] = configString

    return configData


def checkSelectOnlySql(sql):
    import sqlparse

    # SQLの解析
    parsedSql = sqlparse.parse(sql)
    isSelectOnly = True
    hasDmlList = []
    ret = {}
    for sqlStr in parsedSql:
        # SELECT文以外の存在をチェック
        if(sqlStr.get_type() != 'SELECT'):
            isSelectOnly = False
        hasDmlList.append(sqlStr.get_type())

    # 結果を整形
    ret['IS_SELECT_ONLY'] = isSelectOnly
    ret['HAS_DML_LIST'] = hasDmlList
    ret['SQL'] = sqlparse.split(sql)
    return ret

# SQL文とコメントを分割する [{'sql': ..., 'comment': ...}, ...]
def splitSqlAndComment(sql_text: str):
    results = []

    # 改行統一
    sql_text = sql_text.replace('\r\n', '\n')

    # コメントを含めた行単位に分解
    lines = sql_text.split('\n')

    current_sql = []
    current_comment = ""

    for line in lines:
        stripped = line.strip()
        if not stripped:
            continue

        # コメント行（-- で始まる）
        if stripped.startswith('--'):
            # 現在のSQL文がまだ途中なら、コメントは「そのSQLの説明」として扱う
            if current_sql:
                # コメントは連結（複数行コメント対応）
                if current_comment:
                    current_comment += " "
                current_comment += stripped[2:].strip()
            else:
                # SQLが始まっていない状態のコメントは無視（または次のSQLに付与）
                current_comment = stripped[2:].strip()
            continue

        # SQL本体行
        current_sql.append(line)

        # 文末にセミコロンがある場合 → SQL文として確定
        if ';' in line:
            sql_stmt = '\n'.join(current_sql).strip()
            # セミコロンの後のコメントを抽出
            m = re.search(r';\s*--\s*(.*)$', sql_stmt)
            if m:
                inline_comment = m.group(1).strip()
                sql_stmt = re.sub(r';\s*--.*$', ';', sql_stmt).strip()
                if current_comment:
                    current_comment += " " + inline_comment
                else:
                    current_comment = inline_comment

            results.append({
                'sql': sql_stmt,
                'comment': current_comment.strip()
            })

            # 次のSQLに備えてリセット
            current_sql = []
            current_comment = ""

    return results


def execSql(config,sql):
    # ODBCで接続
    connection = pyodbc.connect(config['CONNECTION_STRING'])
    cursor = connection.cursor()

    # SQLの実行
    try:
        ret = cursor.execute(sql)
    except pyodbc.Error as e:
        err_str = str(e).encode("utf-8", errors="replace").decode("utf-8")
        print("Safe error string:", err_str)
        sys.exit()
    # 実行結果の整形 ↓の形式
    # {
    #     'rowInfo':
    #         [
    #             {column1:value1-1,column2:value1-2},
    #             {column1:value2-1,column2:value2-2}
    #         ],
    #     'comment':
    #         "comment1,comment2"
    # }
    rows = cursor.fetchall()
    rowInfo = []
    i = 0

    for row in rows:
        rowInfo.insert(i,{})
        j = 0
        for column in row:
            rowInfo[i][row.cursor_description[j][0]] = column
            j += 1
        i += 1

    # ODBCを切断
    connection.close()

    return rowInfo


# 辞書型データをCSVとして保存
def dict2csv(data,filePath='.\output.csv'):
    import csv

    # CSVファイルに書き込む
    with open(filePath,'w',newline='',encoding='utf-8_sig') as file:
        # フィールド名（CSVのヘッダー）を辞書のキーから取得
        fieldnames = data[0].keys()

        # csv.DictWriterを使って辞書を書き込む準備
        writer = csv.DictWriter(file, fieldnames=fieldnames)

        # ヘッダーを書き込む
        writer.writeheader()

        # 各辞書の行を書き込む
        writer.writerows(data)


# 辞書型データをTXTとして保存
def dict2txt(data,filePath='.\output.txt'):
    # pprintで整形した文字列を取得
    outputText = pprint.pformat(data)

    # 整形した文字列をファイルに書き込む
    with open(filePath,'w',encoding='utf-8') as file:
        file.write(outputText)



# 指定文字列をTXTとして保存
def str2txt(data,filePath='.\output.txt'):
    # 整形した文字列をファイルに書き込む
    with open(filePath,'w',encoding='utf-8') as file:
        file.write(data)



if __name__ == '__main__':
    import tkinter, tkinter.filedialog, tkinter.messagebox, tkinter.ttk
    import pprint
    def execSqlFile(file):
        # SQLファイルを読み込み
        with open(file, 'r', encoding='UTF-8') as f:
            sqlText = f.read()

        # SQL文を文単位で分割
        sqlStatements = splitSqlAndComment(sqlText)


        # 出力先ディレクトリ作成
        outPath = mkdir_datetime('SQL_RESULT_')

        # 各SQL文を実行
        for i, sqlData in enumerate(sqlStatements):

            # SELECT文以外を含むかチェック
            parsedSql = checkSelectOnlySql(sqlData['sql'])
            if not parsedSql['IS_SELECT_ONLY']:
                tkinter.messagebox.showinfo(
                    'SQLエラー',
                    '指定したSQLにSELECT文以外のステートメントが含まれています。\n'
                    'SELECT文のみ利用可能です。\n含まれているSQL：[' + ','.join(parsedSql['HAS_DML_LIST']) + ']'
                )
                return -1

            print(f'{i+1}件目 SQLを実行中...')
            execRet = execSql(config, sqlData['sql'])
            execDict = {'rowInfo':execRet,'comment':sqlData['comment']}

            # 入力SQLと結果をファイル出力
            str2txt(sqlData['sql'], filePath=os.path.join(outPath, f'{i}_input.sql'))
            csv_comment = execDict["comment"] if execDict["comment"] else f'sql_{i}'
            dict2csv(execDict['rowInfo'], filePath=os.path.join(outPath, f'{i}_{csv_comment}.csv'))

            print(f'{i+1}件目 実行完了')

        return len(sqlStatements)
    # 設定ファイルから接続情報を取得
    config = readConfigIni('config.ini')
    if(len(config['CONNECTION_STRING']) < 1):
        tkinter.messagebox.showerror('設定エラー','INIファイルから設定情報が取得できませんでした。\n終了します。')
        print('ERR : INIファイルから設定情報が取得できませんでした。終了します。')
        sys.exit()

    # SQLの指定（引数にない場合はファイルを選択）
    if(len(sys.argv)<2):
        execFlg = True
        while(execFlg):
            print('============ DBから情報取得 ============')
            print('                                 v.1.2.0')
            print('・複数行のSQLに対応')
            print('------------- 実行候補 SQL -------------')
            
            # 現在のディレクトリを取得
            currentDir = os.getcwd()
            # 現在のディレクトリ内のCSVファイルの一覧を取得
            sql_files = [f for f in os.listdir(currentDir) if f.endswith(('.sql','.SQL'))]
            sqlFileList = {}
            for i,sql_file in enumerate(sql_files):
                sqlFileList[i+1] = sql_file
                print('       '+str(i+1)+' : '+sql_file)
            print('       0 : その他のSQLを選択')
            print('    空白 : 終了')
            print('----------------------------------------')
            choiceStr = input('入力してください：')

            # exitの場合は終了
            if(len(choiceStr)<1):
                print('終了します。\n\n')
                execFlg = False

            # 0の場合はファイル選択
            elif(choiceStr=='0'):
                root = tkinter.Tk()
                root.attributes('-topmost', True)
                root.withdraw()
                tkinter.messagebox.showinfo('SQL指定','sqlファイルを選択してください')
                # ファイル選択ダイアログの表示
                fTyp = [('SQLファイル','*.sql')]
                iDir = os.path.abspath(dir_path)
                file = tkinter.filedialog.askopenfilename(filetypes=fTyp,initialdir = iDir)
                fileList = [file]
                ret = execSqlFile(fileList[0])
                print(str(ret)+'件 実行完了\n\n')

            # 選択肢にある場合は実行
            elif(int(choiceStr) in list(sqlFileList.keys())):
                ret = execSqlFile(sqlFileList[int(choiceStr)])
                print(str(ret)+'件 実行完了\n\n')

            # 選択肢に無い場合はもう一度訊く
            else:
                print('選択肢にありません。もう一度、選択してください。\n\n')

    else:
        ret = execSqlFile(sys.argv[1])
        print(str(ret)+'件 実行完了\n\n')
        print('終了します。\n\n')

        # sql = ''.join(sys.argv[1:])

