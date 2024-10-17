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

def execSql(config,sql):
    # ODBCで接続
    connection = pyodbc.connect(config['CONNECTION_STRING'])
    cursor = connection.cursor()

    # SQLの実行
    ret = cursor.execute(sql)

    # 実行結果の整形 ↓の形式
    # [
    #   {column1:value1-1,column2:value1-2},
    #   {column1:value2-1,column2:value2-2},
    # ]
    rows = cursor.fetchall()
    rowInfo = []
    i = 0

    for row in rows:
        print(i,row)
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
    import tkinter, tkinter.filedialog, tkinter.messagebox
    import pprint

    # 設定ファイルから接続情報を取得
    config = readConfigIni('config.ini')
    if(len(config['CONNECTION_STRING']) < 1):
        tkinter.messagebox.showerror('設定エラー','INIファイルから設定情報が取得できませんでした。\n終了します。')
        print('ERR : INIファイルから設定情報が取得できませんでした。終了します。')
        sys.exit()

    # SQLの指定（引数にない場合はファイルを選択）
    if(len(sys.argv)<2):
        tkinter.messagebox.showinfo('SQL指定','sqlファイルを選択してください')
        # ファイル選択ダイアログの表示
        root = tkinter.Tk()
        root.withdraw()
        fTyp = [('SQLファイル','*.sql')]
        iDir = os.path.abspath(dir_path)
        file = tkinter.filedialog.askopenfilename(filetypes=fTyp,initialdir = iDir)
        fileList = [file]
        with open(fileList[0],'r',encoding='UTF-8') as f:
            sql = ''.join(f.read().splitlines())
    else:
        sql = ''.join(sys.argv[1:])
    
    # SQLのチェック（SELECT文以外が含まれているか）
    parsedSql = checkSelectOnlySql(sql)
    if(parsedSql['IS_SELECT_ONLY'] == False):
        tkinter.messagebox.showinfo('SQLエラー','指定したSQLにSELECT文以外のステートメントが含まれています。\nSELECT文のみ利用可能です。\n含まれているSQL：['+','.join(parsedSql['HAS_DML_LIST'])+']')
        sys.exit()

    # SQLの実行
    i = 1
    outPath = mkdir_datetime('SQL_RESULT_')
    for sqlStr in parsedSql['SQL']:
        execDict = execSql(config,sqlStr)
        str2txt(sqlStr,filePath=outPath+'\input_'+str(i)+'.sql')
        dict2txt(execDict,filePath=outPath+'\output_'+str(i)+'.txt')
        dict2csv(execDict,filePath=outPath+'\output_'+str(i)+'.csv')
        i += 1
