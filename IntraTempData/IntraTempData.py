import glob
import re
import os
import csv
import openpyxl
import pprint

# TODO: 引数からパラメータ取得
DIR_INPUT = "Input"
DIR_TEMP = "Template"
DIR_OUTPUT = "Output"
COL_PRIM = "PrimaryKey".upper()
PAT_REPLACE = '<<<%%(\w+)%%>>>'
SET_LANG = "JA"


class ErrMsg:
    msg = []

    def addMsg(self, msg):
        # エラーメッセージ登録
        self.msg.append(msg)


def initDirCheck():
    # 入力フォルダの有無確認
    if not os.path.exists(DIR_INPUT):
        print("No Input folder")
        exit(1)
    if not os.path.exists(DIR_TEMP):
        print("No Template folder")
        exit(1)


# 入力フォルダからファイル一覧を取得
def listInputFileName():
    retLstFile = []

    lst = glob.glob(DIR_INPUT + os.sep + "*.txt")
    retLstFile.extend(lst)
    lst = glob.glob(DIR_INPUT + os.sep + "*.csv")
    retLstFile.extend(lst)
    lst = glob.glob(DIR_INPUT + os.sep + "*.xlsx")
    retLstFile.extend(lst)

    return retLstFile


# 　テンプレートフォルダからファイル一覧を取得
def listTempFileName():
    retLstFile = []

    lst = glob.glob(DIR_TEMP + os.sep + "*.txt")
    retLstFile.extend(lst)
    lst = glob.glob(DIR_TEMP + os.sep + "*.csv")
    retLstFile.extend(lst)

    return retLstFile


# 入力ファイルからデータを取得、二次元配列([ファイル単位][行単位])
def GetInputData(arglstInFile, argPrimKey):
    # 戻り値の初期化
    retLstInRec = []

    for file in arglstInFile:
        if len(os.path.splitext(file)) < 1:
            # 拡張子がない場合
            continue

        # 拡張子の取得
        strExt = os.path.splitext(file)[1]
        # 複数行データの初期化
        lstLines = []

        if strExt == ".txt":
            # テキストファイルの読み込み
            lstLines = GetInputFromText(file)
        elif strExt == ".csv":
            # CSVファイルの読み込み
            lstLines = GetInputFromCsv(file)
        elif strExt == ".xlsx":
            # Excelファイルの読み込み(xls不可)
            lstLines = GetInputFromXlsxWithOpenpyxl(file, argPrimKey)

        if len(lstLines) > 0:
            retLstInRec.append(lstLines)

    # print(retLstInRec)

    return retLstInRec


def GetInputFromText(file):
    f = open(file, "r")
    lstLines = f.readlines()
    f.close()

    for idx, value in enumerate(lstLines):
        # 改行コードの削除
        lstLines[idx] = lstLines[idx].rstrip('\n')
        # SPLIT処理（TAB毎）
        lstLines[idx] = lstLines[idx].split('\t')

    return lstLines


def GetInputFromCsv(file):
    with open(file) as f:
        reader = csv.reader(f)
        lstLines = []
        # print(reader)
        for row in reader:
            lstLines.append(row)

        # print(lstLines)

    return lstLines


# Excelデータ取得
def GetInputFromXlsxWithOpenpyxl(strXls, strPrimKey):
    # xlsxファイルを開く
    wb = openpyxl.load_workbook(strXls)

    # シート名一覧
    sheets = wb.sheetnames

    # 戻りデータ
    retLst = []

    for sheet in sheets:
        ws = wb[sheet]

        # PrimaryKeyのY座標
        keyRow = -1
        # PrimaryKeyのX座標
        keyCol = -1
        # データ取得開始X座標
        stX = -1
        # データ取得終了X座標
        edX = -1

        # プライマリキーが存在するヘッダーの取得
        for row in ws.rows:

            for ix in range(1, len(row)):
                if row[ix].value is None:
                    # 空白の場合は次のセルへ進む
                    continue

                strValue = str(row[ix].value)
                strValue = strValue.upper()
                strValue = strValue.strip()

                if strValue == strPrimKey:
                    keyRow = row[ix].row
                    keyCol = ix
                    break

        # debug
        # print(keyRow, keyCol)

        # プライマリキー取得の有無確認
        if keyCol < 0:
            # プライマリキーがなければ次のシートに移動
            break

        # PrimaryKey行から開始X座標を取得
        for ix in range(1, len(row)):
            strValue = str(row[ix].value)
            strValue = strValue.upper()
            strValue = strValue.strip()

            # 開始X座標を取得
            if strValue != "":
                if stX < 0:
                    stX = ix

        # debug
        # print(keyRow, keyCol, stX)

        if stX < 0:
            # 開始X座標が取れなければ次のシートに移動
            break

        # PrimaryKey行から終了X座標を取得
        for ix in range(stX, len(row)):
            strValue = str(row[ix].value)
            strValue = strValue.upper()
            strValue = strValue.strip()

            # 終了X座標を取得
            if strValue != "":
                edX = ix
            else:
                # 終了X座標の終了
                break

        # debug
        # print(keyRow, keyCol, stX, edX)

        flgStart = False

        for row in ws.rows:
            if flgStart == False:
                if row[0].row < keyRow:
                    # データ取得開始前の行であれば次の行へすすむ
                    continue

                else:
                    flgStart = True

            lstRow = []
            flgBlank = True

            # 行データの取得、すべて空白であれば処理終了
            for ix in range(stX, edX):
                strValue = ""

                if row[ix].value is not None:
                    strValue = str(row[ix].value)

                if strValue != "":
                    flgBlank = False

                # 行データを戻りデータに登録
                lstRow.append(strValue)

            # 行データがすべてブランクであれば処理終了
            if flgBlank == True:
                break

            # 戻りデータに行データを登録
            retLst.append(lstRow)

    # debug
    # pprint.pprint(retLst)

    return retLst


def GetTempData(lstTempFiles):
    lstTemp = []

    for file in lstTempFiles:
        docTemp = []

        f = open(file, "r")
        docTemp = f.readlines()
        f.close()

        for idx, value in enumerate(docTemp):
            # 改行コード削除
            docTemp[idx] = docTemp[idx].rstrip('\n')

        lstTemp.append(docTemp)

    # print(lstTemp)

    return lstTemp


# 引数の文字列から禁止文字を削除して戻す
def removeNgCharsFromPrimKey(argKey):
    # 禁止文字
    strBad = '\/:*?"<>|\t'
    retKey = argKey

    # 大文字に変換
    retKey = retKey.upper()
    # 左右の空白を削除
    retKey = retKey.strip()

    for c in strBad:
        # 禁止文字を削除
        retKey = retKey.replace(c, "")

    return retKey


# プライマリキー配列の作成
def MakeInputRecords(arglstInLine, arglstInFile):
    # 戻り値の初期化
    retDicInput = {}
    # print("files=",lstInLines)

    # ファイル毎に処理
    for idxFile, idxValue in enumerate(arglstInLine):
        lstKey = []
        posPrim = -1

        # print(idxFile)
        # 行データ毎に処理
        for idxLine, vLine in enumerate(arglstInLine[idxFile]):
            if idxLine == 0:
                # 一行目のカラム行における処理
                for idx, col in enumerate(arglstInLine[idxFile][idxLine]):
                    # PrimaryKeyから不適切な文字を削除
                    col = removeNgCharsFromPrimKey(col)
                    lstKey.append(col)

                    if col == COL_PRIM:
                        posPrim = idx
                        # print("posPrim:",posPrim)

                if posPrim < 0:
                    print("Primary Key is not found in file '%s'" % arglstInFile[idxFile])
                    break

            else:
                # 2行目以降のデータ行における処理
                # 行辞書レコードの初期化
                dicRecord = {}
                # 行データの取得
                lstRec = arglstInLine[idxFile][idxLine]
                # PrimaryKeyの取得
                valPrim = lstRec[posPrim]

                for idx, value in enumerate(lstRec):
                    # 行辞書レコードに行データを登録
                    dicRecord[lstKey[idx]] = value

                # ファイル辞書レコードに行辞書レコードを追加
                retDicInput.setdefault(valPrim, []).append(dicRecord)

    # ファイル辞書レコードを戻す
    return retDicInput


# テンプレートフォルダパスから出力フォルダパスに変換
def ConvTempToOutFile(argTempFile, argKey):
    # フォルダ名の取得（テンプレート）
    strDir = os.path.dirname(argTempFile)

    # テンプレートフォルダから出力フォルダへ変換
    strDir = strDir.replace(DIR_TEMP, DIR_OUTPUT)

    # 最後のサブフォルダ名をテンプレートフォルダから出力フォルダ名に変更
    r = re.compile(r'%s$' % DIR_TEMP, re.IGNORECASE)
    strDir = re.sub(r, DIR_OUTPUT, strDir)

    # ファイル名の取得
    strFile = os.path.basename(argTempFile)

    # ファイル名にPrimaryKeyを付与
    lstFile = os.path.splitext(strFile)
    strFile = lstFile[0] + "_" + argKey + lstFile[1]

    # ファイルパスの再生成
    strRet = os.path.join(strDir, strFile)
    # print(strRet)
    return strRet


def makeOutDir():
    if os.path.exists(DIR_OUTPUT) == False:
        os.mkdir(DIR_OUTPUT)


def printFileList(strMsg, lstFile):
    if len(lstFile) <= 0:
        return

    for strFile in lstFile:
        if strFile is None or strFile == "":
            continue

        print(strMsg % strFile)


class FileInfo(object):
    name = ""
    lstLine = []


if __name__ == "__main__":
    # 入力フォルダがあることを確認
    initDirCheck()

    # 入力ファイルの収集
    lstInFiles = listInputFileName()

    if len(lstInFiles) <= 0:
        # 入力ファイルがない場合エラー終了
        print("No input files")
        exit(1)

    # 入力ファイルメッセージの出力
    printFileList("Inputファイル「%s」を読み込みました", lstInFiles)

    # テンプレートファイルの収集
    lstTempFileName = listTempFileName()
    # print(lstTempFiles)

    if len(lstTempFileName) <= 0:
        # テンプレートファイルがない場合エラー終了
        print("No template files")
        exit(1)

    # テンプレートファイルメッセージの出力
    printFileList("Templateファイル「%s」を読み込みました", lstTempFileName)

    # 入力ファイルの中身を一次配列に格納
    lstInAllLine = GetInputData(lstInFiles, COL_PRIM)
    # テンプレートファイルの中身を一次配列に格納
    lstTempAllLine = GetTempData(lstTempFileName)

    # 入力データ配列をPrimaryKeyをもとに辞書配列に変換
    dicInRecord = MakeInputRecords(lstInAllLine, lstInFiles)

    # print(dicInRecord)
    # for keyPrim, lstValue in enumerate(dicInRecord):
    for key in dicInRecord:
        # print("key=",key, dicInRecord[key])
        # print("keyPrim=",keyPrim,"lstValue",lstValue)
        for idx, value in enumerate(dicInRecord[key]):
            dicRecord = dicInRecord[key][idx]
            # print("key=",key,"idx=",idx,"dic=",dicRecord)

    # print(lstInLines)

    # 出力フォルダの作成
    makeOutDir()

    # print(lstTempAllLine)
    # print(lstTempFileName)

    lstOutFileName = []
    lstOutAllLine = []

    for idxFile, lstLines in enumerate(lstTempAllLine):

        # print(lstTempFiles[idxFile])
        # print(idxFile, lstLines)
        for key in dicInRecord:
            # 出力データ(1ファイル分)
            lstOutFileLine = []

            # 出力ファイル名の生成（テンプレート＋PrimaryKey）
            strFile = ConvTempToOutFile(lstTempFileName[idxFile], key)
            lstOutFileName.append(strFile)
            # lstOutFile.append(lstTempFiles[idxFile] + "_" + key)

            # print("file:", strFile,"mainKey",key)

            # テンプレートの行データ毎のループ処理
            # strTemp: 行データ
            for idxTemp, strTemp in enumerate(lstLines):
                # print(idxTemp, strTemp)

                # pattern = re.compile(r'<<<%%(\w+)%%>>>')
                pattern = re.compile(r'%s' % PAT_REPLACE)

                # 置換対象キーワードの検索
                ite = pattern.finditer(strTemp)
                # iterator の長さを取得する。
                iteCount = sum(1 for _ in re.finditer(pattern, strTemp))

                if ite is not None and iteCount > 0:
                    # 置換対象キーワードがある場合
                    lstOld = []
                    lstKey = []
                    strReplace = ""

                    for match in ite:
                        # 置換対象のキーワードを登録
                        lstOld.append(match.group(0).upper())
                        lstKey.append(match.group(1).upper())

                    for idx, value in enumerate(dicInRecord[key]):
                        strOutLine = strTemp
                        # print(idx, value)

                        # for match in ite:
                        for idxOld, valueOld in enumerate(lstOld):
                            # 置き換え対象文字列
                            oldWord = lstOld[idxOld]
                            # 置き換え対象文字列内のキーワード
                            keyWord = lstKey[idxOld]
                            value["TempData_Number".upper()] = str(idx + 1)
                            value["TempData_Index".upper()] = str(idx)

                            # 置き換える文字列
                            newWord = oldWord

                            if keyWord not in value:
                                # print("キーワード「%s」のデータが見つかりませんでした" % oldWord)
                                pass
                            else:
                                newWord = value[keyWord]

                            # print("key=",oldWord, newWord)

                            r = re.compile(r"%s" % oldWord, re.IGNORECASE)
                            strOutLine = re.sub(r, newWord, strOutLine)
                            # out = out.replace(oldWord, newWord)

                        # print("out2=",out)

                        if strReplace != "":
                            strReplace += "\n"

                        # 出力データ(1行)に置換した行を登録
                        strReplace += strOutLine

                    # 行データを出力データ(1ファイル)に登録
                    lstOutFileLine.append(strReplace)
                    # print("STR:", strReplace)

                else:
                    # 行データを出力データ(1ファイル)に登録
                    lstOutFileLine.append(strTemp)

                for line in lstOutFileLine:
                    # print(line)
                    pass

            # 出力データ(全ファイル)に登録
            if len(lstOutFileLine) > 0:
                # print("lstOutFileLine",lstOutFileLine)
                lstOutAllLine.append(lstOutFileLine)

    # print(lstOutFileName)
    # print(lstOutAllLine)
    # print(dicInRecord)

    for (idxFile, valFile) in enumerate(lstOutAllLine):
        strFile = lstOutFileName[idxFile]

        # ファイル名が空白であれば次の要素へ移動
        if strFile is None or strFile == "":
            continue

        outFile = open(strFile, "w")

        for idxLine, valLine in enumerate(lstOutAllLine[idxFile]):
            outFile.write(valLine + "\n")

        outFile.flush()
        outFile.close()

    # 出力ファイルメッセージの出力
    printFileList("Outputファイル「%s」を出力しました", lstOutFileName)

    # print("end")
