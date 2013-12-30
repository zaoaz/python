#encoding:utf8				
#test
import	xlrd
import sys
reload(sys)
sys.setdefaultencoding( "utf-8" )

def LoadExcel(path,paramPath):	
        params = {}
        paramf = open(paramPath)
        sParams = paramf.read().split("|")
        print sParams
        for sparam in sParams:
            sparam_ = sparam.split(",")
            if len(sparam_) == 1:
                break
            params["sheet"+sparam_[0]]={
                "sheetIndex":int(sparam_[0]),
                "beginRow":int(sparam_[1]),
                "colname":sparam_[2],
                }
        rfile = xlrd.open_workbook(path)
        sqlFile = open("result.sql","w");
        for param in params.values():
            sheetIndex = param["sheetIndex"];
            table = rfile.sheet_by_index(sheetIndex)
            beginRow = param["beginRow"];
            colNum = colname_to_num(param["colname"])
            tableName = str(table.cell(beginRow-1,colNum).value).split(" ")[2].split("(")[0];
            sqlFile.write("TRUNCATE TABLE "+tableName+";\n");
            for rowIndex in range(beginRow-1,table.nrows):
                sqlFile.write(str(table.cell(rowIndex,colNum).value)+"\n");
            

def colname_to_num(colname):
    if type(colname) is not str:
        return colname
    col = 0
    power  = 1
    for i in xrange(len(colname) - 1, -1, -1):
        ch = colname[i]
        col += (ord(ch) - ord('A') +  1 ) * power
        power *= 26
    return col - 1
LoadExcel('sql.xls','param.txt')
