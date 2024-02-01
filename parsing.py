# * 开源小脚本 顺手点了 strar
inputFileName = 'xxx.xlsx' # 需要输出的文件名称
outputFileName = 'xxxa.xlsx' # 需要输出的文件名称
colnamekey = 'xx' # 作为需要分出的sheet的 heard 列名
fliterkey = 'xx' # 作为需要过滤的数据的 heard 列名 
# islegal 根据 fliterkey 校验合法数据 需要自定义表达式 sing原数据行 sing[fliterindex]原数据一行中的 fliterkey 这一列 有多个条件 需要用多少列去原数据 从0开始 例:sing[0]
def islegal(sing,fliterindex):
   if sing[fliterindex] > 6:
      return True
   else:
      return False
# 以上是需要修改的变量 和 过滤的表达式 
# **************************************   
# 以下代码是实现逻辑非配置项
import pandas as pd
excel = pd.ExcelFile(inputFileName)
Data = pd.read_excel(excel)
# print(Data,'原文件数据浏览')
# Data.head()
adds = Data.values.tolist()
columnsList = Data.columns.tolist()
colsindex = columnsList.index(colnamekey)
def isNull(fliterkey):
   if fliterkey:
      return columnsList.index(fliterkey)
   else:
      return 0
colsfliterindex = isNull(fliterkey)
waritColumnsList = pd.DataFrame(columns=columnsList)
waritColumnsList.to_excel(outputFileName, index=False)
def fliterObj(seq,sing):
   if fliterkey:
      if seq[colsindex] == sing and islegal(seq,colsfliterindex):
        return seq
   else:
      if seq[colsindex] == sing:
        return seq
def f(x):
   return x[colsindex]
nameList = list(map(f,adds))
new_list = list(set(nameList))
with pd.ExcelWriter(outputFileName) as writer:
    for sing in new_list:
      waritdataList = pd.DataFrame(list(filter(lambda seq:fliterObj(seq,sing),adds)))
      waritdataList.to_excel(writer,sheet_name=sing,index=False,header=columnsList,startrow=0)
print('success 成功分表')


