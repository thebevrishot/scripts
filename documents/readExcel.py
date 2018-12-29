import xlrd

# usage:
# path:str path of excel file
# sheetIdx: integer index of book
# return dict key topicname => value set of document within the topic
def extract(path,sheetIdx):
  book = xlrd.open_workbook(path)
  sh = book.sheet_by_index(sheetIdx)
  topicIdx = _getTopicNameCol(sh)
  res = {}
  for k in topicIdx:
    res[k] = _getAllInSet(sh,topicIdx[k])
  print(res)

# get all members of the topic by col inx
# ignore document that have length no more than 1
def _getAllInSet(sh,idx):
  cnt = 1
  res = set()
  for i in range(1,sh.nrows):
    cont = sh.cell_value(rowx=i,colx=0)
    val = sh.cell_value(rowx=i,colx=idx)
    print(len(cont))

    if len(cont) > 1:
      if len(val) >= 1:
        res.add(cnt)
      cnt += 1

  return res

# get dict of topicname => col index
def _getTopicNameCol(sh):
  res = {}
  for i in range (1,sh.ncols):
    col = sh.col(i)[0].value
    res[col] = i
  return res

if __name__ == '__main__':
  extract('./Gold_Standard_Dataset_Annotated_Example.xlsx',0)
