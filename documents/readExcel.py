import xlrd

# transpose of extractTopicNameDocSet
# usage:
# path:str path of excel file
# sheetIdx: integer index of book
# return dict key doc id => value set of topics
def extract(path,sheetIdx):
  topicDoc = extractTopicNameDocSet(path,sheetIdx)
  res = {}
  for k in topicDoc:
    docs = topicDoc[k]
    for doc in docs:
      if doc not in res:
        res[doc] = set()
      res[doc].add(k)
  return res

# usage:
# path:str path of excel file
# sheetIdx: integer index of book
# return dict key topicname => value set of document within the topic
def extractTopicNameDocSet(path,sheetIdx):
  book = xlrd.open_workbook(path)
  sh = book.sheet_by_index(sheetIdx)
  topicIdx = _getTopicNameCol(sh)
  res = {}
  for k in topicIdx:
    res[k] = _getAllInSet(sh,topicIdx[k])
  return res

# get all members of the topic by col inx
# ignore document that have length no more than 1
def _getAllInSet(sh,idx):
  cnt = 1
  res = set()
  for i in range(1,sh.nrows):
    cont = sh.cell_value(rowx=i,colx=0)
    val = sh.cell_value(rowx=i,colx=idx)

    if len(cont) > 1:
      if len(val) >= 1:
        res.add(cnt)
      cnt += 1

  return res

def _getDocTopics(sh,rowIdx):
  for c in sh.row(rowIdx)[1:]:
    print(c)

# get dict of topicname => col index
def _getTopicNameCol(sh):
  res = {}
  for i in range (1,sh.ncols):
    col = sh.col(i)[0].value
    res[col] = i
  return res


# get dict of topicname => col index
def _getColIdxTopicName(sh):
  res = {}
  for i in range (1,sh.ncols):
    col = sh.col(i)[0].value
    res[i] = col
  return res

if __name__ == '__main__':
  res = extract('./Gold_Standard_Dataset_Annotated_Example.xlsx',0)
  print(res)
