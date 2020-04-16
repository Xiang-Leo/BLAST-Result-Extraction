'''
在blastn中查找相关的信息并整理到excel中
writen by 李祥
'''
import re
import xlwt
path = 'D:\\works\\20191003AvianRetroVirus\\Result20191008\\resulted20191024nonretrovirusid80.txt'       #读取文件路径

with open(path) as raw_data:               #将文件每一行的内容读取到列表中
    everylines = raw_data.readlines()

'''
def xlsxwt():   #怼excel表格进行写操作
                  没想好函数的参数传入怎么写
                  query.group(0) != None:
'''

data = xlwt.Workbook()
table=data.add_sheet('info')


n = 1     #用于在excel表格中计数
for line in everylines:                    #使用正则表达式查找需要处理的内容
                                 #在表格中确定单元格的位置
    query = re.search(r'Query=\s\w+\S\w+\S+',line)
    if re.search(r'Query=\s\w+\S\w+\S+',line):               #写入query
        table.write(n-1,1,query.group(0)[7:])
        n = n + 1 
    align = re.search(r'>\s\w+',line)          #写入查找到的物种基因组的名字
    if re.search(r'>\s\w+',line):
        table.write(n,2,align.group(0)[2:])
      #  n = n + 1 
    score = re.search(r'Score = \d+',line)
    if re.search(r'Score = \d+',line):
        table.write(n,3,score.group(0)[8:])
     #   n = n + 1 
    expect = re.search(r'Expect\s=\s\w+-\d+',line)
    if re.search(r'Expect\s=\s\w+-\d+',line):
        table.write(n,4,expect.group(0)[8:])
     #   n =n + 1
    elif re.search(r'Expect\(\d\)\s=\s\w+-\d+',line):
        expect = re.search(r'Expect\(\d\)\s=\s\w+-\d+',line)
        table.write(n,4,expect.group(0)[12:])
    ide = re.search(r'Identities\s=\s\w+/\d+\s\(\d+%\)',line)
    if re.search(r'Identities\s=\s\w+/\d+\s\(\d+%\)',line):
        table.write(n,5,ide.group(0)[13:])
        pos = re.search(r'Positives\s=\s\w+/\d+\s\(\d+%\)',line)
        table.write(n,6,pos.group(0)[12:])
        gaps = re.search(r'Gaps\s=\s\d+/\d+\s\(\d+%\)',line)
        table.write(n,7,gaps.group(0)[7:])
        n = n + 1
    
data.save('D:\\works\\20191003AvianRetroVirus\\Result20191008\\info.xls')