from docx import Document
import pandas as pd
import re

from openpyxl import Workbook

pattern = re.compile('[0-9]+')
doc = Document('data/hr.docx')
paragraphs=doc.paragraphs
table_list = []
columns = []
df = pd.DataFrame()
count=0
for paragraph in paragraphs:
    text = paragraph.text
    project = "项目"
    domain ="主题域"

    #nuFlag = bool(re.search(r'\d',text))
    #flag0 = (project in text) & (domain in text) & (len(text)<=8)


    flag1 = text.endswith(domain) & (len(text)>6)

    if(flag1):
        #添加columns ,then  list.append  tablename str, then reset list  , end
        #
        columns.append(text)
        if(count>0):
            table_list.append(content)
        content = []
        count =+1
    dbname = "/DWD_HR_"
    flag = dbname in text
    if(flag):
        match = pattern.findall(text)
        if match:
            print("")
        else:

            content.append(text)



table_list.append(content)
series = pd.Series(table_list)
frame = pd.DataFrame(table_list[0:],index=columns)

hr_data = pd.read_csv("data/data_hr.csv")

for index, row in frame.iterrows():

    count=0
    for item in row:
        rowcontent = row[count]
        if(rowcontent is None):

            print("")
        else:
            tablecontent = row[count].split("/");

            #print(tablecontent[0])
            #print(tablecontent[1])
            hr_data.loc[(hr_data['表名称'] == tablecontent[1].lower()), '所属子领域'] = index

            hr_data.loc[(hr_data['表名称'] == tablecontent[1].lower()), '表名称'] = tablecontent[0]
            '''
            def fun(x):
                if x==tablecontent[1].lower():
                    return tablecontent[0]
                else:
                    return x
            hr_data['表名称'] = hr_data['表名称'].apply(lambda x: fun(x))
            '''
        count += 1

print(hr_data.head());

hr_data.to_csv('./a.csv', sep=',', header=True, index=False)