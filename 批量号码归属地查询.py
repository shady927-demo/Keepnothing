# coding:utf-8
from requests import get
from demjson import decode
import pandas as pd
from xlwt import Workbook
from datetime import datetime
from sys import exit

file_path = input("请复制目标excel的绝对路径：\n")
file_name = input("请输入文件名称(必须为xlxs文件)：\n")
df = pd.read_excel(file_path+'/'+file_name+'.xlsx',usecols=[3])
url = "http://tcc.taobao.com/cc/json/mobile_tel_segment.htm?"
workbook = Workbook(encoding = 'utf-8')
worksheet = workbook.add_sheet('i')
print(df.head(5))
choice = input('请确认是不是手机号的数据（Y/N）:\n')
if choice == 'y' or choice =='Y':
    pass
else:
    exit()
df = df.values.tolist()
for i in range(len(df)):
    df[i] = str(df[i])[1:-1]

def result(phone_nums):
    for i in range(len(phone_nums)):
        d = {"tel":phone_nums[i]}
        rep = get(url,d)
        show = rep.content[19:-1].decode('gb2312')
        nu = "\n\t"+" "
        for j in nu:
            show = show.replace(j,'')
        show = decode(show)
        worksheet.write(i, 0, show['carrier'])
        if i%50 == 0:
            print("剩余"+str(len(phone_nums)-i)+"条")
print("数据量较大，请您耐心等待。。。")
result(df)
print("处理完成！请在D盘查看,文件名为：数据结果"+str(datetime.now().second))
workbook.save('D:/'+'数据结果'+str(datetime.now().second)+'.xlsx')
input("程序结束，任意键退出：")
