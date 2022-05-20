from faker import Faker
import pandas as pd
import numpy as np

#df = pd.DataFrame(columns=['姓名','手机号','地址','身份证号'])
n=int(input('请输入需要生成数据的数量：'))
fake = Faker('zh_CN')
lists = []
for i in range(1,n+1):
    list = [i,fake.name(),fake.phone_number(),fake.address(),fake.ssn(),fake.free_email()]
    lists.append(list)
    #print(list)
userdf = pd.DataFrame(lists,columns=['序号','姓名','手机号','地址','身份证号','邮箱'])
#print(lists)
#print(userdf)
userdf.to_excel('userdf.xlsx',sheet_name=str(n),index=False)
