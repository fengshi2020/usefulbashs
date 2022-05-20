# -*- coding: utf-8 -*-

import time
import pandas as pd
import os
from pandas.core.frame import DataFrame

from gooey import GooeyParser, Gooey
form PIL import Image


@Gooey(encoding='utf-8', program_name='云资源使用率分析', program_description='分析云资源使用率', optional_cols=1)
def analyseExcel():
	parse = GooeyParser(description='分析云资源使用率')
	parse.add_argument('file',help='选择需要分析的excel文件（.xlsx格式）', widget='FileChooser', default=os.getcwd())
	parse.add_argument('save',help='保存分析处理后的文件（.xlsx格式）', widget='DirChooser', default=os.getcwd())
	parse.add_argument('dep',help='请输入团队列列名（对照文件中列名正确输入）')
	parse.add_argument('mem',help='请输入内存列名称（对照文件中列名正确输入，不能为null）')
	parse.add_argument('jud',help='请输判断列（对照文件中列名正确输入，需要是比列，数字，不能为null）')
	args = parse.parse_args()

	print('________________________________\n' + time.strftime('%Y-%M-%d %H:%M:%S')+'\n')
	filePath = args.file
	saveFile = args.save+'\\' + os.path.basename(filePath).split('.')[0]+'-processed.xlsx'
	dep = args.dep
	mem = args.mem
	jud = args.jud
	print('分析的文件：'+filePath+'\n')
	print('处理后的文件：'+saveFile+'\n')

	# 设置全局变量
	listDepCon = []  # 团队数量列表
	listCount = []  # 各个数据变量列表

	#创建新表
	newexcel = pd.ExcelWriter(saveFile)

	#打开并读取excel文件
	df_excelAll = pd.read_excel(filePath,converters={jud:int})

	#统一内存单位为GB
	if df_excelAll[mem].min() >= 1024 :
		df_excelAll[mem] = df_excelAll[mem]/1024

	#获取部门列表
	listDep = df_excelAll[dep].unique()

	#获取各部门各种数量
	for i in range(len(listDep)):
		listDepCon.append(df_excelAll[dep].value_counts()[listDep[i]])
		getNum(df_excelAll,listDep[i],newexcel,dep,mem,jud,listCount)

	#生成部门及其主机数量
	listRes = {'部门':listDep.tolist(),'总数':listDepCon}
	df = DataFrame(listRes)

	#生成各比率数量
	dfCount = DataFrame(listCount)
	dfCount.columns = ['≥50','30-50','≤30','Null']

	#合并部门主机数量和各个使用率数量
	dfAll = pd.concat([df,dfCount],axis=1)

	#计算各个使用率的比率
	dfAll['≥50P']=(dfAll['≥50']/dfAll['总数']).apply(lambda x: format(x, '.1%'))
	dfAll['30-50P']=(dfAll['30-50']/dfAll['总数']).apply(lambda x: format(x, '.1%'))
	dfAll['≤30P']=(dfAll['≤30']/dfAll['总数']).apply(lambda x: format(x, '.1%'))

	#导出结果到excel
	dfAll.to_excel(newexcel,index=False,sheet_name='统计')
	#保存结果
	newexcel.save()
	newexcel.close()
	print('处理完成！\n_______________________________________\n\n')

#计算各部门各种使用率数量
def getNum(df1,str,df2,dep,mem,jud,list):
	df_excel = df1[df1[dep] == str]
	df_excel.to_excel(df2,index=False,sheet_name=str) #按照部门生成表
	df_temp = df_excel[df_excel[mem] >2]
	temp = df_temp[df_temp[jud] >= 50]
	if str in temp[dep].values:
		n = temp[dep].value_counts()[str]
	else:
		n = 0
	temp = df_temp[(df_temp[jud] < 50) & (df_temp[jud] > 30)]
	if str in temp[dep].values:
		m = temp[dep].value_counts()[str]
	else:
		m = 0
	temp = df_temp[df_temp[jud] <= 30]
	if str in temp[dep].values:
		p = temp[dep].value_counts()[str]
	else:
		p = 0
	temp = df_temp[df_temp[jud].isnull()]
	if str in temp[dep].values:
		k = temp[dep].value_counts()[str]
	else:
		k = 0
	listT=[n,m,p,k]
	list.append(listT)

def showExamle():
	img = Image.open('example.jpg')
	img.show()

if __name__ == '__main__':
	#showExamle()
	analyseExcel()