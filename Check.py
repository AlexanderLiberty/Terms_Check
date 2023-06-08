#!/usr/bin/env python
#-- coding: utf-8 --
import pandas as pd
import time
import os

'''
1.读取excel ok
2.获取结果记录和符合项列 ok
3.判断sheet和结果记录是否一致，过滤不适用 ok
4.判断符合项和结果记录条件语句是否一致 ok
5.输出有问题的内容 ok
'''
print("/************************************************\n密评/等保，结果记录标准型检测。\nBy：AlexanderLiberty\nDate：2023-01-10\n************************************************/\n\n\n")
path = input("输入“待查项路径,支持xlsx文件”开始检查:)，：")

bad = ["不满足","未定期","未限制","无法","未采用","未对","未禁止","未设置","不合理","未实现","未关闭","未重命名","不适用"]
#bad[0]
okey=["配置合理","已开启","需输入","不存在","可保证","可检测","可实现","可通过","可防止","可避免","满足","可随","可以","可对","不适用"]

#sheets，sheet列表；sheet，单个sheet；
#pand，判定结论，主键；jieguo，判定条件，结论，核查项。

#符合项和部分符合项中出现不符合的条件则执行这一条
def ookey(pand_str,jieguo_str,j):
	for i in bad:
		ookey_cz = jieguo_str.find(i)
		if ookey_cz != -1:
			print("*************************************************************************************\n符合&部分符合情况\n复查：《不满足，未定期，未限制，无法，未采用，未对，未禁止，未设置，不合理，未实现，未关闭，未重命名》判断条件，与判定结果不符\n*************************************************************************************\n")
			print("“" + j + "”"+"结果记录疑似存在问题！")
			print("符合情况：" + pand_str + "\n" + jieguo_str)
			#print(jieguo_str)
			print("\n\n\n")

#不符合中不合理项
def bbad(pand_str,jieguo_str,j):
	for nobd in okey:
		nobd_cz = jieguo_str.find(nobd)
		if nobd_cz != -1:
			print("**********************************************************************************\n不符合情况\n复查：《配置合理，已开启，需输入，不存在，可保证，可检测，可实现，可通过，可防止，可避免，满足，可随，可以，可对》判断条件，与判定结果不符\n**********************************************************************************\n")
			print("“" + j + "”"+"结果记录存在问题！")
			print("符合情况：" + pand_str + "\n" + jieguo_str)
			print("\n\n\n")

#bad sheet
def sheet_bad(pand,jieguo,j):
	count_name = 0
	while(count_name < len(jieguo)):
		pand_str_tmp = pand[count_name]#键
		jieguo_str_tmp = jieguo[count_name]#值
		sheet_bad_cz = jieguo_str_tmp.find(str(j))#在结果中查询标题
		if sheet_bad_cz == -1 and pand_str_tmp != "不适用":
			print("**********************************************************************************\n标题不符合情况\n复查：主语与结果记录不同\n**********************************************************************************\n")
			print("“" + j + "”"+"结果记录《主语》存在问题！")
			print("结果记录主语：\n" + jieguo_str_tmp + "\n\n")
		count_name = count_name + 1
	print("\n\n\n")

#select sheet name for excel
def select(sheet,path,j):
	dataframe = pd.read_excel(path, sheet_name = sheet)
	pand = dataframe['符合情况'].values.tolist()
	jieguo = dataframe['结果记录'].values.tolist()
	count = 0
	#包含性检测
	sheet_bad(pand,jieguo,j)
	while(count < len(jieguo)):
		pand_str = pand[count]#键
		jieguo_str = jieguo[count]#值
		if pand_str == "符合" or pand_str == "部分符合":
			ookey(pand_str, jieguo_str,j)
		elif pand_str == "不符合":
			bbad(pand_str,jieguo_str,j)
		#开关语句
		input("输入“没问题”检查下一个结果记录：")
		count = count + 1

#读取excel，sheet
sheets=pd.read_excel(path,sheet_name=None)
#print(list(sheets.keys()))

def start():
	for j in sheets.keys():
		print("设备名：" + j)
		sheet = j
		select(sheet,path,j)
		print("\n\n\n")
		#contin = input("输入“没问题”开始下一个系统：")
	#sleep(100)

if __name__ == "__main__":
	start()
