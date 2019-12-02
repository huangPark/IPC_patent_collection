import requests
from bs4 import BeautifulSoup
import re
import xlwings as xw
import pandas as pd
import time
import os

import openpyxl
import json

# get the total number of all patents
def parse(data):
	if data.status_code==200:
		data = json.loads(data.text)
		num = data['resultPagination']['totalCount']# get the total number of all patents
		# print(num)
	return num

def page_parse(data):
	# After each page is crawled, the results are returned
	# Collect all the results of a category and save it in a file
	datas = []
	if data.status_code==200:
		data = json.loads(data.text)
		num = data['resultPagination']['totalCount']# get the total number of all patents
		print(num)
		
		results = data['searchResultDTO']['searchResultRecord']# Get specific data for each page
		for result in results:
			
			TIVIEW = result['fieldMap']['TIVIEW']# patent name
			APO = result['fieldMap']['APO']# application no. <FONT>CN</FONT>
			APO = APO.replace('<FONT>CN</FONT>','CN')

			APD = result['fieldMap']['APD']# application date

			PN = result['fieldMap']['PN']# Public (Announcement) Number<FONT>CN</FONT>
			PN = PN.replace('<FONT>CN</FONT>','CN')

			PD = result['fieldMap']['PD']# Public (Announcement) Number
			IC = result['fieldMap']['IC']# IPC classification number
			PAVIEW = result['fieldMap']['PAVIEW']# Applicant (patent)
			INVIEW = result['fieldMap']['INVIEW']# inventor
			AC = result['fieldMap']['AC']# Applicant's Country (Province)
			PRD = result['fieldMap']['PRD']# Priority date
			lawStatus = result['lawStatus']# Literature type

			hh = [TIVIEW,APO,APD,PN,PD,IC,PAVIEW,INVIEW,AC,PRD,lawStatus]
			datas.append(hh)
	return datas
			


def main():
	url = 'http://pss-system.cnipa.gov.cn/sipopublicsearch/patentsearch/executeTableSearch0529-executeCommandSearch.shtml'
	
	headers = {
			'user-agent':'Mozilla/5.0 (Windows NT 6.1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/73.0.3683.103 Safari/537.36',
			'cookie':'Cookie: WEE_SID=zMoK_Avc6LN7Fgrt5FmGFBggcxxuld7IMlLYH-fIdD3Y_OYuMy1I!-692928076!683571502!1567847353308; IS_LOGIN=true; _gscu_1645064636=634844537jc21u20; avoid_declare=declare_pass; _gscu_2023327167=54302220lsfvz115; _va_ref=%5B%22%22%2C%22%22%2C1566544459%2C%22https%3A%2F%2Fwww.google.com%2F%22%5D; _va_id=0b2db2a502d00ea1.1566540769.2.1566544459.1566544459.; _gscbrs_1645064636=1; _gscs_1645064636=67847843t8spjq21|pv:1; JSESSIONID=zMoK_Avc6LN7Fgrt5FmGFBggcxxuld7IMlLYH-fIdD3Y_OYuMy1I!-692928076!683571502'
			}
	# wb = openpyxl.Workbook()
	# ws = wb.active
	# ws.append(['发明名称','申请号','申请日','公开（公告）号','公开（公告）日','IPC分类号','申请（专利权）人','发明人','申请人所在国（省）','优先权日','文献类型'])
	wb = openpyxl.load_workbook('results.xlsx')
	ws = wb[u'专利数据']

	start=1
	# The query parameter searchExp is passed in from the file
	# Each parameter start is fetched from 1. The server sets a limit. Only 10 results can be returned at a time.
	# (公开（公告）日=20070101:20181231 AND IPC分类号=(B60K6/00)) AND (发明类型=("U") AND 公开国家/地区/组织=(CN))
	searchExp = '(公开（公告）日=20070101:20181231 AND IPC分类号=({})) AND (发明类型=("U") AND 公开国家/地区/组织=(CN))'
	with open('2.txt','r')as f:
		for line in f.readlines():
			results = []
			# print(line)
			searchExp_t = searchExp.format(line.strip())
			# print(searchExp_t)
			data = {
					'searchCondition.searchExp':searchExp_t,
					'searchCondition.dbId':'VDB',
					'searchCondition.searchType':'Sino_foreign',
					"searchCondition.extendInfo['MODE']":'SEARCH_MODE',
					"searchCondition.extendInfo['STRATEGY']":'STRATEGY_CALCULATE',
					'resultPagination.start':start,
			}
			# The first time is to get the total number of pages per category
			# The above has constructed each category url and start=0
			html = requests.post(url=url,headers=headers,data=data,verify=False)
			html.encoding = 'utf-8'

			num = parse(html)# Get the total number of pages per category
			num = int(num)
			# print(num)
			num_page = num//10# Can only return 10 results at a time
			for i in range(0,num_page+1):
				start = 10*i+1
				data = {
					'searchCondition.searchExp':searchExp_t,
					'searchCondition.dbId':'VDB',
					'searchCondition.searchType':'Sino_foreign',
					"searchCondition.extendInfo['MODE']":'SEARCH_MODE',
					"searchCondition.extendInfo['STRATEGY']":'STRATEGY_CALCULATE',
					'resultPagination.start':start,
				}
				html = requests.post(url=url,headers=headers,data=data,verify=False)
				html.encoding = 'utf-8'
				result = page_parse(html)# Results for each page returned
				results.append(result)# Save the results of each page together

			# Save to file
			
			for result in results:
				for row in result:
					ws.append(row)

			wb.save('results.xlsx')

	
	


if __name__ == '__main__':
	main()