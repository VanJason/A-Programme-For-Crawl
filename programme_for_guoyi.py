#-*- coding:utf-8 -*-

"""
国义招标网
"""

import re
import urllib
from urllib import request
from bs4 import BeautifulSoup
import xlwt
from selenium import webdriver
from bs4 import BeautifulSoup
import time
import random

class Programme_guoyi():
	datalist = []
	
	def get_Web_guoyi(self,url,typetext,keyword):
		"""
		打开国义招标网并进行关键词搜索
		返回网站
		无UI浏览器访问
		"""
		driver = webdriver.PhantomJS()
		driver.set_window_size(800,600)
		# driver.set_page_load_timeout(30)
		# driver.set_script_timeout(30)
		driver.get(url)
		if typetext == "招标公告":			
			driver.find_element_by_link_text(typetext).click()
		else:
			driver.find_element_by_link_text(u"招标公告").click()
			# driver.find_element_by_xpath("//form[@id='aspnetForm']/table/tbody/tr[1]/td[2]/a/img").click()
		driver.find_element_by_name("Keyword").clear()
		driver.find_element_by_name("Keyword").send_keys(keyword)
		driver.find_element_by_css_selector("a > img").click()

		return driver
		driver.quite()

	def get_newWeb_guoyi(url):
		driver = webdriver.PhantomJS()
		driver.set_window_size(800,600)
		driver.get(url)

		return driver

	def get_newpageWeb_guoyi(url):
		driver = webdriver.PhantomJS()
		driver.set_window_size(800,600)
		driver.get(url)
		iframe = driver.find_element_by_xpath("iframe")
		driver.switch_to_frame(iframe)
		# driver.switch_to_default_content()

		return driver
		driver.quite()

	def get_programme_guoyi(self,info,startime,endtime,state):
		weblist = []

		if state == True:
			while True:
				pagedate = []

				soup = BeautifulSoup(info.page_source)
				website = soup.find_all(href=re.compile("snid"))

				for web in website:
					gotoweb = web['href']
					weblist.append(gotoweb)

				dataall = soup.find_all("td",text = re.compile("\d{2}-\d+-\d+"))
				for data in dataall:
					datastr = data.get_text()
					self.datalist.append(int("20"+datastr[0]+datastr[1]+datastr[3]+datastr[4]+datastr[6]+datastr[7]))

				try:
					nextpage = soup.find("input", id = "ctl00_PageContent_btnNextPage")
					judge = nextpage['disabled']
					break
				except:
					info.find_element_by_id("ctl00_PageContent_btnNextPage").click()
					print("翻1页")
					time.sleep(random.randint(3,5))
		else:
			while True:
				webindex = []
				dateindex = []
				site = 0

				soup = BeautifulSoup(info.page_source)

				dateevery = soup.find_all("td",text = re.compile("\d{2}-\d+-\d+"))
				website = soup.find_all(href=re.compile("snid"))

				# print(len(dateevery),len(website))
				# break

				for web in website:
					gotoweb = web['href']
					webindex.append(gotoweb)

				for date_unit in dateevery:
					datestr = date_unit.get_text()
					dateindex.append(int("20"+datestr[0]+datestr[1]+datestr[3]+datestr[4]+datestr[6]+datestr[7]))

				for dateindex_unit in dateindex:
					if startime<= dateindex_unit <= endtime:
						self.datalist.append(dateindex_unit)
						weblist.append(webindex[site])
					else:
						None
					site +=1

				if startime < dateindex[len(dateindex)-1]:
					try:
						nextpage = soup.find("input", id = "ctl00_PageContent_btnNextPage")
						judge = nextpage['disabled']
						break
					except:
						info.find_element_by_id("ctl00_PageContent_btnNextPage").click()
						print("翻1页")
						time.sleep(random.randint(3,5))
				else:
					break


		return weblist




				
	def get_title_guoyi(self,soup):
		"""
		获取项目名称的方法
		返回项目名称
		"""
		soup_name = soup.find('span', id="ctl00_PageContent_Label_Title")
		title = soup_name.get_text()

		return title
	def get_beginningtime_guoyi(self,soup):
		"""
		获取公布时间的方法
		返回公布时间
		"""
		soup_beginningtime = soup.find('span', id="ctl00_PageContent_Label_ShowDate")
		beginningtime = soup_beginningtime.get_text()

		return beginningtime

	def get_number_guoyi(self,soup):
		"""
		获取项目编号的方法
		返回项目编号
		"""
		soup_number = soup.find('span', id="ctl00_PageContent_Label_Code")
		number = soup_number.get_text()

		return number

	def get_account_guoyi(self,soup):
		"""
		招标预算不准确，无法抓取
		"""
		return ""

	def get_showtime_guoyi(self,soup):
		programmelist = []
		i = 0
		showtime = None
		programmeALL = soup.find('span', id="ctl00_PageContent_Label_Content")
		if programmeALL is not None:
			for programme_unit in programmeALL.stripped_strings:
				programmelist.append(programme_unit)
		else:
			None
		while i < len(programmelist):
			try:
				showtimeindex = programmelist[i].index(u"开标时间")
				if showtimeindex ==7:
					showtime = programmelist[i][13:]
				elif showtimeindex == 2:
					showtime = programmelist[i][7:]
				elif showtimeindex == 3:
					showtime = programmelist[i][8:]
				else:
					showtime = programmelist[i][14:]
				break
			except:
				i +=1
		return showtime

	def get_agentcompany_guoyi(self,soup):
		return "国际招标股份有限公司"

	def get_buyer_guoyi(self,soup):
		programmelist = []
		i = 0
		buyer = None
		programmeALL = soup.find('span', id="ctl00_PageContent_Label_Content")
		if programmeALL is not None:
			for programme_unit in programmeALL.stripped_strings:
				programmelist.append(programme_unit)
		else:
			None
		while i < len(programmelist):
			try:
				buyerindex = programmelist[i].index(u"招标人名称")
				buyer = programmelist[i][6:]
				break
			except:
				try:
					buyerindex = programmelist[i].index(u"采购单位：")
					buyer = programmelist[i][5:]
					break
				except:
					try:
						buyerindex = programmelist[i].index(u"招标人:")
						buyer = programmelist[i][4:]
						break
					except:
						i +=1
		return buyer

	def get_money_guoyi(self,soup):
		"""
		暂不实现
		"""
		return ""


	def get_detail_guoyi(self,WBall,filename,state1,state2,state3,state4,state5,state6,state7,state8):
		"""
		对项目内容进行细化操作
		返回
		"""
		wbk = xlwt.Workbook()
		sheet = wbk.add_sheet('sheet 1',cell_overwrite_ok=True)
		sheet.write(0,0,"招标公告日期")
		sheet.write(0,1,"链接")
		sheet.write(0,2,"地区")
		sheet.write(0,3,"招标机构")
		sheet.write(0,4,"采购单位")
		sheet.write(0,5,"项目名称")
		sheet.write(0,6,"采购内容")
		sheet.write(0,7,"开标日期")
		sheet.write(0,8,"中标公告时间")
		sheet.write(0,9,"中标公司")
		sheet.write(0,10,"总金额")
		sheet.write(0,11,"中标金额")
		sheet.write(0,12,"预算（元）")
		sheet.write(0,13,"中标公告链接")

		i = 1

		for web in WBall:
			html = urllib.request.urlopen('http://www.gmgit.com/Notice/BidInfo/' + web)
			content = html.read()
			html.close()

			soup = BeautifulSoup(content)

			message_list = []

			soup_message = soup.find('span', id="ctl00_PageContent_Label_Content")

			if soup_message is None:
				None
				# new_content = soup.find(src=re.compile("HtmShow"))
				# content_web = new_content['src']
				# new_page = 'http://www.gmgit.com/Notice/BidInfo/' + content_web

				# # soupother = BeautifulSoup(self.get_newWeb_guoyi(new_page).page_source)
				# html2 = urllib.request.urlopen(new_page)
				# content2 = 
				# newsoup_message = soupother.find('div', class_="Section1")

				# for mes1 in newsoup_message.stripped_strings:
				# 	message_list.append(mes1)
			else:
				for mes2 in soup_message.stripped_strings:
					message_list.append(mes2)
					"""将内容分块保存至数组"""

			if state1 is True:
				sheet.write(i,5,self.get_title_guoyi(soup))
			else:
				None
			if state2 is True:	
				sheet.write(i,12,self.get_account_guoyi(soup))
			else:
				None
			if state3 is True:
				sheet.write(i,0,self.get_beginningtime_guoyi(soup))
			else:
				None
			if state4 is True:	
				sheet.write(i,3,self.get_agentcompany_guoyi(soup))
			else:
				None
			if state5 is True:
				sheet.write(i,1,"http://www.gmgit.com/Notice/BidInfo/" + web)
			else:
				None
			if state6 is True:
				sheet.write(i,11,self.get_money_guoyi(soup))
			else:
				None
			if state7 is True:
				sheet.write(i,7,self.get_showtime_guoyi(soup))
			else:
				None
			if state8 is True:
				sheet.write(i,4,self.get_buyer_guoyi(soup))
			else:
				None
				
			i +=1

		print("sucess")

		wbk.save(filename)



