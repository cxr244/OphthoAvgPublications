import pandas as pd 
from googlesearch import search
import urllib.request
from bs4 import BeautifulSoup
from urllib.request import Request, urlopen 

df = pd.read_excel("C:/Chandru/CWRU/Research/CCF/OphthoAvgPubsPaper/IFJournals.xlsx", sheet_name="Sheet1")
df = pd.DataFrame(df, columns = ["Journal", 'IF','Manual'])

FinalList = pd.DataFrame()
for index, row in df.iterrows():
	# if index < 8:
	query = df.at[index,"Journal"]
	print(query)
	ctr = 0 
	row = pd.DataFrame()
	IF = 0 
	for i in search(query + " Impact factor 2019 scijournal", tld="com", lang = "en", num = 5, start = 0, stop = 4):
		if "www.scijournal.org" in i and ctr == 0:
			ctr = 1 
			url = i 
			req = Request(url, headers={'User-Agent': 'Mozilla/5.0'})
			content = urlopen(req).read()
			parse = BeautifulSoup(content, 'html.parser')
			# Provide html elements' attributes to extract the data 
			text1 = parse.find_all('div', attrs={'class': 'num'})
			if(len(text1) > 0):
				nbr = list(text1[0])
				IF = nbr[0]
	data = { "Journal" : [query],
			  "ImpactFactor": [IF]}
	row = pd.DataFrame(data)
	# row = pd.DataFrame([[query, IF]], columns = ["Journal", "ImpactFactor"])
	FinalList = pd.concat([FinalList, row])
writer = pd.ExcelWriter("C:/Chandru/CWRU/Research/CCF/OphthoAvgPubsPaper/IFJournals2.xlsx", engine = "xlsxwriter")
FinalList.to_excel(writer, 'Sheet1')
writer.save()