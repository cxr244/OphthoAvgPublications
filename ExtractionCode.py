from Bio import Entrez
import pandas as pd 
import time 

def search(query):
	Entrez.email = "cxr244@case.edu"
	handle = Entrez.esearch(db='pubmed', 
                            sort='relevance', 
                            retmax='20',
                            retmode='xml', 
                            term=query)
	results = Entrez.read(handle)
	return results

def fetch_details(id_list):
    ids = ','.join(id_list)
    Entrez.email = 'cxr244@case.edu'
    handle = Entrez.efetch(db='pubmed',
                           retmode='xml',
                           id=ids)
    results = Entrez.read(handle)
    return results

if __name__ == '__main__':
	# initialData = {"Skip": ["N","N"], "Name": ["Chandruganesh Rasendran", "Alexander Haueisen"]}
	df = pd.read_excel("C:/Chandru/CWRU/Research/CCF/OphthoAvgPubsPaper/SusAuthorsDataAnalysis.xlsx", sheet_name="InputForCode")
	df = pd.DataFrame(df, columns = ["Skip", 'Name','Affliation','LastName'])
	# results=search("Chandruganesh Rasendran")
	# df = pd.DataFrame(initialData, columns = ["Skip", "Name"])
	FinalList = pd.DataFrame()
	for index, row in df.iterrows():
		if(df.at[index,'Skip'] == "N"):
			# Do stuff 
			AuthorList = pd.DataFrame()
			print(row['Affliation'])
			if(row['Affliation'] == "Ignore"):
				print("Searching just name")
				results = search(row['Name'] + " [AU]")
			else:
				results = search(row['Name'] + " [AU] AND " + row['Affliation'] + " [Affiliation]")
			# results = search(row['Name'] + " [AU]")
			id_list = results['IdList']
			title = []
			date = []
			year = []
			month = []
			day = [] 
			authors = [] 
			journal = [] 
			firstauthorlastname = []
			firstauthorfirstname = []
			lastauthorlastname = []
			lastauthorfirstname = []
			names = []
			if(len(id_list) > 0):
				papers = fetch_details(id_list)
				# TODO for below, if there is an error, append nothing and move on 
				# Adding new rows if columns are too high (likely use length of columns and then add same number of rows )
				# Change how dates are recorded to be more user friendly if possible 
				for i, paper in enumerate(papers['PubmedArticle']):
					size = len(authors)
					add = "TRUE"
					finishLoop = "FALSE"
					for i in range(len(paper['MedlineCitation']['Article']['AuthorList'])):	
						try:
							if add == "TRUE" and paper['MedlineCitation']['Article']['AuthorList'][i].get('LastName',"None") != "None" and paper['MedlineCitation']['Article']['AuthorList'][i]['LastName'] == row['LastName']:
								if(paper['MedlineCitation']['Article']['AuthorList'][i]['ForeName'][0] == row['Name'][0]) and size == len(authors):
									if(len(paper['MedlineCitation']['Article']['AuthorList'][i]['AffiliationInfo']) == 0):
										authors.append("Error")
									else:
										authors.append(paper['MedlineCitation']['Article']['AuthorList'][i]['AffiliationInfo'][0]['Affiliation'])	
								else:
									add = "FALSE"
						except (KeyError,IndexError,ValueError):
							# Enters here if author of interest has no affliation OR there is an author without last name 
							nearestAffliation = "FALSE"
							for i in range(len(paper['MedlineCitation']['Article']['AuthorList'])):
								if nearestAffliation == "FALSE" and paper['MedlineCitation']['Article']['AuthorList'][i]['AffiliationInfo'] and add == "TRUE":
									authors.append(paper['MedlineCitation']['Article']['AuthorList'][i]['AffiliationInfo'][0]['Affiliation'])	
									nearestAffliation = "TRUE"
					if (size + 1 == len(authors)):
						title.append("{}) {}".format(i+1, paper['MedlineCitation']['Article']['ArticleTitle']))
						try:
							# date.append(paper['MedlineCitation']['Article']['ArticleDate']['Year'])
							if(len(paper['MedlineCitation']['Article']['ArticleDate']) == 0):
								date.append("Error")
							else:
								year.append(paper['MedlineCitation']['Article']['ArticleDate'][0]['Year'])
								month.append(paper['MedlineCitation']['Article']['ArticleDate'][0]['Month'])
								day.append(paper['MedlineCitation']['Article']['ArticleDate'][0]['Day'])
								date.append(month[-1] + "/"+ day[-1]+ "/" + year[-1])
						except (KeyError,IndexError,ValueError):
							date.append("Error")
						try:
							journal.append(paper['MedlineCitation']['Article']['Journal']['Title'])
						except (KeyError,IndexError,ValueError):
							journal.append("Error")
							# authors.append(paper['MedlineCitation']['Article']['AuthorList'])
						try:
							firstauthorlastname.append(paper['MedlineCitation']['Article']['AuthorList'][0].get('LastName',"None"))
						except (KeyError,IndexError,ValueError):
							try:
								firstauthorlastname.append(paper['MedlineCitation']['Article']['AuthorList'][1].get('LastName',"None"))
							except:
								firstauthorlastname.append("Error")
						try:
							firstauthorfirstname.append(paper['MedlineCitation']['Article']['AuthorList'][0].get('ForeName',"None"))
						except (KeyError,IndexError,ValueError):
							try:
								firstauthorfirstname.append(paper['MedlineCitation']['Article']['AuthorList'][1].get('ForeName',"None"))
							except:
								firstauthorfirstname.append("Error")
						try:
							lastauthorlastname.append(paper['MedlineCitation']['Article']['AuthorList'][-1].get('LastName',"None"))
						except (KeyError,IndexError,ValueError):
							try:
								lastauthorlastname.append(paper['MedlineCitation']['Article']['AuthorList'][-2].get('LastName',"None"))
							except:
								lastauthorlastname.append("Error")
						try:
							lastauthorfirstname.append(paper['MedlineCitation']['Article']['AuthorList'][-1].get('ForeName',"None"))
						except (KeyError,IndexError,ValueError):
							try:
								lastauthorfirstname.append(paper['MedlineCitation']['Article']['AuthorList'][-2].get('ForeName',"None"))
							except:
								lastauthorfirstname.append("Error")
				titleLength = len(title)
				for i in range(len(title)):
					names.append(row['Name'])	
					# df = pd.concat([df,df.tail(1)])
				# print(df)
			if(len(title) == 0):
				names.append(row['Name'])
			AuthorList["Title"] = title 
			AuthorList["Date"] = date
			AuthorList["Journal"] = journal 
			AuthorList["First Last Name"] = firstauthorlastname
			AuthorList["First First Name"] = firstauthorfirstname
			AuthorList["Last Last Name"] = lastauthorlastname
			AuthorList["Last First Name"] = lastauthorfirstname
			AuthorList["Authors"] = authors
			print(AuthorList)
			print(names)
			AuthorList.insert(0,"Name", names)

			FinalList = pd.concat([FinalList,AuthorList])		
			
	writer = pd.ExcelWriter("C:/Chandru/CWRU/Research/CCF/OphthoAvgPubsPaper/DataAnalysis2.xlsx", engine = "xlsxwriter")
	FinalList.to_excel(writer, 'Sheet1')
	writer.save()


 #         print("{}) {}".format(i+1, paper['MedlineCitation']['Article']['ArticleTitle']))
 #         print(paper['MedlineCitation']['Article']['ArticleDate'])
 #         print(paper['MedlineCitation']['Article']['Journal']['Title'])
 #         print(paper['MedlineCitation']['Article']['AuthorList'][0]['LastName'])
 #         print(paper['MedlineCitation']['Article']['AuthorList'][0]['ForeName'])
 #         print(paper['MedlineCitation']['Article']['AuthorList'][-1]['LastName'])
 #         print(paper['MedlineCitation']['Article']['AuthorList'][-1]['ForeName'])



# if __name__ == '__main__':
# 	df = pd.read_excel("C:/Chandru/CWRU/Research/CCF/OphthoAvgPubsPaper/DataAnalysis.xlsx", sheet_name="Sheet3")
# 	df = pd.DataFrame(df, columns = ["Skip", 'Resident'])
# 	for index, row in df.iterrows():
# 		if df.at[index,'Skip'] == "N":
# 			results=search(df.at(index,'Resident'))
# 			id_list = results['IdList']
# 			papers = fetch_details(id_list)
# 			title = []
# 			date = []
# 			firstauthorlastname = []
# 			firstauthorfirstname = []
# 			lastauthorlastname = []
# 			lastauthorfirstname = []
# 			for i, paper in enumerate(papers['PubmedArticle']):
# 				title.append("{}) {}".format(i+1, paper['MedlineCitation']['Article']['ArticleTitle']))
# 				date.append(paper['MedlineCitation']['Article']['ArticleDate'])
# 				firstauthorlastname.append(paper['MedlineCitation']['Article']['ArticleDate'])
# 				firstauthorfirstname.append(paper['MedlineCitation']['Article']['ArticleDate'])
# 				lastauthorlastname.append(paper['MedlineCitation']['Article']['ArticleDate'])
# 				lastauthorfirstname.append(paper['MedlineCitation']['Article']['ArticleDate'])
# 		writer = pd.ExcelWriter("C:/Chandru/CWRU/Research/CCF/OphthoAvgPubsPaper/DataAnalysis2.xlsx", engine = "xlsxwriter")
# 		df.to_excel(writer, 'Sheet1')
# 		writer.save()













			# pubmed search the name 
	# results = search('Alexander Haueisen')
	# id_list = results['IdList']
	# papers = fetch_details(id_list)
	# for i, paper in enumerate(papers['PubmedArticle']):
 #         print("{}) {}".format(i+1, paper['MedlineCitation']['Article']['ArticleTitle']))
 #         print(paper['MedlineCitation']['Article']['ArticleDate'])
 #         print(paper['MedlineCitation']['Article']['Journal']['Title'])
 #         print(paper['MedlineCitation']['Article']['AuthorList'][0]['LastName'])
 #         print(paper['MedlineCitation']['Article']['AuthorList'][0]['ForeName'])
 #         print(paper['MedlineCitation']['Article']['AuthorList'][-1]['LastName'])
 #         print(paper['MedlineCitation']['Article']['AuthorList'][-1]['ForeName'])



