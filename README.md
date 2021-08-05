# OphthoAvgPublications

The purpose of this project is to determine the average number of publications for ophthalmology applicants that matched in 2020-2021 academic year. Additionally, applicants were compared based on gender, attending a US News Top 40 Medical School, and matching at a top 20 doximity residency. 

ExtractionCode.py utilizes a python script to search for authors based on first and last name as well as affliation on Pubmed and download the data onto an excel sheet.

ExtractionCodeNoAffliation.py utilizes a python script to search for authors only based on first and last name on Pubmed and download the data onto an excel sheet.

IFMasterFile.xlsx was utilized as a master file that stored the impact factors of various journals and cross-referenced with the journals extracted from the python scripts. 

ResultsOfPCA.txt contains the coefficent results of principal component analysis of variables that are highly multicollinear with each other.

stats2.R was used to analyze the final data and produce the tables seen in the final manuscript. 
