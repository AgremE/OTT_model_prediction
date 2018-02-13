"""
Cleaning up the file for complaining data
"""

import pandas as pd
import numpy as np
import glob
FILE_NAMES = []
#dummmy data for file name
FILE_NAMES.append("01 - Jan'17 Complaint.xlsx")
SHEET_NAME = ['CTT','SR','AMDOCS']

def get_file_name(path):
	files  = glob.glob(path+"/*.xlsx")
	return files
"""
Function to read all the name of the file for cleaning up
Input: List of file name (String)
Output: List of dataframe from execl file split into sheet name
"""
"""
Read the excel file from provided path
"""
def read_excel_files(file_names):
	complete_data = {}
	for x in range(len(file_names)):
		for y in range(len(SHEET_NAME)):
			data = pd.read_excel(file_names[x],dtype=object,sheet_name=SHEET_NAME[y])
			complete_data[file_names[x]+"|"+SHEET_NAME[y]] = data
			print(file_names[x]+"|"+SHEET_NAME[y])
	return complete_data

"""
This function will product a data in csv form based on individual sheet name
Input:
	data: a pandas dataframe for specific sheet_name
	sheet_name: a string value
Output:
	string value represent in the form of CSV if condition is meet otherwise is NULL value
"""
def clean_based_sheet_name(data,sheet_name):
	result = ""
	if sheet_name == 'CTT':
		for i in range(len(data)):
			result=result+"CTT"+","+str(data.at[i,"TT_NUMBER"])+","+str(data.at[i,"PRE_OR_POST"])\
						+","+str(data.at[i,"SERIAL_NUMBER"])+","+str(data.at[i,"CUSTOMER_SEGMENT_GROUP"])\
							+","+str(data.at[i,"SOURCE"])+","+str(data.at[i,"STATUS"])+","+str(data.at[i,"TT_TYPE"])\
								+","+str(data.at[i,"CATEGORY"])+",,"+str(data.at[i,"PLAN"])+","+str(data.at[i,"PRODUCT"])+","\
									+","+str(data.at[i,"PROBLEM_CAUSED"])+","+str(data.at[i,"RESOLUTION_CODE"])+","+str(data.at[i,"OWNER_ID"])\
										+","+str(data.at[i,"CREATOR_ID"])+","+str(data.at[i,"CREATED_DATE"])+","+str(data.at[i,"CLOSED_DATE"])\
											+","+str(data.at[i,"ESCALATION_TIME"])+",,"+str(data.at[i,"NETWORK_INDICATOR"])+","+str(data.at[i,"REGION"])\
												+","+str(data.at[i,"SERVICE_AFFECTED"])+"\n"
	elif sheet_name == 'SR':
		for i in range(len(data)):
			result=result+"SR"+","+str(data.at[i,"SR_NUMBER"])+","+str(data.at[i,"CON_TYPE"])\
						+","+str(data.at[i,"SERVICE_NUMBER"])+","+str(data.at[i,"CUSTOMER_SEGMENT_GROUP"])\
							+","+str(data.at[i,"SOURCE"])+","+str(data.at[i,"SR_STATUS"])+","+str(data.at[i,"SR_TYPE"])\
								+","+str(data.at[i,"CATEGORY"])+","+str(data.at[i,"SUBCATEGORY"])+","+str(data.at[i,"PLAN"])+","+str(data.at[i,"PRODUCT_NAME"])\
									+","+str(data.at[i,"PRODUCT_TYPE"])+","+str(data.at[i,"PROBLEM_CAUSE"])+","+str(data.at[i,"RESOLUTION_CODE"])+","+str(data.at[i,"OWNER_ID"])\
										+","+str(data.at[i,"CREATOR_ID"])+","+str(data.at[i,"CREATED_DATE"])+","+str(data.at[i,"CLOSED_DATE"])\
											+","+str(data.at[i,"ESCALATION_TIME"])+","+str(data.at[i,"DISPUTE_AMOUNT"])+",,"+str(data.at[i,"SD_REGION"])+",\n"
	elif sheet_name == 'AMDOCS':
		for i in range(len(data)):
			result=result+"AMDOCS"+",,,"+str(data.at[i,"Service Num"])+",,,"+str(data.at[i,"Status"])+","+str(data.at[i,"Type"])\
								+",,,,,"+str(data.at[i,"Prod Type"])+","+str(data.at[i,"Prob Category"])\
									+",,"+str(data.at[i,"Owner"])+",,"+str(data.at[i,"Creation Time"])+","+str(data.at[i,"X Close Date"])\
											+",,,,,\n"
	return result
"""
"""
	
def cleaning_up(complete_data):
	complete_clean_up = ""
	for key,value in complete_data.items():
		file_name,sheet_name = key.split("|")
		complete_clean_up = complete_clean_up + clean_based_sheet_name(value,sheet_name)
	return complete_clean_up

def main():
	print("Job begins")
	data = "FILE_TYPE,COMPL_NUMBER,CON_TYPE,SERVICE_NUMBER,CUSTOMER_SEGMENT_GROUP,\
				SOURCE,STATUS,TYPE,CATEGORY,SUBCATEGORY,PLAN,PRODUCT,PRODUCT_TYPE,PROBLEM_CAUSE,\
					RESOLUTION_CODE,OWNER_ID,CREATOR_ID,CREATED_DATE,CLOSED_DATE,ESCALATION_TIME,DISPUTE_AMOUNT,\
						NETWORK_INDICATOR,REGION,SERVICE_AFFECTED\n"
	data = data + cleaning_up(read_excel_files(get_file_name("../Data")))
	with open("complain_file_after_clean_up.csv","w") as open_file:
		open_file.write(data)
	print("Job Finish")
main()
		
	