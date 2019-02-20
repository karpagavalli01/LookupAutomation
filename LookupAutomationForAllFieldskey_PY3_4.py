''''
Program Description: python LookupAutomationForAllFieldskey.py E:\LENSOPTIC\Packages\LensOptic-5.14.1.0-Optic-3.4.3.2-Win64.eng-AUS_NZL.UV configGeneralANZRC2.ini

ConfigFile Description: configGeneral.ini file contains the datafiles,path for the datafiles,xpath,Delimiter

Author : KARPAGAVALLI K
Dated : 31-October-2017

Edit Log:
		1) Changed the scripts from Python 2.7 to 3.4 - KIK:06-Apr-2018
'''
import os
import re
import sys
import xlsxwriter
import xlrd
from configobj import ConfigObj
from collections import defaultdict
from more_itertools import unique_everseen
from more_itertools import unique_everseen
from collections import Counter
#from UserString import MutableString


#Required Variables
dict_of_list=defaultdict(list)#Dictionary to store each module
listofdata=[]
uniq_dict = {}
uniq_dict_std ={}
uniq_dict_Major = {}

listofdataIndustry=list()
listIndustryRuleid = list()
anzlistofdataIndustry=list()
anzlistIndustryRuleid = list()
uslistofdataIndustry=list()
uslistIndustryRuleid = list()


skillsMapping_dict = {}
dict_location = defaultdict(list)
dictSA = defaultdict(list)
dictLMR = defaultdict(list)
dictLL = defaultdict(list)
listLL = list()
dictX = defaultdict(list)
dictRE = defaultdict(list)

nzdictSA = defaultdict(list)
nzdictLMR = defaultdict(list)
nzdictLL = defaultdict(list)
nzlistLL = list()
nzdictX = defaultdict(list)
nzdictRE = defaultdict(list)

#USLocationmodule
dictcity = defaultdict(list)
dictlocations = defaultdict(list)
dictzips = defaultdict(list)
dictcounty = defaultdict(list)
dictmsa = defaultdict(list)
dictlma = defaultdict(list)

instance_path=sys.argv[1] # Get from the user
config=ConfigObj(sys.argv[2]) # config file name get from the user

value=config["Output"]
Outputfile=value["Outputfile"]

def SkillsMapping(section):
		#print "maping func"
		value = config[section]
		datafiles=[v for k,v in value.iteritems() if k.startswith('Data')]
		for file in (datafiles):
			datafile=file.split(';')
			filename=datafile[0]
			fname = "%s\\%s"%(instance_path,filename)
			if (os.stat(fname).st_size == 0):
							#print fname 
							continue
			else:
							key = int(datafile[1])-1
							value = int(datafile[2])-1
							filecontent=open(instance_path+"\\"+filename,"r")
							delimit1 = datafile[3].lstrip(' ')
							delimit2 = datafile[4].lstrip(' ')
							#print delimit1
							#print delimit2
							if(re.match(r'\'\\t',delimit1)):
									for line in filecontent:
											#print line
											content=line.split("\t")
											if( content[key] != "NA"):
													if(re.match(r'\'\:',delimit2)):
															temp = content[key].split(':')
															col = int(datafile[5])-1
															skillsMapping_dict[temp[col]] = content[value]
													elif(re.match(r'\'\;',delimit2)):
															temp = content[key].split(';')
															col = int(datafile[5])-1
															skillsMapping_dict[temp[col]] = content[value]
													elif(re.match(r'\'\ ',delimit2)):
															temp = content[key].split(' ')
															col = int(datafile[5])-1
															skillsMapping_dict[temp[col]] = content[value]
							elif(re.match(r'\'\:',delimit1)):
									for line in filecontent:
											#print line
											content=line.split(":")
											if( content[key] != "NA"):
													if(re.match(r'\'\:',delimit2)):
															temp = content[key].split(':')
															col = int(datafile[5])-1
															skillsMapping_dict[temp[col]] = content[value]
													elif(re.match(r'\'\;',delimit2)):
															temp = content[key].split(';')
															col = int(datafile[5])-1
															skillsMapping_dict[temp[col]] = content[value]
													elif(re.match(r'\'\ ',delimit2)):
															temp = content[key].split(' ')
															col = int(datafile[5])-1
															skillsMapping_dict[temp[col]] = content[value]
							elif(re.match(r'\'\;',delimit1)):
									for line in filecontent:
											#print line
											content=line.split(";")
											if( content[key] != "NA"):
													if(re.match(r'\'\:',delimit2)):
															temp = content[key].split(':')
															col = int(datafile[5])-1
															skillsMapping_dict[temp[col]] = content[value]
													elif(re.match(r'\'\;',delimit2)):
															temp = content[key].split(';')
															col = int(datafile[5])-1
															skillsMapping_dict[temp[col]] = content[value]
													elif(re.match(r'\'\ ',delimit2)):
															temp = content[key].split(' ')
															col = int(datafile[5])-1
															skillsMapping_dict[temp[col]] = content[value]				
											else:
													continue
												
										
def Skills(section):
		value = config[section]
		skills = list()
		datafiles=[v for k,v in value.iteritems() if k.startswith('Data')]
		for file in (datafiles):
			datafile=file.split(';')
			filename=datafile[0]
			fname = "%s\\%s"%(instance_path,filename)
			if (os.stat(fname).st_size == 0):
							#print fname 
							continue
			else:
							columntoread=datafile[1]
							filecontent=open(instance_path+"\\"+filename,"r")
							delimit1 = datafile[2].lstrip(' ')
							delimit2 = datafile[3].lstrip(' ')
							#print delimit2
							if(re.match(r'\'\\t',delimit1)):
									for line in filecontent:
											#print line
											content=line.split("\t")
											col=int(columntoread)-1
											if( content[col] != "NA"):
													if(re.match(r'\'\:',delimit2)):
															#print "mathched"
															temp = content[col].split(':')
															if( section == "SkillsRuleId"):
																	incol = len(temp)-1
																	#incol = 1
																	skills.append(temp[incol])
															else:
																	if(len(temp) > 2):
																			var = "%s:%s"%(temp[0],temp[1])
																			skills.append(var)
																	else:
																			skills.append(temp[0])
													elif(re.match(r'\'\;',delimit2)):
															temp = content[col].split(';')
															if( section == "SkillsRuleId"):
																	incol = len(temp)-1
																	skills.append(temp[incol])
															else:
																	if(len(temp) > 2):
																			var = "%s:%s"%(temp[0],temp[1])
																			skills.append(var)
																	else:
																			skills.append(temp[0])
													elif(re.match(r'\'\ ',delimit2)):
															temp = content[col].split(' ')
															if( section == "SkillsRuleId"):
																	incol = len(temp)-1
																	skills.append(temp[incol])
															else:
																	if(len(temp) > 2):
																			var = "%s:%s"%(temp[0],temp[1])
																			skills.append(var)
																	else:
																			skills.append(temp[0])
							elif(re.match(r'\'\:',delimit1)):
									for line in filecontent:
											#print line
											content=line.split(":")
											col=int(columntoread)-1
											if( content[col] != "NA"):
													if(re.match(r'\'\:',delimit2)):
															#print "mathched"
															temp = content[col].split(':')
															if( section == "SkillsRuleId"):
																	incol = len(temp)-1
																	skills.append(temp[incol])
															else:
																	if(len(temp) > 2):
																			var = "%s:%s"%(temp[0],temp[1])
																			skills.append(var)
																	else:
																			skills.append(temp[0])
													elif(re.match(r'\'\;',delimit2)):
															temp = content[col].split(';')
															if( section == "SkillsRuleId"):
																	incol = len(temp)-1
																	skills.append(temp[incol])
															else:
																	if(len(temp) > 2):
																			var = "%s:%s"%(temp[0],temp[1])
																			skills.append(var)
																	else:
																			skills.append(temp[0])
													elif(re.match(r'\'\ ',delimit2)):
															temp = content[col].split(' ')
															if( section == "SkillsRuleId"):
																	incol = len(temp)-1
																	skills.append(temp[incol])
															else:
																	if(len(temp) > 2):
																			var = "%s:%s"%(temp[0],temp[1])
																			skills.append(var)
																	else:
																			skills.append(temp[0])
							elif(re.match(r'\'\;',delimit1)):
									for line in filecontent:
											#print line
											content=line.split(";")
											col=int(columntoread)-1
											if( content[col] != "NA"):
													if(re.match(r'\'\:',delimit2)):
															#print "mathched"
															temp = content[col].split(':')
															if( section == "SkillsRuleId"):
																	incol = len(temp)-1
																	skills.append(temp[incol])
															else:
																	if(len(temp) > 2):
																			var = "%s:%s"%(temp[0],temp[1])
																			skills.append(var)
																	else:
																			skills.append(temp[0])
													elif(re.match(r'\'\;',delimit2)):
															temp = content[col].split(';')
															if( section == "SkillsRuleId"):
																	incol = len(temp)-1
																	skills.append(temp[incol])
															else:
																	if(len(temp) > 2):
																			var = "%s:%s"%(temp[0],temp[1])
																			skills.append(var)
																	else:
																			skills.append(temp[0])
													elif(re.match(r'\'\, ',delimit2)):
															temp = content[col].split(' ')
															if( section == "SkillsRuleId"):
																	incol = len(temp)-1
																	skills.append(temp[incol])
															else:
																	if(len(temp) > 2):
																			var = "%s:%s"%(temp[0],temp[1])
																			skills.append(var)
																	else:
																			skills.append(temp[0])
																	
							else:
									continue
							result = list(unique_everseen(skills))
							#print result
							for var in result:
								dict_of_list[section].append(var)
				
def handle_two_delimiter(file,section):
	re_section=section
	datafile=file
	print(datafile)
	datafilecontent=datafile.split(';')
	#print datafilecontent
	filename=datafilecontent[0]
	columntoread=datafilecontent[1]
	Delimiter=datafilecontent[2]
	#Delimiter = delimiter.lstrip(' ')
	contentDelimiter=datafilecontent[3]
	col_extract=datafilecontent[4]
	filecontent=open(instance_path+"\\"+filename,"r")
	if(re.match(r'\'\\t',Delimiter)):
			for line in filecontent:
					#print line
					content=line.split("\t")
					col=int(columntoread)-1
					temp=content[col].split(":")
					if(temp[int(col_extract)]!="NA"):
							#print temp[int(col_extract)]
							dict_of_list[re_section].append(temp[int(col_extract)])
def AULocation(section):
	value=config[section]
	delimit=[v for k,v in value.iteritems() if k.startswith('Delimiter')]
	datafiles=[v for k,v in value.iteritems() if k.startswith('Data')]
	count = 0
	
	for file in (datafiles):
		datafile=file.split(';')
		filename=datafile[0]
		fname = "%s\\%s"%(instance_path,filename)
		if (os.stat(fname).st_size == 0):
						#print fname 
						continue
		elif(count == 0):
						Delimiter = delimit[count]
						count = count + 1
						columntoread=datafile[1]
						filecontent=open(instance_path+"\\"+filename,"r")
						if(re.match(r'\\t',Delimiter)):
								for line in filecontent:
										if line.startswith('#'):#to ignore the comment line
												#print line
												continue
										else:
												content=line.split("\t")
												col=int(columntoread)-1 #key 
												for val in range(len(datafile)):
														if val in [0,1]:
																continue
														else:
																colval = int(datafile[val])-1
																if( colval == -1):
																		dictSA[content[col]].append("NA")#dummy values
																else:
																		if(val == len(datafile)-1):
																				SA2 = content[colval].rstrip('\n')
																				SA3 = int(str(SA2)[:5])
																				SA4 = int(str(SA2)[:3])
																				dictSA[content[col]].append(SA2)
																				dictSA[content[col]].append(SA3)
																				dictSA[content[col]].append(SA4)
																		else:
																				dictSA[content[col]].append(content[colval])
		elif(count == 1):
						Delimiter = delimit[count]
						count = count + 1
						columntoread=datafile[1]
						filecontent=open(instance_path+"\\"+filename,"r")
						if(re.match(r'\\t',Delimiter)):
								for line in filecontent:
										if line.startswith('#'):#to ignore the comment line
												#print line
												continue
										else:
												content=line.split("\t")
												col=int(columntoread)-1 #key
												if content[col] in dictLMR.keys():
														continue
												else:
														for val in range(len(datafile)):
																if val in [0,1]:
																		continue
																else:
																		colval = int(datafile[val])-1
																		if( colval == -1):
																				dictLMR[content[col]].append("NA")#dummy values
																		else:
																				if(val == len(datafile)-1):
																						final = content[colval].rstrip('\n')
																						dictLMR[content[col]].append(final)
																						dictLMR[content[col]].append("AUS")
																						dictLMR[content[col]].append("000")
																				else:
																						dictLMR[content[col]].append(content[colval])
		elif(count == 2):
						Delimiter = delimit[count]
						count = count + 1
						columntoread=datafile[1]
						filecontent=open(instance_path+"\\"+filename,"r")
						if(re.match(r'\\t',Delimiter)):
								for line in filecontent:
										if line.startswith('#'):#to ignore the comment line
												#print line
												continue
										else:
												content=line.split("\t")
												col=int(columntoread)-1
												listLL.append(content[col])
		elif(count == 3):
						Delimiter = delimit[count]
						count = count + 1
						columntoread=datafile[1]
						filecontent=open(instance_path+"\\"+filename,"r")
						if(re.match(r',',Delimiter)):
								for line in filecontent:
										if line.startswith('#'):#to ignore the comment line
												#print line
												continue
										else:
												content=line.split(",")
												col=int(columntoread)-1
												temp = content[col].lstrip('"')
												temp2 = temp.rstrip('"')
												if temp2 in listLL:
														if temp2 in dictLL.keys():
																continue
														else:
																for val in range(len(datafile)):
																		if val in [0,1]:
																				continue
																		else:
																				colval = int(datafile[val])-1
																				if( colval == -1):
																						dictLL[temp2].append("NA")#dummy values
																				else:
																						if(val == len(datafile)-1):
																								final = content[colval].rstrip('\n')
																								dictLL[temp2].append(final)
																						else:
																								dictLL[temp2].append(content[colval])
												else:
														continue
		elif(count == 4):
						Delimiter = delimit[count]
						count = count + 1
						columntoread=datafile[1]
						filecontent=open(instance_path+"\\"+filename,"r")
						if(re.match(r'\\t',Delimiter)):
								for line in filecontent:
										if line.startswith('#'):#to ignore the comment line
												#print line
												continue
										else:
												content=line.split("\t")
												col=int(columntoread)-1
												if content[col] in dictX.keys():
														continue
												else:
														for val in range(len(datafile)):
																if val in [0,1]:
																		continue
																else:
																		colval = int(datafile[val])-1
																		if( colval == -1):
																				dictX[content[col]].append("NA")#dummy values
																		else:
																				dictX[content[col]].append(content[colval])
						for key in dictX.keys():
								dictX[key].append("NA")
								dictX[key].append("NA")
								dictX[key].append("NA")
		elif(count == 5):
						Delimiter = delimit[count]
						count = count + 1
						columntoread=datafile[1]
						filecontent=open(instance_path+"\\"+filename,"r")
						if(re.match(r'\\t',Delimiter)):
								for line in filecontent:
										if line.startswith('#'):#to ignore the comment line
												#print line
												continue
										else:
												content=line.split("\t")
												col=int(columntoread)-1 #key
												if content[col] in dictRE.keys():
														continue
												else:
														for val in range(len(datafile)):
																if val in [0,1]:
																		continue
																else:
																		colval = int(datafile[val])-1
																		if( colval == -1):
																				dictRE[content[col]].append("NA")#dummy values
																		else:
																				#print max(datafile[0:)
																				if(val == len(datafile)-4):
																						final = content[colval].rstrip('\n')
																						dictRE[content[col]].append(final)
																				else:
																						 dictRE[content[col]].append(content[colval])
														
												
		for keyX in dictX.keys():
				keyX = keyX.rstrip('\n')
				if keyX in dictSA.keys():
						continue
				else:
						dictSA[keyX]=dictX[keyX]
						
'''		for key in dictLMR.keys():
				print "%s------%s"%(key,dictLMR[key])
'''

def NZLocation(section):
	value=config[section]
	delimit=[v for k,v in value.iteritems() if k.startswith('Delimiter')]
	datafiles=[v for k,v in value.iteritems() if k.startswith('Data')]
	count = 0
	
	for file in (datafiles):
				datafile=file.split(';')
				filename=datafile[0]
				fname = "%s\\%s"%(instance_path,filename)
				if (os.stat(fname).st_size == 0):
						#print fname 
						continue
				elif(count == 0):
						Delimiter = delimit[count]
						count = count + 1
						columntoread=datafile[1]
						filecontent=open(instance_path+"\\"+filename,"r")
						if(re.match(r'\\t',Delimiter)):
								for line in filecontent:
										if line.startswith('#'):#to ignore the comment line
												#print line
												continue
										else:
												content=line.split("\t")
												col=int(columntoread)-1 #key 
												for val in range(len(datafile)):
														if val in [0,1]:
																continue
														else:
																colval = int(datafile[val])-1
																if( colval == -1):
																		nzdictSA[content[col]].append("NA")#dummy values
																else:
																		if(val == len(datafile)-1):
																				SA2 = content[colval]
																				SA3 = "NA"
																				SA4 = "NA"
																				nzdictSA[content[col]].append(SA2)
																				nzdictSA[content[col]].append(SA3)
																				nzdictSA[content[col]].append(SA4)
																		else:
																				nzdictSA[content[col]].append(content[colval])
				elif(count == 1):
						Delimiter = delimit[count]
						count = count + 1
						columntoread=datafile[1]
						filecontent=open(instance_path+"\\"+filename,"r")
						if(re.match(r'\\t',Delimiter)):
								for line in filecontent:
										if line.startswith('#'):#to ignore the comment line
												#print line
												continue
										else:
												content=line.split("\t")
												col=int(columntoread)-1 #key
												if content[col] in nzdictLMR.keys():
														continue
												else:
														for val in range(len(datafile)):
																if val in [0,1]:
																		continue
																else:
																		colval = int(datafile[val])-1
																		if( colval == -1):
																				nzdictLMR[content[col]].append("NA")#dummy values
																		else:
																				if(val == len(datafile)-1):
																						final = content[colval].rstrip('\n')
																						nzdictLMR[content[col]].append(final)
																						nzdictLMR[content[col]].append("NZL")
																						nzdictLMR[content[col]].append("000")
																				else:
																						nzdictLMR[content[col]].append(content[colval])
				elif(count == 2):
						Delimiter = delimit[count]
						count = count + 1
						columntoread=datafile[1]
						filecontent=open(instance_path+"\\"+filename,"r")
						if(re.match(r'\\t',Delimiter)):
								for line in filecontent:
										if line.startswith('#'):#to ignore the comment line
												#print line
												continue
										else:
												content=line.split("\t")
												col=int(columntoread)-1
												nzlistLL.append(content[col])
				elif(count == 3):
						Delimiter = delimit[count]
						count = count + 1
						columntoread=datafile[1]
						filecontent=open(instance_path+"\\"+filename,"r")
						if(re.match(r',',Delimiter)):
								for line in filecontent:
										if line.startswith('#'):#to ignore the comment line
												#print line
												continue
										else:
												content=line.split(",")
												col=int(columntoread)-1
												temp = content[col].lstrip('"')
												temp2 = temp.rstrip('"')
												if temp2 in nzlistLL:
														if temp2 in nzdictLL.keys():
																continue
														else:
																for val in range(len(datafile)):
																		if val in [0,1]:
																				continue
																		else:
																				colval = int(datafile[val])-1
																				if( colval == -1):
																						nzdictLL[temp2].append("NA")#dummy values
																				else:
																						if(val == len(datafile)-1):
																								final = content[colval].rstrip('\n')
																								nzdictLL[temp2].append(final)
																						else:
																								nzdictLL[temp2].append(content[colval])
												else:
														continue
				elif(count == 4):
						Delimiter = delimit[count]
						count = count + 1
						columntoread=datafile[1]
						filecontent=open(instance_path+"\\"+filename,"r")
						if(re.match(r'\\t',Delimiter)):
								for line in filecontent:
										if line.startswith('#'):#to ignore the comment line
												#print line
												continue
										else:
												content=line.split("\t")
												col=int(columntoread)-1
												if content[col] in nzdictX.keys():
														continue
												else:
														for val in range(len(datafile)):
																if val in [0,1]:
																		continue
																else:
																		colval = int(datafile[val])-1
																		if( colval == -1):
																				nzdictX[content[col]].append("NA")#dummy values
																		else:
																				nzdictX[content[col]].append(content[colval])
						for key in nzdictX.keys():
								nzdictX[key].append("NA")
								nzdictX[key].append("NA")
								nzdictX[key].append("NA")
				elif(count == 5):
						Delimiter = delimit[count]
						count = count + 1
						columntoread=datafile[1]
						filecontent=open(instance_path+"\\"+filename,"r")
						if(re.match(r'\\t',Delimiter)):
								for line in filecontent:
										if line.startswith('#'):#to ignore the comment line
												#print line
												continue
										else:
												content=line.split("\t")
												col=int(columntoread)-1 #key
												if content[col] in nzdictRE.keys():
														continue
												else:
														for val in range(len(datafile)):
																if val in [0,1]:
																		continue
																else:
																		colval = int(datafile[val])-1
																		if( colval == -1):
																				nzdictRE[content[col]].append("NA")#dummy values
																		else:
																				#print max(datafile[0:)
																				if(val == len(datafile)-4):
																						final = content[colval].rstrip('\n')
																						nzdictRE[content[col]].append(final)
																				else:
																						 nzdictRE[content[col]].append(content[colval])
														
												
				for keyX in nzdictX.keys():
					keyX = keyX.rstrip('\n')
					if keyX in nzdictSA.keys():
						continue
					else:
						nzdictSA[keyX]=nzdictX[keyX]
def USLocation(section):
	value=config[section]
	delimit=[v for k,v in value.iteritems() if k.startswith('Delimiter')]
	datafiles=[v for k,v in value.iteritems() if k.startswith('Data')]
	count = 0
	
	for file in (datafiles):
				datafile=file.split(';')
				filename=datafile[0]
				fname = "%s\\%s"%(instance_path,filename)
				if (os.stat(fname).st_size == 0):
						#print fname 
						continue
				elif(count == 0):
						print("city")
						Delimiter = delimit[count]
						count = count + 1
						#continue
						columntoread=datafile[1]
						filecontent=open(instance_path+"\\"+filename,"r")
						if(re.match(r'\\t',Delimiter)):
								for line in filecontent:
										if line.startswith('#'):#to ignore the comment line
												#print line
												continue
										else:
												content=line.split("\t")
												col=int(columntoread)-1 #key
												
												
												if content[col] in dictcity.keys():
														continue
												else:
														for val in range(len(datafile)):
																if val in [0,1]:
																		continue
																else:
																		colval = int(datafile[val])-1
																		if( colval == -1):
																				dictcity[content[col]].append("NA")#dummy values
																		else:
																				if(val == len(datafile)-3):
																						final = content[colval].rstrip('\n')
																						dictcity[content[col]].append(final)
																				else:
																						dictcity[content[col]].append(content[colval])
												dictcity[content[col]].append("USA")
							   
				elif(count == 1):
						print("zips")
						Delimiter = delimit[count]
						count = count + 1
						#continue
						columntoread=datafile[1]
						filecontent=open(instance_path+"\\"+filename,"r")
						if(re.match(r',',Delimiter)):
								for line in filecontent:
										if line.startswith('#'):#to ignore the comment line
												#print line
												continue
										else:
												content=line.split(",")
												col=int(columntoread)-1 #key
												temp = content[col].lstrip('"')
												key = temp.rstrip('"')
												
												if key in dictzips.keys():
														continue
												else:
														for val in range(len(datafile)):
																if val in [0,1]:
																		continue
																else:
																		colval = int(datafile[val])-1
																		if( colval == -1):
																				dictzips[key].append("NA")#dummy values
																		else:
																				if(val == len(datafile)-2):
																						final = content[colval].rstrip('\n')
																						dictzips[key].append(final)
																				else:
																						dictzips[key].append(content[colval])
				elif(count == 2):
						print("locations")
						Delimiter = delimit[count]
						count = count + 1
						#continue
						columntoread=datafile[1]
						filecontent=open(instance_path+"\\"+filename,"r")
						if(re.match(r',',Delimiter)):
								for line in filecontent:
										if line.startswith('#'):#to ignore the comment line
												#print line
												continue
										else:
												content=line.split(",")
												col=int(columntoread)-1 #key
												temp = content[col].lstrip('"')
												key = temp.rstrip('"')
												if key in dictlocations.keys():
														continue
												else:
														for val in range(len(datafile)):
																if val in [0,1]:
																		continue
																else:
																		colval = int(datafile[val])-1
																		if( colval == -1):
																				dictlocations[key].append("NA")#dummy values
																		else:
																				if(val == len(datafile)-2):
																						final = content[colval].rstrip('\n')
																						dictlocations[key].append(final)
																				else:
																						dictlocations[key].append(content[colval])
				elif(count == 3):
						print("county")
						Delimiter = delimit[count]
						count = count + 1
						#continue
						columntoread=datafile[1]
						filecontent=open(instance_path+"\\"+filename,"r")
						if(re.match(r'\\t',Delimiter)):
								for line in filecontent:
										if line.startswith('#'):#to ignore the comment line
												#print line
												continue
										else:
												content=line.split("\t")
												col=int(columntoread)-1 #key
												if content[col] in dictcounty.keys():
														continue
												else:
														for val in range(len(datafile)):
																if val in [0,1]:
																		continue
																else:
																		colval = int(datafile[val])-1
																		if( colval == -1):
																				dictcounty[content[col]].append("NA")#dummy values
																		else:
																				if(val == len(datafile)-1):
																						final = content[colval].rstrip('\n')
																						dictcounty[content[col]].append(final)
																				else:
																						dictcounty[content[col]].append(content[colval])
				elif(count == 4):
						print("msa")
						Delimiter = delimit[count]
						count = count + 1
						columntoread=datafile[1]
						filecontent=open(instance_path+"\\"+filename,"r")
						if(re.match(r'\\t',Delimiter)):
								for line in filecontent:
										if line.startswith('#'):#to ignore the comment line
												#print line
												continue
										else:
												content=line.split("\t")
												col=int(columntoread)-1 #key
												keysplit = content[col].split(' ')
												
												length = len(keysplit)
												if(length == 2):
														temp = keysplit[0]
														key = temp.title()
												elif(length == 3):
														citycounty = "%s %s"%(keysplit[1],keysplit[2])
														if(keysplit[2] == "CITY"):
																temp = "%s %s"%(keysplit[0],keysplit[1])
																key = temp.title()
														elif(keysplit[1] == "CITY"):
																temp = keysplit[0]
																key = temp.title()
														elif(citycounty == "CITY COUNTY"):
																temp = keysplit[0]
																key = temp.title()
												elif(length == 4):
														citycounty = "%s %s"%(keysplit[2],keysplit[3])
														if(citycounty == "CITY COUNTY"):
																temp = "%s %s"%(keysplit[0],keysplit[1])
																key = temp.title()
														else:
																temp = "%s %s %s"%(keysplit[0],keysplit[1],keysplit[2])
																key = temp.title()
												if key in dictmsa.keys():
														continue
												else:
														for val in range(len(datafile)):
																if val in [0,1]:
																		continue
																else:
																		colval = int(datafile[val])-1
																		if( colval == -1):
																				dictmsa[key].append("NA")#dummy values
																		else:
																				if(val == len(datafile)-1):
																						final = content[colval].rstrip('\n')
																						dictmsa[key].append(final)
																				else:
																						dictmsa[key].append(content[colval])

				elif(count == 5):
						print("lma")
						Delimiter = delimit[count]
						count = count + 1
						#continue
						columntoread=datafile[1]
						filecontent=open(instance_path+"\\"+filename,"r")
						if(re.match(r'\\t',Delimiter)):
								for line in filecontent:
										if line.startswith('#'):#to ignore the comment line
												#print line
												continue
										else:
												content=line.split("\t")
												col=int(columntoread)-1 #key
												keysplit = content[col].split(' ')
												
												length = len(keysplit)
												if(length == 2):
														temp = keysplit[0]
														key = temp.title()
												elif(length == 3):
														citycounty = "%s %s"%(keysplit[1],keysplit[2])
														if(keysplit[2] == "CITY"):
																temp = "%s %s"%(keysplit[0],keysplit[1])
																key = temp.title()
														elif(keysplit[1] == "CITY"):
																temp = keysplit[0]
																key = temp.title()
														elif(citycounty == "CITY COUNTY"):
																temp = keysplit[0]
																key = temp.title()
												elif(length == 4):
														citycounty = "%s %s"%(keysplit[2],keysplit[3])
														if(citycounty == "CITY COUNTY"):
																temp = "%s %s"%(keysplit[0],keysplit[1])
																key = temp.title()
														else:
																temp = "%s %s %s"%(keysplit[0],keysplit[1],keysplit[2])
																key = temp.title()
												if key in dictlma.keys():
														continue
												else:
														for val in range(len(datafile)):
																if val in [0,1]:
																		continue
																else:
																		colval = int(datafile[val])-1
																		if( colval == -1):
																				dictlma[key].append("NA")#dummy values
																		else:
																				if(val == len(datafile)-1):
																						final = content[colval].rstrip('\n')
																						dictlma[key].append(final)
																				else:
																						dictlma[key].append(content[colval])
		#for key in dictmsa.keys():
				#print "%s ----------- %s"%(key,dictmsa[key])
				

def Location(section):
	value=config[section]
	Delimiter=value["Delimiter"]
	datafiles=[v for k,v in value.iteritems() if k.startswith('Data')]
	for file in (datafiles):
				datafile=file.split(';')
				filename=datafile[0]
				fname = "%s\\%s"%(instance_path,filename)
				if (os.stat(fname).st_size == 0):
						#print fname 
						continue
				else:
						columntoread=datafile[1]
						filecontent=open(instance_path+"\\"+filename,"r")
						lastval = datafile[len(datafile)-1].lstrip()
						if(lastval.isdigit()):
								if(re.match(r'\\t',Delimiter)):
										for line in filecontent:
												if line.startswith('#'):#to ignore the comment line
														#print line
														continue
												else:
														content=line.split("\t")
														col=int(columntoread)-1 #key 
														for val in range(len(datafile)):
																if val in [0,1]:
																		continue
																else:
																		colval = int(datafile[val])-1
																		if( colval == -1):
																				dict_location[content[col]].append("NA")#dummy values
																		else:
																				dict_location[content[col]].append(content[colval])
								elif(re.match(r'\'\;',Delimiter)):
										for line in filecontent:
												if line.startswith('#'):#to ignore the comment line
														#print line
														continue
												else:
														content=line.split(";")
														col=int(columntoread)-1 #key 
														for val in range(len(datafile)):
																if val in [0,1]:
																		continue
																else:
																		colval = int(datafile[val])-1
																		if( colval == -1):
																				dict_location[content[col]].append("NA")#dummy values
																		else:
																				dict_location[content[col]].append(content[colval])
								elif(re.match(r'\'\:',Delimiter)):
										for line in filecontent:
												if line.startswith('#'):#to ignore the comment line
														#print line
														continue
												else:
														content=line.split(":")
														col=int(columntoread)-1 #key 
														for val in range(len(datafile)):
																if val in [0,1]:
																		continue
																else:
																		colval = int(datafile[val])-1
																		if( colval == -1):
																				dict_location[content[col]].append("NA")#dummy values
																		else:
																				dict_location[content[col]].append(content[colval])
								elif(re.match(r'\'\,',Delimiter)):
										for line in filecontent:
												if line.startswith('#'):#to ignore the comment line
														#print line
														continue
												else:
														content=line.split(",")
														col=int(columntoread)-1 #key 
														for val in range(len(datafile)):
																if val in [0,1]:
																		continue
																else:
																		colval = int(datafile[val])-1
																		if( colval == -1):
																				dict_location[content[col]].append("NA")#dummy values
																		else:
																				dict_location[content[col]].append(content[colval])

def CanonIntermediary(section):
		value = config[section]
		inter = list()
		datafiles=[v for k,v in value.iteritems() if k.startswith('Data')]
		for file in (datafiles):
			datafile=file.split(';')
			filename=datafile[0]
			fname = "%s\\%s"%(instance_path,filename)
			if (os.stat(fname).st_size == 0):
						#print fname 
						continue
			elif (len(datafile)>2):
						columntoread=datafile[1]
						filecontent=open(instance_path+"\\"+filename,"r")
						delimit1 = datafile[2].lstrip(' ')
						delimit2 = datafile[3].lstrip(' ')
						lsincol = datafile[4].lstrip(' ')
						rsincol = lsincol.rstrip(' ')
						incol = int(rsincol)
						#print incol
						if(re.match(r'\'\\t',delimit1)):
								for line in filecontent:
										#print line
										content=line.split("\t")
										col=int(columntoread)-1
										if( content[col] != "NA"):
												if(re.match(r'\'\:',delimit2)):
														#print "mathched"
														temp = content[col].split(':')
														inter.append(temp[incol])
														#print temp[incol]
												elif(re.match(r'\'\;',delimit2)):
														temp = content[col].split(';')
														inter.append(temp[incol])
												elif(re.match(r'\'\ ',delimit2)):
														temp = content[col].split(' ')
														inter.append(temp[incol])
						elif(re.match(r':',delimit1)):
								for line in filecontent:
										#print line
										content=line.split("\t")
										col=int(columntoread)-1
										if( content[col] != "NA"):
												if(re.match(r'\'\:',delimit2)):
														#print "mathched"
														temp = content[col].split(':')
														inter.append(temp[incol])
														
												elif(re.match(r'\'\;',delimit2)):
														temp = content[col].split(';')
														inter.append(temp[incol])
												elif(re.match(r'\'\ ',delimit2)):
														temp = content[col].split(' ')
														inter.append(temp[incol])
						elif(re.match(r'\';',Delimiter)):
								for line in filecontent:
										#print line
										content=line.split("\t")
										col=int(columntoread)-1
										if( content[col] != "NA"):
												if(re.match(r'\'\:',delimit2)):
														#print "mathched"
														temp = content[col].split(':')
														inter.append(temp[incol])
														
												elif(re.match(r'\'\;',delimit2)):
														temp = content[col].split(';')
														inter.append(temp[incol])
												elif(re.match(r'\'\ ',delimit2)):
														temp = content[col].split(' ')
														inter.append(temp[incol])
			else:
						columntoread=datafile[1]
						filecontent=open(instance_path+"\\"+filename,"r")
						Delimiter=value["Delimiter"]
						#print filecontent
						if(re.match(r'\\t',Delimiter)):
								for line in filecontent:
										if line.startswith('#'):#Ignore the lines starting with '#'
											continue
										else:
												content=line.split("\t")
												col=int(columntoread)-1
												dict_of_list[section].append(content[col]) #creates a dictionary where the key = section name , values = deduplicated list of section
												
						elif(re.match(r':',Delimiter)):
								for line in filecontent:
										if line.startswith('#'):#Ignore the lines starting with '#'
											continue
										else:
												content=line.split(":")
												col=int(columntoread)-1
												dict_of_list[section].append(content[col])
						elif(re.match(r' ',Delimiter)):
								for line in filecontent:
										if line.startswith('#'):#Ignore the lines starting with '#'
											continue
										else:
												content=line.split(" ")
												col=int(columntoread)-1
												dict_of_list[section].append(content[col])
												
						elif(re.match(r'\';',Delimiter)):
								for line in filecontent:
										if line.startswith('#'):#Ignore the lines starting with '#'
											continue
										else:
												content=line.split(";")
												col=int(columntoread)-1
												dict_of_list[section].append(content[col])
		result = list(unique_everseen(inter))
		for var in result:
				dict_of_list[section].append(var)
def extract_module_content(section):
		value=config[section]
		Delimiter=value["Delimiter"]
		datafiles=[v for k,v in value.iteritems() if k.startswith('Data')]
		for file in (datafiles):
			datafile=file.split(';')
			filename=datafile[0]
			fname = "%s\\%s"%(instance_path,filename)
			if (os.stat(fname).st_size == 0):
						#print fname 
						continue
				
			else:
						
						columntoread=datafile[1]
						filecontent=open(instance_path+"\\"+filename,"r")
						#print filecontent
						if(re.match(r'\\t',Delimiter)):
								for line in filecontent:
										if line.startswith('#'):#Ignore the lines starting with '#'
											continue
										elif line in ['\n', '\r\n']:
												continue
										else:
												content=line.split("\t")
												col=int(columntoread)-1
												length=len(file.split(';'))
												#print length
												if(length>2):
														
														handle_two_delimiter(file,section)
												else:
														#print content[col]
														dict_of_list[section].append(content[col]) #creates a dictionary where the key = section name , values = deduplicated list of section
												
						elif(re.match(r':',Delimiter)):
								for line in filecontent:
										if line.startswith('#'):#Ignore the lines starting with '#'
											continue
										elif line in ['\n', '\r\n']:
												continue
										else:
												content=line.split(":")
												col=int(columntoread)-1
												length=len(file.split())
												if(length>2):
														handle_two_delimiter(file,section)
												else:
														dict_of_list[section].append(content[col])
						elif(re.match(r' ',Delimiter)):
								for line in filecontent:
										if line.startswith('#'):#Ignore the lines starting with '#'
											continue
										elif line in ['\n', '\r\n']:
												continue
										else:
												content=line.split(" ")
												col=int(columntoread)-1
												length=len(file.split())
												if(length>2):
														handle_two_delimiter(file,section)
												else:
														dict_of_list[section].append(content[col])
												
						elif(re.match(r'\';',Delimiter)):
								for line in filecontent:
										if line.startswith('#'):#Ignore the lines starting with '#'
											continue
										elif line in ['\n', '\r\n']:
												continue
										else:
												content=line.split(";")
												col=int(columntoread)-1
												length=len(file.split())
												if(length>2):
														handle_two_delimiter(file,section)
												else:
														dict_of_list[section].append(content[col])
					
def StdMajorCode(section):#Standard Majorcode Section
		value=config[section]
		Delimiter=value["Delimiter"]
		data=[v for k,v in value.iteritems() if k.startswith('Data')]
		for element in (data):
			#print element
			datafile=element.split(';')
			filename=datafile[0]
			fname = "%s\\%s"%(instance_path,filename)
			if (os.stat(fname).st_size == 0):
						#print fname 
						continue
			else:
						key = int(datafile[1])-1
						value = int(datafile[2])-1
						filecontent=open(instance_path+"\\"+filename,"r")
						if(Delimiter == r'\t'):
								print(Delimiter)
								for line in filecontent:
										if line.startswith('#'):#Ignore the lines starting with '#'
												continue
										else:
												content=line.split("\t")
												uniq_dict_std[content[key]]=content[value]
						elif(Delimiter== r':'):
								for line in filecontent:
										if line.startswith('#'):#Ignore the lines starting with '#'
												continue
										else:
												content=line.split(":")
												uniq_dict_std[content[key]]=content[value]
						elif(Delimiter== r' '):
								for line in filecontent:
										if line.startswith('#'):#Ignore the lines starting with '#'
												continue
										else:
												content=line.split(" ")
												uniq_dict_std[content[key]]=content[value]
						elif(Delimiter== r';'):
								for line in filecontent:
										if line.startswith('#'):#Ignore the lines starting with '#'
												continue
										else:
												content=line.split(";")
												uniq_dict_std[content[key]]=content[value]
def MajorCode(section):#Major code section
	value=config[section]
	delimit=[v for k,v in value.iteritems() if k.startswith('Delimiter')]
	data=[v for k,v in value.iteritems() if k.startswith('Data')]
	count = 0
	for element in (data):
		datafile=element.split(';')
		filename=datafile[0]
		fname = "%s\\%s"%(instance_path,filename)
		if (os.stat(fname).st_size == 0):
			continue
		else:
			Delimiter = delimit[count]
			count = count + 1
			value = int(datafile[1])-1
			key = int(datafile[2])-1
			filecontent=open(instance_path+"\\"+filename,"r")
			if(Delimiter == r'\t'):
				for line in filecontent:
					if line.startswith('#'):#Ignore the lines starting with '#'
						continue
					else:
						content=line.split("\t")
						#print content
						if "|" in content[key]:
							keys= content[key].split('|')
							for i in range(len(keys)):
								#print keys[i]
								uniq_dict_Major[keys[i]] = content[value]
						else:
							uniq_dict_Major[content[key]] = content[value]
			elif(Delimiter== r':'):
				#print Delimiter
				for line in filecontent:
					if line.startswith('#'):
						continue
					else:
						content=line.split(":")
						if "|" in content[key]:
							keys= content[key].split('|')
							for i in range(len(keys)):
								uniq_dict_Major[keys[i]] = content[value]
						else:
							uniq_dict_Major[content[key]] = content[value]
			elif(Delimiter== r' '):
				#print Delimiter
				for line in filecontent:
					if line.startswith('#'):
						continue
					else:
						content=line.split(" ")
						if "|" in content[key]:
							keys= content[key].split('|')
							for i in range(len(keys)):
								uniq_dict_Major[keys[i]] = content[value]
						else:
							uniq_dict_Major[content[key]] = content[value]
			elif(Delimiter== r';'):
				#print Delimiter
				for line in filecontent:
					if line.startswith('#'):
						continue
					else:
						content=line.split(";")
						if "|" in content[key]:
							keys= content[key].split('|')
							for i in range(len(keys)):
								uniq_dict_Major[keys[i]] = content[value]
						else:
							uniq_dict_Major[content[key]] = content[value]
def ConsolidatedInferredUKSIC(section):#Consolidated Inferred UKSIC Industry code
		value=config[section]
		Delimiter=value["Delimiter"]
		data=[v for k,v in value.iteritems() if k.startswith('Data')]
		for element in (data):
			datafile=element.split(';')
			filename=datafile[0]
			fname = "%s\\%s"%(instance_path,filename)
			if (os.stat(fname).st_size == 0):
						#print fname 
						continue
			else:
						columntoread=datafile[1]
						filecontent=open(instance_path+"\\"+filename,"r")
						if(Delimiter == r'\t'):
								#print Delimiter
								for line in filecontent:
										if line in ['\n', '\r\n','#']: #if the file has empty lines to ignore it 
												continue
										else:
												content=line.split("\t")
												col=int(columntoread)-1
												listofdataIndustry.append(content[col])
						elif(Delimiter== r':'):
								for line in filecontent:
										if line in ['\n', '\r\n','#']:
												continue
										else:
												content=line.split(":")
												col=int(columntoread)-1
												listofdataIndustry.append(content[col])
						elif(Delimiter== r' '):
								for line in filecontent:
										if line in ['\n', '\r\n','#']:
												continue
										else:
												content=line.split(" ")
												col=int(columntoread)-1
												listofdataIndustry.append(content[col])
						elif(Delimiter== r';'):
								for line in filecontent:
										if line in ['\n', '\r\n','#']:
												continue
										else:
												content=line.split(";")
												col=int(columntoread)-1
												listofdataIndustry.append(content[col])

def ConsolidatedInferredANZSIC(section):#Consolidated Inferred ANZSIC Industry code
		value=config[section]
		Delimiter=value["Delimiter"]
		data=[v for k,v in value.iteritems() if k.startswith('Data')]
		for element in (data):
			datafile=element.split(';')
			filename=datafile[0]
			fname = "%s\\%s"%(instance_path,filename)
			if (os.stat(fname).st_size == 0):
						#print fname 
						continue
			else:
						columntoread=datafile[1]
						filecontent=open(instance_path+"\\"+filename,"r")
						if(Delimiter == r'\t'):
								#print Delimiter
								for line in filecontent:
										if line in ['\n', '\r\n','#']: #if the file has empty lines to ignore it 
												continue
										elif line.startswith('#'):
												continue
										else:
												content=line.split("\t")
												col=int(columntoread)-1
												anzlistofdataIndustry.append(content[col])
						elif(Delimiter== r':'):
								for line in filecontent:
										if line in ['\n', '\r\n','#']:
												continue
										elif line.startswith('#'):
												continue
										else:
												content=line.split(":")
												col=int(columntoread)-1
												anzlistofdataIndustry.append(content[col])
						elif(Delimiter== r' '):
								for line in filecontent:
										if line in ['\n', '\r\n','#']:
												continue
										elif line.startswith('#'):
												continue
										else:
												content=line.split(" ")
												col=int(columntoread)-1
												anzlistofdataIndustry.append(content[col])
						elif(Delimiter== r';'):
								for line in filecontent:
										if line in ['\n', '\r\n','#']:
												continue
										elif line.startswith('#'):
												continue
										else:
												content=line.split(";")
												col=int(columntoread)-1
												anzlistofdataIndustry.append(content[col])
def ConsolidatedInferredNAICS(section):#Consolidated Inferred NAICS Industry code
		value=config[section]
		Delimiter=value["Delimiter"]
		data=[v for k,v in value.iteritems() if k.startswith('Data')]
		for element in (data):
			datafile=element.split(';')
			filename=datafile[0]
			fname = "%s\\%s"%(instance_path,filename)
			if (os.stat(fname).st_size == 0):
						#print fname 
						continue
			else:
						columntoread=datafile[1]
						filecontent=open(instance_path+"\\"+filename,"r")
						if(Delimiter == r'\t'):
								#print Delimiter
								for line in filecontent:
										if line in ['\n', '\r\n','#']: #if the file has empty lines to ignore it 
												continue
										elif line.startswith('#'):
												continue
										else:
												content=line.split("\t")
												col=int(columntoread)-1
												uslistofdataIndustry.append(content[col])
						elif(Delimiter== r':'):
								for line in filecontent:
										if line in ['\n', '\r\n','#']:
												continue
										elif line.startswith('#'):
												continue
										else:
												content=line.split(":")
												col=int(columntoread)-1
												uslistofdataIndustry.append(content[col])
						elif(Delimiter== r' '):
								for line in filecontent:
										if line in ['\n', '\r\n','#']:
												continue
										elif line.startswith('#'):
												continue
										else:
												content=line.split(" ")
												col=int(columntoread)-1
												uslistofdataIndustry.append(content[col])
						elif(Delimiter== r';'):
								for line in filecontent:
										if line in ['\n', '\r\n','#']:
												continue
										elif line.startswith('#'):
												continue
										else:
												content=line.split(";")
												col=int(columntoread)-1
												uslistofdataIndustry.append(content[col])																								
def ConsolidatedUKSIC_Ruleid(section):#Consolidated UKSIC RuleId section
		value=config[section]
		Delimiter=value["Delimiter"]
		data=[v for k,v in value.iteritems() if k.startswith('Data')]
		for element in (data):
			datafile=element.split(';')
			filename=datafile[0]
			fname = "%s\\%s"%(instance_path,filename)
			if (os.stat(fname).st_size == 0):
						#print fname 
						continue
			else:
						columntoread=datafile[1]
						filecontent=open(instance_path+"\\"+filename,"r")
						if(Delimiter == r'\t'):
								#print Delimiter
								for line in filecontent:
										if line in ['\n', '\r\n','#']: #if the file has empty lines to ignore it 
											continue
										else:
											content=line.split("\t")
											col=int(columntoread)-1
											listIndustryRuleid.append(content[col])
						elif(Delimiter== r':'):
								for line in filecontent:
										if line in ['\n', '\r\n','#']:
											continue
										else:
											content=line.split(":")
											col=int(columntoread)-1
											listIndustryRuleid.append(content[col])
						elif(Delimiter== r' '):
								for line in filecontent:
										if line in ['\n', '\r\n','#']:
											continue
										else:
											content=line.split(" ")
											col=int(columntoread)-1
											listIndustryRuleid.append(content[col])
						elif(Delimiter== r';'):
								for line in filecontent:
										if line in ['\n', '\r\n','#']:
											continue
										else:
											content=line.split(";")
											col=int(columntoread)-1
											listIndustryRuleid.append(content[col])
def ConsolidatedNAICSRuleid(section):#Consolidated NAICS RuleId section
		value=config[section]
		Delimiter=value["Delimiter"]
		data=[v for k,v in value.iteritems() if k.startswith('Data')]
		for element in (data):
			datafile=element.split(';')
			filename=datafile[0]
			fname = "%s\\%s"%(instance_path,filename)
			if (os.stat(fname).st_size == 0):
						#print fname 
						continue
			else:
						columntoread=datafile[1]
						filecontent=open(instance_path+"\\"+filename,"r")
						if(Delimiter == r'\t'):
								#print Delimiter
								for line in filecontent:
										if line in ['\n', '\r\n','#']: #if the file has empty lines to ignore it 
											continue
										else:
											content=line.split("\t")
											col=int(columntoread)-1
											uslistIndustryRuleid.append(content[col])
						elif(Delimiter== r':'):
								for line in filecontent:
										if line in ['\n', '\r\n','#']:
											continue
										else:
											content=line.split(":")
											col=int(columntoread)-1
											uslistIndustryRuleid.append(content[col])
						elif(Delimiter== r' '):
								for line in filecontent:
										if line in ['\n', '\r\n','#']:
											continue
										else:
											content=line.split(" ")
											col=int(columntoread)-1
											uslistIndustryRuleid.append(content[col])
						elif(Delimiter== r';'):
								for line in filecontent:
										if line in ['\n', '\r\n','#']:
											continue
										else:
											content=line.split(";")
											col=int(columntoread)-1
											uslistIndustryRuleid.append(content[col])
#for onet,BGTOcc,SSOC and LocalGovt				
def forOccupation(section):
	sheetList=list()
	value=config[section]
	WorkbookPath=value["WorkbookPath"]
	Sheets=[v for k,v in value.iteritems() if k.startswith('Sheet')]
	for element in (Sheets):
		sheetSplit = element.split(';')
		sheetName = sheetSplit[0].rstrip(' ')
		workbook1=xlrd.open_workbook(WorkbookPath)#creating a dictionary from the existing workbook
		sh= workbook1.sheet_by_name(sheetName)
		for i in range(sh.nrows):
			if(i==0):
				continue
			else:
				addTab= list()
				for j in range(len(sheetSplit)):
					if(j==0):
						continue
					else:
						colval = int(sheetSplit[j])
						value=sh.cell(i,colval).value
						addTab.append(str(value))
				#colValue= MutableString() #Changing python 2.7 to python 3.4 - KIK 06/04/2018
				colValue = ''
				# for k in range(len(addTab)):
					# colValue +="%s/gap"%(addTab[k])
				colValue = '/gap'.join(addTab)
				colValue=colValue.rstrip('/gap')
				sheetList.append(colValue)
	uniq_list = list(unique_everseen(sheetList))
	for val in uniq_list:
		valSplit = val.split('/gap')
		uniq_dict[valSplit[0]] = valSplit[1]
	value=config[section]
	Datafiles=[v for k,v in value.iteritems() if k.startswith('Datafile')]
	for file in (Datafiles):
		#print file 
		datafile = file.split(';')
		filename=datafile[0]
		fname = "%s\\%s"%(instance_path,filename)
		if (os.stat(fname).st_size == 0):
			continue
		else:
			columntoread=datafile[1]
			filecontent=open(instance_path+"\\"+filename,"r")
			for line in filecontent:
				if line.startswith('#'):
					continue
				else:
					#print line
					content=line.split("\t")
					#print content
					col=int(columntoread)-1
					#print col
					dict_of_list[section].append(content[col])

#for BGTOcc in SGP Locale
def SGP_BGTOcc(section):
	sheetList=list()
	value=config[section]
	#WorkbookPath=value["WorkbookPath"]
	Sheets=[v for k,v in value.iteritems() if k.startswith('Sheet')]
	for element in (Sheets):
		sheetSplit = element.split(';')
		sheetName = sheetSplit[1].rstrip(' ')
		sheetName = sheetName.lstrip(' ')
		workbook = sheetSplit[0].rstrip(' ')
		value=config[section]
		WorkbookPath = value[workbook]
		workbook1=xlrd.open_workbook(WorkbookPath)#creating a dictionary from the existing workbook
		sh= workbook1.sheet_by_name(sheetName)
		for i in range(sh.nrows):
			if(i==0):
				continue
			else:
				addTab= list()
				for j in range(len(sheetSplit)):
					if j in [0,1]:
						continue
					else:
						colval = int(sheetSplit[j])
						value=sh.cell(i,colval).value
						addTab.append(str(value))
				#colValue= MutableString() #Changing python 2.7 to python 3.4 - KIK 06/04/2018
				colValue = ''
				# for k in range(len(addTab)):
					# colValue +="%s/gap"%(addTab[k])
				colValue = '/gap'.join(addTab)
				colValue=colValue.rstrip('/gap')
				sheetList.append(colValue)
	uniq_list = list(unique_everseen(sheetList))
	for val in uniq_list:
		valSplit = val.split('/gap')
		uniq_dict[valSplit[0]] = valSplit[1]
	value=config[section]
	Datafiles=[v for k,v in value.iteritems() if k.startswith('Datafile')]
	for file in (Datafiles):
		#print file 
		datafile = file.split(';')
		filename=datafile[0]
		fname = "%s\\%s"%(instance_path,filename)
		if (os.stat(fname).st_size == 0):
			continue
		else:
			columntoread=datafile[1]
			filecontent=open(instance_path+"\\"+filename,"r")
			for line in filecontent:
				if line.startswith('#'):
					continue
				else:
					#print line
					content=line.split("\t")
					#print content
					col=int(columntoread)-1
					#print col
					dict_of_list[section].append(content[col])

#Read datafiles from configfile and extract the module content
for section in (config.keys()):
		print(section)
		if(section == "Output"):
				continue
		else:
				value=config[section]
				Presence = value["Presence"]
				if Presence in ["Yes","yes"]:
						if(section in ["Output","CanonYearsOfExperience","CanonSalary","CanonJobType","ConsolidatedOnetRank","ANZSCORank","NumberOfOpenings","CanonNumberOfOpenings","SSOCRank"]):
								continue
						elif(section in ["Onet","BGTOcc","LocalGovt","ANZSCO","ANZBGTOcc","SSOC"]):
								forOccupation(section)
						elif( section == "BGTOccSGP"):
								SGP_BGTOcc(section)
						elif( section == "StdMajorCIPCode"):
								StdMajorCode(section)
						elif( section == "MajorCode"):
								MajorCode(section)
						elif( section == "ConsolidatedInferredUKSIC"):
								ConsolidatedInferredUKSIC(section)
						elif( section == "ConsolidatedInferredANZSIC"):
								ConsolidatedInferredANZSIC(section)
						elif( section == "ConsolidatedInferredNAICS"):
								ConsolidatedInferredNAICS(section)
						elif( section == "ConsolidatedNAICSRuleid"):
								ConsolidatedNAICSRuleid(section)
						elif( section == "ConsolidatedUKSICRuleid"):
								ConsolidatedUKSIC_Ruleid(section)
						elif( section in ["CanonSkills","SkillsRuleId"]):
								Skills(section)
						elif( section == "CanonSkilltoSkillClusterMapping"):
								SkillsMapping(section)
						elif( section == "LocationSpecificInformation"):
								Location(section)
						elif( section == "AULocationSpecificInformation"):
								AULocation(section)
						elif( section == "NZLocationSpecificInformation"):
								NZLocation(section)
						elif section in ["USLocationSpecificInformation","SGPLocationSpecificInformation"]:
								USLocation(section)
						elif( section == "CanonIntermediary"):
								CanonIntermediary(section)
						else:
								extract_module_content(section)
				else:
						continue

#Write extracted content into an spreadsheet		
workbook = xlsxwriter.Workbook(str(Outputfile))	
format = workbook.add_format({'bold': True ,'bg_color':"#FFB6C1" , 'border':True})
format.set_font_size(11)
entryFormat = workbook.add_format({'border':True})
entryFormat.set_font_size(9)
hyperlink = workbook.add_format({'bold': True ,'font_color': 'blue' , 'underline': True , 'border':True })
hyperlink.set_font_size(9)
format1 = workbook.add_format({'bold': True, 'font_color': 'red'})
format1.set_font_size(9)
sectionlist = list()
for section in (config.keys()):
		if(section == "Output"):
				continue
		else:
				value=config[section]
				Presence = value["Presence"]
				
				if Presence in ["Yes","yes"]:
						sectionlist.append(section)
				else:
						continue
		
worksheet=workbook.add_worksheet("Summary")
worksheet.write(0,0," DE fields",format)
worksheet.write(0,1," XPath",format)
worksheet.write(0,2," Value Type",format)

value=config["Output"]
version = value["Version"]
worksheet.write(1,0,"DataElements Version",entryFormat)
worksheet.write(1,1,"/JobDoc/DataElementsRollup/@version",entryFormat)
worksheet.write(1,2,version,entryFormat)

worksheet.write(2,0,"JobID",entryFormat)
worksheet.write(2,1,"/JobDoc/DataElementsRollup/JobID",entryFormat)
worksheet.write(2,2,"Based on raw input document",entryFormat)

row = 3
for section in sectionlist:
		if(section == "ConsolidatedDegree"):
				value = config[section]
				Source = value["Source"]
				worksheet.write_url(row,0,  'internal:ConsolidatedDegree!A1',hyperlink,"ConsolidatedDegree")
				worksheet.write(row,1,"//JobDoc/DataElementsRollup/ConsolidatedDegree",entryFormat)
				worksheet.write(row,2,Source,entryFormat)
				row = row + 1
				worksheet.write_url(row,0,  'internal:ConsolidatedDegree!A2',hyperlink,"MaxDegree")
				worksheet.write(row,1,"//JobDoc/DataElementsRollup/CanonMaximumDegree",entryFormat)
				worksheet.write(row,2,Source,entryFormat)
				row = row + 1
				worksheet.write_url(row,0,  'internal:ConsolidatedDegree!A3',hyperlink,"MinDegree")
				worksheet.write(row,1,"//JobDoc/DataElementsRollup/CanonMinimumDegree",entryFormat)
				worksheet.write(row,2,Source,entryFormat)
				row = row + 1
				worksheet.write_url(row,0,  'internal:ConsolidatedDegree!A4',hyperlink,"PreferredDegree")
				worksheet.write(row,1,"//JobDoc/DataElementsRollup/CanonPreferredDegree",entryFormat)
				worksheet.write(row,2,Source,entryFormat)
				row = row + 1
				worksheet.write_url(row,0,  'internal:ConsolidatedDegree!A5',hyperlink,"RequiredDegree")
				worksheet.write(row,1,"//JobDoc/DataElementsRollup/CanonRequiredDegree",entryFormat)
				worksheet.write(row,2,Source,entryFormat)
				row = row + 1
				worksheet.write(row,0,"CanonOtherDegrees",entryFormat)
				worksheet.write(row,1,"/JobDoc/DataElementsRollup/CanonOtherDegrees",entryFormat)
				worksheet.write(row,2,"",entryFormat)
				row = row + 1
				worksheet.write(row,0,"OtherDegreeLevels",entryFormat)
				worksheet.write(row,1,"/JobDoc/DataElementsRollup/OtherDegreeLevels",entryFormat)
				worksheet.write(row,2,"",entryFormat)
				row = row + 1
		elif(section == "CanonJobtitle"):
				value = config[section]
				Source = value["Source"]
				XPath = value["XPath"]
				worksheet.write_url(row,0,  'internal:CanonJobtitle!A1',hyperlink,"CanonJobtitle")
				worksheet.write(row,1,XPath,entryFormat)
				worksheet.write(row,2,Source,entryFormat)
				row = row + 1
				worksheet.write(row,0,"CleanJobTitle",entryFormat)
				worksheet.write(row,1,"/JobDoc/DataElementsRollup/CleanJobTitle",entryFormat)
				worksheet.write(row,2,"Based on raw input document",entryFormat)
				row = row + 1
		elif(section == "AQFLevel"):
				value = config[section]
				Source = value["Source"]
				MinXPath = value["MinXPath"]
				MaxXPath = value["MaxXPath"]
				PreferredXPath = value["PreferredXPath"]
				RequiredXPath = value["RequiredXPath"]
				worksheet.write_url(row,0,  'internal:AQFLevel!A1',hyperlink,"MinXPath")
				worksheet.write(row,1,MinXPath,entryFormat)
				worksheet.write(row,2,Source,entryFormat)
				row = row + 1
				worksheet.write_url(row,0,  'internal:AQFLevel!A1',hyperlink,"MaxXPath")
				worksheet.write(row,1,MaxXPath,entryFormat)
				worksheet.write(row,2,Source,entryFormat)
				row = row + 1
				worksheet.write_url(row,0,  'internal:AQFLevel!A1',hyperlink,"PreferredXPath")
				worksheet.write(row,1,PreferredXPath,entryFormat)
				worksheet.write(row,2,Source,entryFormat)
				row = row + 1
				worksheet.write_url(row,0,  'internal:AQFLevel!A1',hyperlink,"RequiredXPath")
				worksheet.write(row,1,RequiredXPath,entryFormat)
				worksheet.write(row,2,Source,entryFormat)
				row = row + 1
		elif(section == "LocalGovt"):
				value = config[section]
				Source = value["Source"]
				XPath = value["XPath"]
				worksheet.write_url(row,0,  'internal:LocalGovt!A1',hyperlink,"LocalGovt")
				worksheet.write(row,1,XPath,entryFormat)
				worksheet.write(row,2,Source,entryFormat)
				row = row + 1
				worksheet.write(row,0,"Certification",entryFormat)
				worksheet.write(row,1,"/JobDoc/DataElementsRollup/Certification",entryFormat)
				worksheet.write(row,2,"Based on raw input document",entryFormat)
				row = row + 1
		elif(section == "ANZSCO"):
				value = config[section]
				Source = value["Source"]
				XPath = value["XPath"]
				worksheet.write_url(row,0,  'internal:ANZSCO!A1',hyperlink,"ANZSCO")
				worksheet.write(row,1,XPath,entryFormat)
				worksheet.write(row,2,Source,entryFormat)
				row = row + 1
				worksheet.write(row,0,"Certification",entryFormat)
				worksheet.write(row,1,"/JobDoc/DataElementsRollup/Certification",entryFormat)
				worksheet.write(row,2,"Based on raw input document",entryFormat)
				row = row + 1
		elif(section == "ConsolidatedUKSICRuleid"):
				value = config[section]
				Source = value["Source"]
				XPath = value["XPath"]
				worksheet.write_url(row,0,  'internal:ConsolidatedUKSICRuleid!A1',hyperlink,"ConsolidatedUKSICRuleid")
				worksheet.write(row,1,XPath,entryFormat)
				worksheet.write(row,2,Source,entryFormat)
				row = row + 1
				worksheet.write(row,0,"InternshipFlag",entryFormat)
				worksheet.write(row,1,"/JobDoc/DataElementsRollup/InternshipFlag",entryFormat)
				worksheet.write(row,2,"Boolean (True or False)",entryFormat)
				row = row + 1
				worksheet.write(row,0,"Salary",entryFormat)
				worksheet.write(row,1,"/JobDoc/DataElementsRollup/Salary",entryFormat)
				worksheet.write(row,2,"Based on raw input document",entryFormat)
				row = row + 1
		elif(section == "ConsolidatedNAICSRuleid"):
				value = config[section]
				Source = value["Source"]
				XPath = value["XPath"]
				worksheet.write_url(row,0,  'internal:ConsolidatedNAICSRuleid!A1',hyperlink,"ConsolidatedNAICSRuleid")
				worksheet.write(row,1,XPath,entryFormat)
				worksheet.write(row,2,Source,entryFormat)
				row = row + 1
				worksheet.write(row,0,"InternshipFlag",entryFormat)
				worksheet.write(row,1,"/JobDoc/DataElementsRollup/InternshipFlag",entryFormat)
				worksheet.write(row,2,"Boolean (True or False)",entryFormat)
				row = row + 1
				worksheet.write(row,0,"Salary",entryFormat)
				worksheet.write(row,1,"/JobDoc/DataElementsRollup/Salary",entryFormat)
				worksheet.write(row,2,"Based on raw input document",entryFormat)
				row = row + 1
		elif(section == "RootTitle"):
				value = config[section]
				Source = value["Source"]
				XPath = value["XPath"]
				worksheet.write_url(row,0,  'internal:RootTitle!A1',hyperlink,"RootTitle")
				worksheet.write(row,1,XPath,entryFormat)
				worksheet.write(row,2,Source,entryFormat)
				row = row + 1
				worksheet.write(row,0,"EMail",entryFormat)
				worksheet.write(row,1,"/JobDoc/DataElementsRollup/EMail",entryFormat)
				worksheet.write(row,2,"Based on raw input document",entryFormat)
				row = row + 1
				worksheet.write(row,0,"JobUrl",entryFormat)
				worksheet.write(row,1,"/JobDoc/DataElementsRollup/JobUrl",entryFormat)
				worksheet.write(row,2,"Based on raw input document",entryFormat)
				row = row + 1
				worksheet.write(row,0,"JobDomain",entryFormat)
				worksheet.write(row,1,"/JobDoc/DataElementsRollup/JobDomain",entryFormat)
				worksheet.write(row,2,"Based on raw input document",entryFormat)
				row = row + 1
				worksheet.write(row,0,"JobTitleRule",entryFormat)
				worksheet.write(row,1,"/JobDoc/DataElementsRollup/JobTitleRule",entryFormat)
				worksheet.write(row,2,"",entryFormat)
				row = row + 1
				worksheet.write(row,0,"EmployerRule",entryFormat)
				worksheet.write(row,1,"/JobDoc/DataElementsRollup/EmployerRule",entryFormat)
				worksheet.write(row,2,"",entryFormat)
				row = row + 1
				worksheet.write(row,0,"JobReferenceId",entryFormat)
				worksheet.write(row,1,"/JobDoc/DataElementsRollup/JobReferenceId",entryFormat)
				worksheet.write(row,2,"Based on raw input document",entryFormat)
				row = row + 1
				worksheet.write(row,0,"OpeningDate",entryFormat)
				worksheet.write(row,1,"/JobDoc/DataElementsRollup/OpeningDate",entryFormat)
				worksheet.write(row,2,"Based on raw input document",entryFormat)
				row = row + 1
				worksheet.write(row,0,"CanonOpeningDate",entryFormat)
				worksheet.write(row,1,"/JobDoc/DataElementsRollup/CanonOpeningDate",entryFormat)
				worksheet.write(row,2,"Based on raw input document",entryFormat)
				row = row + 1
				worksheet.write(row,0,"ClosingDate",entryFormat)
				worksheet.write(row,1,"/JobDoc/DataElementsRollup/ClosingDate",entryFormat)
				worksheet.write(row,2,"Based on raw input document",entryFormat)
				row = row + 1
				worksheet.write(row,0,"CanonClosingDate",entryFormat)
				worksheet.write(row,1,"/JobDoc/DataElementsRollup/CanonClosingDate",entryFormat)
				worksheet.write(row,2,"Based on raw input document",entryFormat)
				row = row + 1
				worksheet.write(row,0,"Blacklist",entryFormat)
				worksheet.write(row,1,"/JobDoc/DataElementsRollup/Blacklist",entryFormat)
				worksheet.write(row,2,"Boolean (0 or 1)",entryFormat)
				row = row + 1
		elif(section == "StdMajorCIPCode"):
				value = config[section]
				Source = value["Source"]
				worksheet.write_url(row,0,  'internal:StdMajorCIPCode!A1',hyperlink,"StdMajor")
				worksheet.write(row,1,"//JobDoc/DataElementsRollup/StdMajor",entryFormat)
				worksheet.write(row,2,Source,entryFormat)
				row = row + 1
				worksheet.write_url(row,0,  'internal:StdMajorCIPCode!A2',hyperlink,"CIPCode")
				worksheet.write(row,1,"//JobDoc/DataElementsRollup/CIPCode",entryFormat)
				worksheet.write(row,2,Source,entryFormat)
				row = row + 1
		elif(section == "CanonYearsOfExperience"):
				value = config[section]
				Source = value["Source"]
				worksheet.write_url(row,0,  'internal:CanonYearsOfExperience!A1',hyperlink,"CanonYearsOfExperience-Level")
				worksheet.write(row,1,"//JobDoc/DataElementsRollup/CanonYearsOfExperience/level",entryFormat)
				worksheet.write(row,2,Source,entryFormat)
				row = row + 1
				worksheet.write_url(row,0,  'internal:CanonYearsOfExperience!A2',hyperlink,"CanonYearsOfExperience-CanonLevel")
				worksheet.write(row,1,"//JobDoc/DataElementsRollup/CanonYearsOfExperience/canonlevel",entryFormat)
				worksheet.write(row,2,Source,entryFormat)
				row = row + 1
		elif(section == "CanonSalary"):
				value = config[section]
				Source = value["Source"]
				worksheet.write_url(row,0,  'internal:CanonSalary!A1',hyperlink,"CanonSalaryISO-4217")
				worksheet.write(row,1,"//JobDoc/DataElementsRollup/CanonSalary/iso4217",entryFormat)
				worksheet.write(row,2,Source,entryFormat)
				row = row + 1
				worksheet.write(row,0,"CanonMinSalary",entryFormat)
				worksheet.write(row,1,"/JobDoc/DataElementsRollup/CanonSalary/min",entryFormat)
				worksheet.write(row,2,"Based on raw input document",entryFormat)
				row = row + 1
				worksheet.write(row,0,"CanonMaxSalary",entryFormat)
				worksheet.write(row,1,"/JobDoc/DataElementsRollup/CanonSalary/max",entryFormat)
				worksheet.write(row,2,"Based on raw input document",entryFormat)
				row = row + 1
				worksheet.write(row,0,"CanonMinAnnualSalary",entryFormat)
				worksheet.write(row,1,"/JobDoc/DataElementsRollup/CanonSalary/minannualsal",entryFormat)
				worksheet.write(row,2,"Based on raw input document",entryFormat)
				row = row + 1
				worksheet.write(row,0,"CanonMaxAnnualSalary",entryFormat)
				worksheet.write(row,1,"/JobDoc/DataElementsRollup/CanonSalary/maxannualsal",entryFormat)
				worksheet.write(row,2,"Based on raw input document",entryFormat)
				row = row + 1
				worksheet.write(row,0,"CanonMinHourlySalary",entryFormat)
				worksheet.write(row,1,"/JobDoc/DataElementsRollup/CanonSalary/minhourlysal",entryFormat)
				worksheet.write(row,2,"Based on raw input document",entryFormat)
				row = row + 1
				worksheet.write(row,0,"CanonMaxHourlySalary",entryFormat)
				worksheet.write(row,1,"/JobDoc/DataElementsRollup/CanonSalary/maxhourlysal",entryFormat)
				worksheet.write(row,2,"Based on raw input document",entryFormat)
				row = row + 1
				worksheet.write_url(row,0,  'internal:CanonSalary!A3',hyperlink,"CanonSalaryPayFrequency")
				worksheet.write(row,1,"//JobDoc/DataElementsRollup/CanonSalary/payfrequency",entryFormat)
				worksheet.write(row,2,Source,entryFormat)
				row = row + 1
				worksheet.write_url(row,0,  'internal:CanonSalary!A2',hyperlink,"CanonSalType")
				worksheet.write(row,1,"//JobDoc/DataElementsRollup/CanonSalary/saltype",entryFormat)
				worksheet.write(row,2,Source,entryFormat)
				row = row + 1
		elif(section == "CanonJobType"):
				value = config[section]
				Source = value["Source"]
				worksheet.write(row,0,"JobType",entryFormat)
				worksheet.write(row,1,"/JobDoc/DataElementsRollup/JobType",entryFormat)
				worksheet.write(row,2,"Based on raw input document",entryFormat)
				row = row + 1
				worksheet.write_url(row,0,  'internal:CanonJobType!A1',hyperlink,"CanonJobType")
				worksheet.write(row,1,"/JobDoc/DataElementsRollup/CanonJobType/type",entryFormat)
				worksheet.write(row,2,Source,entryFormat)
				row = row + 1
				worksheet.write_url(row,0,  'internal:CanonJobType!A2',hyperlink,"CanonJobHours")
				worksheet.write(row,1,"/JobDoc/DataElementsRollup/CanonJobType/hours",entryFormat)
				worksheet.write(row,2,Source,entryFormat)
				row = row + 1
				worksheet.write_url(row,0,  'internal:CanonJobType!A3',hyperlink,"CanonJobTaxTerm")
				worksheet.write(row,1,"/JobDoc/DataElementsRollup/CanonJobType/taxterm",entryFormat)
				worksheet.write(row,2,Source,entryFormat)
				row = row + 1
				worksheet.write_url(row,0,  'internal:CanonJobType!A4',hyperlink,"CanonJobWorkFromHome")
				worksheet.write(row,1,"/JobDoc/DataElementsRollup/CanonJobType/wfh",entryFormat)
				worksheet.write(row,2,"Boolean (True or False)",entryFormat)
				row = row + 1
				worksheet.write(row,0,"JobDate",entryFormat)
				worksheet.write(row,1,"/JobDoc/DataElementsRollup/JobDate",entryFormat)
				worksheet.write(row,2,"Based on raw input document",entryFormat)
				row = row + 1
		elif(section == "CanonSkilltoSkillClusterMapping"):
				value = config[section]
				Source = value["Source"]
				worksheet.write_url(row,0,  'internal:CanonSkilltoSkillClusterMapping!A1',hyperlink,"CanonSkill-(Name)")
				worksheet.write(row,1,"/JobDoc/DataElementsRollup/CanonSkills/canonskill/@name",entryFormat)
				worksheet.write(row,2,Source,entryFormat)
				row = row + 1
				worksheet.write_url(row,0,  'internal:CanonSkilltoSkillClusterMapping!A2',hyperlink,"CanonSkill-(Skill-Cluster)")
				worksheet.write(row,1,"/JobDoc/DataElementsRollup/CanonSkills/canonskill/@skill-cluster",entryFormat)
				worksheet.write(row,2,Source,entryFormat)
				row = row + 1
		elif(section == "AULocationSpecificInformation"):
				value = config[section]
				Source = value["Source"]
				XPath1 = value["XPath1"]
				XPath2 = value["XPath2"]
				XPath3 = value["XPath3"]
				XPath4 = value["XPath4"]
				XPath5 = value["XPath5"]
				XPath6 = value["XPath6"]
				XPath7 = value["XPath7"]
				XPath8 = value["XPath8"]
				XPath9 = value["XPath9"]
				XPath10 = value["XPath10"]
				XPath11 = value["XPath11"]
				XPath12 = value["XPath12"]
				XPath13 = value["XPath13"]
				XPath14 = value["XPath14"]
				worksheet.write_url(row,0,  'internal:AULocationSpecificInformation!A1',hyperlink,"LocationSpecificInformationcity")
				worksheet.write(row,1,XPath1,entryFormat)
				worksheet.write(row,2,Source,entryFormat)
				row = row + 1
				worksheet.write_url(row,0,  'internal:AULocationSpecificInformation!A1',hyperlink,"LocationSpecificInformationstate")
				worksheet.write(row,1,XPath2,entryFormat)
				worksheet.write(row,2,Source,entryFormat)
				row = row + 1
				worksheet.write_url(row,0,  'internal:AULocationSpecificInformation!A1',hyperlink,"LocationSpecificInformationcounty")
				worksheet.write(row,1,XPath3,entryFormat)
				worksheet.write(row,2,Source,entryFormat)
				row = row + 1
				worksheet.write_url(row,0,  'internal:AULocationSpecificInformation!A1',hyperlink,"LocationSpecificInformationcountry")
				worksheet.write(row,1,XPath4,entryFormat)
				worksheet.write(row,2,Source,entryFormat)
				row = row + 1
				worksheet.write_url(row,0,  'internal:AULocationSpecificInformation!A1',hyperlink,"LocationSpecificInformationzipcode")
				worksheet.write(row,1,XPath5,entryFormat)
				worksheet.write(row,2,Source,entryFormat)
				row = row + 1
				worksheet.write_url(row,0,  'internal:AULocationSpecificInformation!A1',hyperlink,"LocationSpecificInformationgeocodelat")
				worksheet.write(row,1,XPath6,entryFormat)
				worksheet.write(row,2,Source,entryFormat)
				row = row + 1
				worksheet.write_url(row,0,  'internal:AULocationSpecificInformation!A1',hyperlink,"LocationSpecificInformationgeocodelon")
				worksheet.write(row,1,XPath7,entryFormat)
				worksheet.write(row,2,Source,entryFormat)
				row = row + 1
				worksheet.write_url(row,0,  'internal:AULocationSpecificInformation!A1',hyperlink,"LocationSpecificInformationsacodes(sa2code)")
				worksheet.write(row,1,XPath8,entryFormat)
				worksheet.write(row,2,Source,entryFormat)
				row = row + 1
				worksheet.write_url(row,0,  'internal:AULocationSpecificInformation!A1',hyperlink,"LocationSpecificInformationsacodes(sa3code)")
				worksheet.write(row,1,XPath9,entryFormat)
				worksheet.write(row,2,Source,entryFormat)
				row = row + 1
				worksheet.write_url(row,0,  'internal:AULocationSpecificInformation!A1',hyperlink,"LocationSpecificInformationsacodes(sa4code)")
				worksheet.write(row,1,XPath10,entryFormat)
				worksheet.write(row,2,Source,entryFormat)
				row = row + 1
				worksheet.write_url(row,0,  'internal:AULocationSpecificInformation!A1',hyperlink,"LocationSpecificInformationsrsregion")
				worksheet.write(row,1,XPath11,entryFormat)
				worksheet.write(row,2,Source,entryFormat)
				row = row + 1
				worksheet.write_url(row,0,  'internal:AULocationSpecificInformation!A1',hyperlink,"LocationSpecificInformationslacode")
				worksheet.write(row,1,XPath12,entryFormat)
				worksheet.write(row,2,Source,entryFormat)
				row = row + 1
				worksheet.write_url(row,0,  'internal:AULocationSpecificInformation!A1',hyperlink,"LocationSpecificInformationlmr")
				worksheet.write(row,1,XPath13,entryFormat)
				worksheet.write(row,2,Source,entryFormat)
				row = row + 1
				worksheet.write_url(row,0,  'internal:AULocationSpecificInformation!A1',hyperlink,"LocationSpecificInformationphoneareacode")
				worksheet.write(row,1,XPath14,entryFormat)
				worksheet.write(row,2,Source,entryFormat)
				row = row + 1
		elif(section == "NZLocationSpecificInformation"):
				continue
		elif(section == "USLocationSpecificInformation"):
				value = config[section]
				Source = value["Source"]
				source = str(Source)
				#worksheet.write(row,0,section)
				worksheet.write_url(row,0,  'internal:USLocationSpecificInformation!A1',hyperlink,"LocationSpecificInformationcity")
				worksheet.write(row,1,"/JobDoc/DataElementsRollup/LocationSpecificInformation/city",entryFormat)
				worksheet.write(row,2,Source,entryFormat)
				row = row + 1
				worksheet.write_url(row,0,  'internal:USLocationSpecificInformation!A1',hyperlink,"LocationSpecificInformationstate")
				worksheet.write(row,1,"/JobDoc/DataElementsRollup/LocationSpecificInformation/state",entryFormat)
				worksheet.write(row,2,Source,entryFormat)
				row = row + 1
				worksheet.write_url(row,0,  'internal:USLocationSpecificInformation!A1',hyperlink,"LocationSpecificInformationcounty")
				worksheet.write(row,1,"/JobDoc/DataElementsRollup/LocationSpecificInformation/county",entryFormat)
				worksheet.write(row,2,Source,entryFormat)
				row = row + 1
				worksheet.write_url(row,0,  'internal:USLocationSpecificInformation!A1',hyperlink,"LocationSpecificInformationcountry")
				worksheet.write(row,1,"/JobDoc/DataElementsRollup/LocationSpecificInformation/country",entryFormat)
				worksheet.write(row,2,Source,entryFormat)
				row = row + 1
				worksheet.write_url(row,0,  'internal:USLocationSpecificInformation!A1',hyperlink,"LocationSpecificInformationzipcode")
				worksheet.write(row,1,"/JobDoc/DataElementsRollup/LocationSpecificInformation/zipcode",entryFormat)
				worksheet.write(row,2,Source,entryFormat)
				row = row + 1
				worksheet.write_url(row,0,  'internal:USLocationSpecificInformation!A1',hyperlink,"LocationSpecificInformationgeocodelat")
				worksheet.write(row,1,"/JobDoc/DataElementsRollup/LocationSpecificInformation/geocode/lat",entryFormat)
				worksheet.write(row,2,Source,entryFormat)
				row = row + 1
				worksheet.write_url(row,0,  'internal:USLocationSpecificInformation!A1',hyperlink,"LocationSpecificInformationgeocodelon")
				worksheet.write(row,1,"/JobDoc/DataElementsRollup/LocationSpecificInformation/geocode/lon",entryFormat)
				worksheet.write(row,2,Source,entryFormat)
				row = row + 1
				worksheet.write_url(row,0,  'internal:USLocationSpecificInformation!A1',hyperlink,"MSA")
				worksheet.write(row,1,"/JobDoc/DataElementsRollup/LocationSpecificInformation/msa",entryFormat)
				worksheet.write(row,2,Source,entryFormat)
				row = row + 1
				worksheet.write_url(row,0,  'internal:USLocationSpecificInformation!A1',hyperlink,"LMA")
				worksheet.write(row,1,"/JobDoc/DataElementsRollup/LocationSpecificInformation/lma",entryFormat)
				worksheet.write(row,2,Source,entryFormat)
				row = row + 1
				worksheet.write_url(row,0,  'internal:USLocationSpecificInformation!A1',hyperlink,"PhoneAreaCode")
				worksheet.write(row,1,"/JobDoc/DataElementsRollup/LocationSpecificInformation/phoneareacode",entryFormat)
				worksheet.write(row,2,Source,entryFormat)
				row = row + 1
				worksheet.write_url(row,0,  'internal:USLocationSpecificInformation!A1',hyperlink,"DivisionCode")
				worksheet.write(row,1,"/JobDoc/DataElementsRollup/LocationSpecificInformation/divisioncode",entryFormat)
				worksheet.write(row,2,Source,entryFormat)
				row = row + 1
				worksheet.write(row,0,"Telephone",entryFormat)
				worksheet.write(row,1,"/JobDoc/DataElementsRollup/Telephone",entryFormat)
				worksheet.write(row,2,"Based on raw document",entryFormat)
				row = row + 1
		elif(section == "SGPLocationSpecificInformation"):
				value = config[section]
				Source = value["Source"]
				source = str(Source)
				#worksheet.write(row,0,section)
				worksheet.write_url(row,0,  'internal:SGPLocationSpecificInformation!A1',hyperlink,"LocationSpecificInformationcity")
				worksheet.write(row,1,"/JobDoc/DataElementsRollup/LocationSpecificInformation/city",entryFormat)
				worksheet.write(row,2,Source,entryFormat)
				row = row + 1
				worksheet.write_url(row,0,  'internal:SGPLocationSpecificInformation!A1',hyperlink,"LocationSpecificInformationstate")
				worksheet.write(row,1,"/JobDoc/DataElementsRollup/LocationSpecificInformation/state",entryFormat)
				worksheet.write(row,2,Source,entryFormat)
				row = row + 1
				worksheet.write_url(row,0,  'internal:SGPLocationSpecificInformation!A1',hyperlink,"LocationSpecificInformationcounty")
				worksheet.write(row,1,"/JobDoc/DataElementsRollup/LocationSpecificInformation/county",entryFormat)
				worksheet.write(row,2,Source,entryFormat)
				row = row + 1
				worksheet.write_url(row,0,  'internal:SGPLocationSpecificInformation!A1',hyperlink,"LocationSpecificInformationcountry")
				worksheet.write(row,1,"/JobDoc/DataElementsRollup/LocationSpecificInformation/country",entryFormat)
				worksheet.write(row,2,Source,entryFormat)
				row = row + 1
				worksheet.write_url(row,0,  'internal:SGPLocationSpecificInformation!A1',hyperlink,"LocationSpecificInformationzipcode")
				worksheet.write(row,1,"/JobDoc/DataElementsRollup/LocationSpecificInformation/zipcode",entryFormat)
				worksheet.write(row,2,Source,entryFormat)
				row = row + 1
				worksheet.write_url(row,0,  'internal:SGPLocationSpecificInformation!A1',hyperlink,"LocationSpecificInformationgeocodelat")
				worksheet.write(row,1,"/JobDoc/DataElementsRollup/LocationSpecificInformation/geocode/lat",entryFormat)
				worksheet.write(row,2,Source,entryFormat)
				row = row + 1
				worksheet.write_url(row,0,  'internal:SGPLocationSpecificInformation!A1',hyperlink,"LocationSpecificInformationgeocodelon")
				worksheet.write(row,1,"/JobDoc/DataElementsRollup/LocationSpecificInformation/geocode/lon",entryFormat)
				worksheet.write(row,2,Source,entryFormat)
				row = row + 1
				worksheet.write_url(row,0,  'internal:SGPLocationSpecificInformation!A1',hyperlink,"MSA")
				worksheet.write(row,1,"/JobDoc/DataElementsRollup/LocationSpecificInformation/msa",entryFormat)
				worksheet.write(row,2,Source,entryFormat)
				row = row + 1
				worksheet.write_url(row,0,  'internal:SGPLocationSpecificInformation!A1',hyperlink,"LMA")
				worksheet.write(row,1,"/JobDoc/DataElementsRollup/LocationSpecificInformation/lma",entryFormat)
				worksheet.write(row,2,Source,entryFormat)
				row = row + 1
				worksheet.write_url(row,0,  'internal:SGPLocationSpecificInformation!A1',hyperlink,"PhoneAreaCode")
				worksheet.write(row,1,"/JobDoc/DataElementsRollup/LocationSpecificInformation/phoneareacode",entryFormat)
				worksheet.write(row,2,Source,entryFormat)
				row = row + 1
				worksheet.write_url(row,0,  'internal:SGPLocationSpecificInformation!A1',hyperlink,"DivisionCode")
				worksheet.write(row,1,"/JobDoc/DataElementsRollup/LocationSpecificInformation/divisioncode",entryFormat)
				worksheet.write(row,2,Source,entryFormat)
				row = row + 1
				worksheet.write(row,0,"Telephone",entryFormat)
				worksheet.write(row,1,"/JobDoc/DataElementsRollup/Telephone",entryFormat)
				worksheet.write(row,2,"Based on raw document",entryFormat)
				row = row + 1
		elif(section == "LocationSpecificInformation"):
				value = config[section]
				Source = value["Source"]
				worksheet.write_url(row,0,  'internal:LocationSpecificInformation!A1',hyperlink,"CanonCity")
				worksheet.write(row,1,"/JobDoc/DataElementsRollup/LocationSpecificInformation/city",entryFormat)
				worksheet.write(row,2,Source,entryFormat)
				row = row + 1
				worksheet.write_url(row,0,  'internal:LocationSpecificInformation!A2',hyperlink,"CanonZipCode")
				worksheet.write(row,1,"/JobDoc/DataElementsRollup/LocationSpecificInformation/zipcode",entryFormat)
				worksheet.write(row,2,Source,entryFormat)
				row = row + 1
				worksheet.write_url(row,0,  'internal:LocationSpecificInformation!A3',hyperlink,"TravelToWorkArea")
				worksheet.write(row,1,"/JobDoc/DataElementsRollup/LocationSpecificInformation/traveltoworkarea",entryFormat)
				worksheet.write(row,2,Source,entryFormat)
				row = row + 1
				worksheet.write_url(row,0,  'internal:LocationSpecificInformation!A4',hyperlink,"LocalAuthorityDistrict")
				worksheet.write(row,1,"/JobDoc/DataElementsRollup/LocationSpecificInformation/localauthoritydistrict",entryFormat)
				worksheet.write(row,2,Source,entryFormat)
				row = row + 1
				worksheet.write_url(row,0,  'internal:LocationSpecificInformation!A5',hyperlink,"CanonCounty")
				worksheet.write(row,1,"/JobDoc/DataElementsRollup/LocationSpecificInformation/county",entryFormat)
				worksheet.write(row,2,Source,entryFormat)
				row = row + 1
				worksheet.write_url(row,0,  'internal:LocationSpecificInformation!A6',hyperlink,"LocalEnterprisePartnership")
				worksheet.write(row,1,"/JobDoc/DataElementsRollup/LocationSpecificInformation/localenterprisepartnership",entryFormat)
				worksheet.write(row,2,Source,entryFormat)
				row = row + 1
				worksheet.write_url(row,0,  'internal:LocationSpecificInformation!A7',hyperlink,"Region")
				worksheet.write(row,1,"/JobDoc/DataElementsRollup/LocationSpecificInformation/region",entryFormat)
				worksheet.write(row,2,Source,entryFormat)
				row = row + 1
				worksheet.write_url(row,0,  'internal:LocationSpecificInformation!A8',hyperlink,"EnglishCountry")
				worksheet.write(row,1,"/JobDoc/DataElementsRollup/LocationSpecificInformation/englishcountry",entryFormat)
				worksheet.write(row,2,Source,entryFormat)
				row = row + 1
				worksheet.write_url(row,0,  'internal:LocationSpecificInformation!A9',hyperlink,"CanonLat")
				worksheet.write(row,1,"/JobDoc/DataElementsRollup/LocationSpecificInformation/geocode/lat",entryFormat)
				worksheet.write(row,2,Source,entryFormat)
				row = row + 1
				worksheet.write_url(row,0,  'internal:LocationSpecificInformation!A10',hyperlink,"CanonLon")
				worksheet.write(row,1,"/JobDoc/DataElementsRollup/LocationSpecificInformation/geocode/lon",entryFormat)
				worksheet.write(row,2,Source,entryFormat)
				row = row + 1
				worksheet.write_url(row,0,  'internal:LocationSpecificInformation!A11',hyperlink,"LowerSuperOutputArea")
				worksheet.write(row,1,"/JobDoc/DataElementsRollup/LocationSpecificInformation/lsoa",entryFormat)
				worksheet.write(row,2,Source,entryFormat)
				row = row + 1
				worksheet.write_url(row,0,  'internal:LocationSpecificInformation!A12',hyperlink,"LocationSpecificInformation-Ruleid")
				worksheet.write(row,1,"/JobDoc/DataElementsRollup/LocationSpecificInformation/@ruleid",entryFormat)
				worksheet.write(row,2,Source,entryFormat)
				row = row + 1
				worksheet.write(row,0,"Telephone",entryFormat)
				worksheet.write(row,1,"/JobDoc/DataElementsRollup/Telephone",entryFormat)
				worksheet.write(row,2,"Based on raw input document",entryFormat)
				row = row + 1
		elif(section == "BGTOccSGP"):
				value = config[section]
				Source = value["Source"]
				worksheet.write_url(row,0,  'internal:BGTOcc!A1',hyperlink,"BGTOcc")
				worksheet.write(row,1,"/JobDoc/DataElementsRollup/BGTOcc",entryFormat)
				worksheet.write(row,2,Source,entryFormat)
				row = row + 1
				
		else:
				#print "------------------------"
				#print section  
				value = config[section]
				xpath=value["XPath"]
				Source = value["Source"]
				source = str(Source)
				#worksheet.write(row,0,section)
				
				worksheet.write_url(row,0,  'internal:%s!C1'%(section),hyperlink,section)
				worksheet.write(row,1,xpath,entryFormat)
				worksheet.write(row,2,source,entryFormat)
				row = row + 1




				
for section in (config.keys()):
		if(section == "Output"):
				continue
		else:
				value=config[section]
				Presence = value["Presence"]
				
				if Presence in ["Yes","yes"]:
						#for Summary Sheet
						#sectionlist.append(section)
						
								
						if(section == "MajorCode"): #for MajorCode module 
								value=config[section]
								XPath = value["XPath"]
								Source = value["Source"]
								source = str(Source)
								OutputColumnHeader=value["OutputColumnHeader"]
								colHead=OutputColumnHeader.split(';')
								worksheet=workbook.add_worksheet(section)
								worksheet.write(0,0,section+" XPath",format)
								worksheet.write(0,1,XPath,entryFormat)
								worksheet.write_url(0,9,  'internal:Summary!A1',hyperlink,"Back To Summary Sheet")
								if(source == "Based on Lookup and raw input document"):
										worksheet.write(0,13,"Note: This is the possible MajorCodes based on the lookup. In addition to this, there will be other Majorcode values emitted based on input raw tags",format1)
								worksheet.write(2,0,colHead[0],format)
								worksheet.write(2,1,colHead[1],format)
								row = 3
								for val in sorted(uniq_dict_Major.keys()):
										key1 = val.lstrip(' ')
										key1 = val.rstrip('\n')
										worksheet.write_string(row,0,key1,entryFormat)
										worksheet.write_string(row,1,uniq_dict_Major[val],entryFormat)
										row = row + 1
						elif( section == "AULocationSpecificInformation"):
								value = config[section]
								Source = value["Source"]
								XPath1 = value["XPath1"]
								XPath2 = value["XPath2"]
								XPath3 = value["XPath3"]
								XPath4 = value["XPath4"]
								XPath5 = value["XPath5"]
								XPath6 = value["XPath6"]
								XPath7 = value["XPath7"]
								XPath8 = value["XPath8"]
								XPath9 = value["XPath9"]
								XPath10 = value["XPath10"]
								XPath11 = value["XPath11"]
								XPath12 = value["XPath12"]
								XPath13 = value["XPath13"]
								XPath14 = value["XPath14"]
								OutputColumnHeader=value["OutputColumnHeader"]
								colHead=OutputColumnHeader.split(';')
								worksheet=workbook.add_worksheet(section)

								worksheet.write(0,0,"LocationSpecificInformation tags",format)
								worksheet.write(0,1,"XPath",format)
								
								worksheet.write(1,0,"LocationSpecificInformationcity",format)
								worksheet.write(1,1,XPath1,entryFormat)
								worksheet.write_url(0,9,  'internal:Summary!A1',hyperlink,"Back To Summary Sheet")
								if(Source == "Based on Lookup and raw input document"):
										worksheet.write(0,13,"Note: This is the possible AULocation Specific Information values based on the lookup. In addition to this, there will be other AULocation Specific Information values emitted based on input raw tags",format1)
								worksheet.write(2,0,"LocationSpecificInformationstate",format)
								worksheet.write(2,1,XPath2,entryFormat)
								worksheet.write(3,0,"LocationSpecificInformationcounty",format)
								worksheet.write(3,1,XPath3,entryFormat)
								worksheet.write(4,0,"LocationSpecificInformationcountry",format)
								worksheet.write(4,1,XPath4,entryFormat)
								worksheet.write(5,0,"LocationSpecificInformationzipcode",format)
								worksheet.write(5,1,XPath5,entryFormat)
								worksheet.write(6,0,"LocationSpecificInformationgeocodelat",format)
								worksheet.write(6,1,XPath6,entryFormat)
								worksheet.write(7,0,"LocationSpecificInformationgeocodelon",format)
								worksheet.write(7,1,XPath7,entryFormat)
								worksheet.write(8,0,"LocationSpecificInformationsacodes(sa2code)",format)
								worksheet.write(8,1,XPath8,entryFormat)
								worksheet.write(9,0,"LocationSpecificInformationsacodes(sa3code)",format)
								worksheet.write(9,1,XPath9,entryFormat)
								worksheet.write(10,0,"LocationSpecificInformationsacodes(sa4code)",format)
								worksheet.write(10,1,XPath10,entryFormat)
								worksheet.write(11,0,"LocationSpecificInformationsrsregion",format)
								worksheet.write(11,1,XPath11,entryFormat)
								worksheet.write(12,0,"LocationSpecificInformationslacode",format)
								worksheet.write(12,1,XPath12,entryFormat)
								worksheet.write(13,0,"LocationSpecificInformationlmr",format)
								worksheet.write(13,1,XPath13,entryFormat)
								worksheet.write(14,0,"LocationSpecificInformationphoneareacode",format)
								worksheet.write(14,1,XPath14,entryFormat)
								column = 0
								for val in colHead:
										worksheet.write(16,column,val,format)
										column = column+1
								row = 17
								for key in sorted(dictSA.keys()):
										listofdata = dictSA.get(key)
										if key in dictLL.keys():
												latlist = dictLL.get(key)
												worksheet.write(row,3,latlist[0],entryFormat)
												worksheet.write(row,4,latlist[1],entryFormat)
										else:
												worksheet.write(row,3,"NA",entryFormat)
												worksheet.write(row,4,"NA",entryFormat)
										
										worksheet.write(row,0,listofdata[0],entryFormat)
										
										worksheet.write(row,1,listofdata[1],entryFormat)
										worksheet.write(row,2,key,entryFormat)
										worksheet.write(row,5,listofdata[2],entryFormat)
										worksheet.write(row,6,listofdata[3],entryFormat)
										worksheet.write(row,7,listofdata[4],entryFormat)
										if key in dictLMR.keys():
												lmrlist = dictLMR.get(key)
												worksheet.write(row,8,lmrlist[0],entryFormat)
												worksheet.write(row,9,lmrlist[1],entryFormat)
												worksheet.write(row,10,lmrlist[2],entryFormat)
												worksheet.write(row,11,lmrlist[3],entryFormat)
												worksheet.write(row,12,lmrlist[4],entryFormat)
										else:
												worksheet.write(row,8,"NA",entryFormat)
												worksheet.write(row,9,"NA",entryFormat)
												worksheet.write(row,10,"NA",entryFormat)
												worksheet.write(row,11,"NA",entryFormat)
												worksheet.write(row,12,"NA",entryFormat)
										row = row+1
								for key in dictRE.keys():
										listRE = dictRE.get(key)
										worksheet.write(row,0,key,entryFormat)
										worksheet.write(row,1,listRE[0],entryFormat)
										worksheet.write(row,2,"NA",entryFormat)
										worksheet.write(row,3,"NA",entryFormat)
										worksheet.write(row,4,"NA",entryFormat)
										worksheet.write(row,5,"NA",entryFormat)
										worksheet.write(row,6,"NA",entryFormat)
										worksheet.write(row,7,"NA",entryFormat)
										worksheet.write(row,8,listRE[1],entryFormat)
										worksheet.write(row,9,listRE[2],entryFormat)
										worksheet.write(row,10,listRE[3],entryFormat)
										if key in ["NA","na"]:
												worksheet.write(row,11,"NA",entryFormat)
												worksheet.write(row,12,"NA",entryFormat)
										else:
												worksheet.write(row,11,"AUS",entryFormat)
												worksheet.write(row,12,"000",entryFormat)
										
										row = row + 1
						elif( section == "NZLocationSpecificInformation"):
								value = config[section]
								Source = value["Source"]
								XPath1 = value["XPath1"]
								XPath2 = value["XPath2"]
								XPath3 = value["XPath3"]
								XPath4 = value["XPath4"]
								XPath5 = value["XPath5"]
								XPath6 = value["XPath6"]
								XPath7 = value["XPath7"]
								XPath8 = value["XPath8"]
								XPath9 = value["XPath9"]
								XPath10 = value["XPath10"]
								XPath11 = value["XPath11"]
								XPath12 = value["XPath12"]
								XPath13 = value["XPath13"]
								XPath14 = value["XPath14"]
								OutputColumnHeader=value["OutputColumnHeader"]
								colHead=OutputColumnHeader.split(';')
								worksheet=workbook.add_worksheet(section)

								worksheet.write(0,0,"LocationSpecificInformation tags",format)
								worksheet.write(0,1,"XPath",format)
								
								worksheet.write(1,0,"LocationSpecificInformationcity",format)
								worksheet.write(1,1,XPath1,entryFormat)
								worksheet.write_url(0,9,  'internal:Summary!A1',hyperlink,"Back To Summary Sheet")
								if(Source == "Based on Lookup and raw input document"):
										worksheet.write(0,13,"Note: This is the possible AULocation Specific Information values based on the lookup. In addition to this, there will be other AULocation Specific Information values emitted based on input raw tags",format1)
								worksheet.write(2,0,"LocationSpecificInformationstate",format)
								worksheet.write(2,1,XPath2,entryFormat)
								worksheet.write(3,0,"LocationSpecificInformationcounty",format)
								worksheet.write(3,1,XPath3,entryFormat)
								worksheet.write(4,0,"LocationSpecificInformationcountry",format)
								worksheet.write(4,1,XPath4,entryFormat)
								worksheet.write(5,0,"LocationSpecificInformationzipcode",format)
								worksheet.write(5,1,XPath5,entryFormat)
								worksheet.write(6,0,"LocationSpecificInformationgeocodelat",format)
								worksheet.write(6,1,XPath6,entryFormat)
								worksheet.write(7,0,"LocationSpecificInformationgeocodelon",format)
								worksheet.write(7,1,XPath7,entryFormat)
								worksheet.write(8,0,"LocationSpecificInformationsacodes(sa2code)",format)
								worksheet.write(8,1,XPath8,entryFormat)
								worksheet.write(9,0,"LocationSpecificInformationsacodes(sa3code)",format)
								worksheet.write(9,1,XPath9,entryFormat)
								worksheet.write(10,0,"LocationSpecificInformationsacodes(sa4code)",format)
								worksheet.write(10,1,XPath10,entryFormat)
								worksheet.write(11,0,"LocationSpecificInformationsrsregion",format)
								worksheet.write(11,1,XPath11,entryFormat)
								worksheet.write(12,0,"LocationSpecificInformationslacode",format)
								worksheet.write(12,1,XPath12,entryFormat)
								worksheet.write(13,0,"LocationSpecificInformationlmr",format)
								worksheet.write(13,1,XPath13,entryFormat)
								worksheet.write(14,0,"LocationSpecificInformationphoneareacode",format)
								worksheet.write(14,1,XPath14,entryFormat)
								column = 0
								for val in colHead:
										worksheet.write(16,column,val,format)
										column = column+1
								row = 17
								for key in sorted(nzdictSA.keys()):
										listofdata = nzdictSA.get(key)
										if key in nzdictLL.keys():
												latlist = nzdictLL.get(key)
												worksheet.write(row,3,latlist[0],entryFormat)
												worksheet.write(row,4,latlist[1],entryFormat)
										else:
												worksheet.write(row,3,"NA",entryFormat)
												worksheet.write(row,4,"NA",entryFormat)
										
										worksheet.write(row,0,listofdata[0],entryFormat)
										#print listofdata[1]
										worksheet.write(row,1,listofdata[1],entryFormat)
										worksheet.write(row,2,key,entryFormat)
										worksheet.write(row,5,listofdata[2],entryFormat)
										worksheet.write(row,6,listofdata[3],entryFormat)
										worksheet.write(row,7,listofdata[4],entryFormat)
										if key in nzdictLMR.keys():
												lmrlist = nzdictLMR.get(key)
												worksheet.write(row,8,lmrlist[0],entryFormat)
												worksheet.write(row,9,lmrlist[1],entryFormat)
												worksheet.write(row,10,lmrlist[2],entryFormat)
												worksheet.write(row,11,lmrlist[3],entryFormat)
												worksheet.write(row,12,lmrlist[4],entryFormat)
										else:
												worksheet.write(row,8,"NA",entryFormat)
												worksheet.write(row,9,"NA",entryFormat)
												worksheet.write(row,10,"NA",entryFormat)
												worksheet.write(row,11,"NA",entryFormat)
												worksheet.write(row,12,"NA",entryFormat)
										row = row+1
								for key in nzdictRE.keys():
										listRE = nzdictRE.get(key)
										worksheet.write(row,0,key,entryFormat)
										#print listRE[0]
										worksheet.write(row,1,listRE[0],entryFormat)
										worksheet.write(row,2,"NA",entryFormat)
										worksheet.write(row,3,"NA",entryFormat)
										worksheet.write(row,4,"NA",entryFormat)
										worksheet.write(row,5,"NA",entryFormat)
										worksheet.write(row,6,"NA",entryFormat)
										worksheet.write(row,7,"NA",entryFormat)
										worksheet.write(row,8,listRE[1],entryFormat)
										worksheet.write(row,9,listRE[2],entryFormat)
										worksheet.write(row,10,listRE[3],entryFormat)
										if key in ["NA","na"]:
												worksheet.write(row,11,"NA",entryFormat)
												worksheet.write(row,12,"NA",entryFormat)
										else:
												worksheet.write(row,11,"NZL",entryFormat)
												worksheet.write(row,12,"000",entryFormat)
										
										row = row + 1
						elif section in ["USLocationSpecificInformation","SGPLocationSpecificInformation"]: #for USLocation module and for SGP Location
								value=config[section]
								Source = value["Source"]
								source = str(Source)
								XPath = value["XPath"]
								OutputColumnHeader=value["OutputColumnHeader"]
								colHead=OutputColumnHeader.split(';')
								worksheet=workbook.add_worksheet(section)
								worksheet.write(0,0,"LocationSpecificInformation XPath",format)
								worksheet.write(0,1,XPath,entryFormat)
								worksheet.write_url(0,9,  'internal:Summary!A1',hyperlink,"Back To Summary Sheet")
								if(source == "Based on Lookup and raw input document"):
										worksheet.write(0,13,"Note: This is the possible Location Specific Information values based on the lookup. In addition to this, there will be other Location Specific Information values emitted based on input raw tags",format1)
								column = 0
								for val in colHead:
										worksheet.write(2,column,val,format)
										column = column+1
								row = 3
								for key in sorted(dictcity.keys()):
										if key not in ["NA","na"]:
												listofdata=dictcity.get(key)
												worksheet.write(row,0,key,entryFormat)
												if key in dictmsa.keys():
														listcity =dictmsa.get(key)
														worksheet.write(row,7,listcity[0],entryFormat)
														worksheet.write(row,8,listcity[1],entryFormat)
												elif key in dictlma.keys():
														listcity =dictlma.get(key)
														worksheet.write(row,9,listcity[0],entryFormat)
												col = 1
												for num in range(len(listofdata)):
														if(num == 1):
																worksheet.write(row,2,listofdata[num],entryFormat)
																if listofdata[num] in dictzips.keys():
																		keylist = dictzips.get(listofdata[num])
																		worksheet.write(row,5,keylist[0],entryFormat)
																		worksheet.write(row,6,keylist[1],entryFormat)
														elif(num == 2):
																worksheet.write(row,3,listofdata[num],entryFormat)
																
														elif(num == 3):
																worksheet.write(row,10,listofdata[num],entryFormat)
														elif(num == 4):
																worksheet.write(row,num,listofdata[num],entryFormat)
																break
														else:
																worksheet.write(row,col,listofdata[num],entryFormat)
																col = col+1
										row = row + 1
								for key in dictcounty.keys():
										worksheet.write(row,3,key,entryFormat)
										if key in dictmsa.keys():
														listcity =dictmsa.get(key)
														worksheet.write(row,7,listcity[0],entryFormat)
														worksheet.write(row,8,listcity[1],entryFormat)
										elif key in dictlma.keys():
														listcity =dictlma.get(key)
														worksheet.write(row,9,listcity[0],entryFormat)
										col = 0
										listofdata = dictcounty.get(key)
										for num in range(len(listofdata)):
												if(num == 0):
														worksheet.write(row,0,listofdata[num],entryFormat)
														if listofdata[num] in dictlocations.keys():
																keylist = dictlocations.get(listofdata[num])
																worksheet.write(row,5,keylist[1],entryFormat)
																worksheet.write(row,6,keylist[0],entryFormat)
														elif listofdata[num] in dictcity.keys():
																keylist = dictcity.get(listofdata[num])
																worksheet.write(row,2,keylist[1],entryFormat)
												elif(num == 2):
														worksheet.write(row,10,listofdata[num],entryFormat)
												else:
														worksheet.write(row,col,listofdata[num],entryFormat)
														col = col+1
										row = row+1  
						elif( section == "LocationSpecificInformation"): #for Location module
								value=config[section]
								Source = value["Source"]
								source = str(Source)
								CanonCity_XPath = value["CanonCity XPath"]
								CanonZipCode_XPath = value["CanonZipCode XPath"]
								TravelToWorkArea_XPath = value["TravelToWorkArea XPath"]
								LocalAuthorityDistrict_XPath = value["LocalAuthorityDistrict XPath"]
								CanonCounty_XPath = value["CanonCounty XPath"]
								LocalEnterprisePartnership_XPath = value["LocalEnterprisePartnership XPath"]
								Region_XPath = value["Region XPath"]
								EnglishCountry_XPath = value["EnglishCountry XPath"]
								CanonLat_XPath = value["CanonLat XPath"]
								CanonLon_XPath = value["CanonLon XPath"]
								LowerSuperOutputArea_XPath = value["LowerSuperOutputArea XPath"]
								Ruleid_XPath = value["Ruleid XPath"]
								OutputColumnHeader=value["OutputColumnHeader"]
								colHead=OutputColumnHeader.split(';')
								worksheet=workbook.add_worksheet(section)
								worksheet.write(0,0,"CanonCity XPath",format)
								worksheet.write(0,1,CanonCity_XPath,entryFormat)
								worksheet.write_url(0,9,  'internal:Summary!A1',hyperlink,"Back To Summary Sheet")
								if(source == "Based on Lookup and raw input document"):
										worksheet.write(0,13,"Note: This is the possible Location Specific Information values based on the lookup. In addition to this, there will be other Location Specific Information values emitted based on input raw tags",format1)
								worksheet.write(1,0,"CanonZipCode XPath",format)
								worksheet.write(1,1,CanonZipCode_XPath,entryFormat)
								worksheet.write(2,0,"TravelToWorkArea XPath",format)
								worksheet.write(2,1,TravelToWorkArea_XPath,entryFormat)
								worksheet.write(3,0,"LocalAuthorityDistrict XPath",format)
								worksheet.write(3,1,LocalAuthorityDistrict_XPath,entryFormat)
								worksheet.write(4,0,"CanonCounty XPath",format)
								worksheet.write(4,1,CanonCounty_XPath,entryFormat)
								worksheet.write(5,0,"LocalEnterprisePartnership XPath",format)
								worksheet.write(5,1,LocalEnterprisePartnership_XPath,entryFormat)
								worksheet.write(6,0,"Region XPath",format)
								worksheet.write(6,1,Region_XPath,entryFormat)
								worksheet.write(7,0,"EnglishCountry XPath",format)
								worksheet.write(7,1,EnglishCountry_XPath,entryFormat)
								worksheet.write(8,0,"CanonLat XPath",format)
								worksheet.write(8,1,CanonLat_XPath,entryFormat)
								worksheet.write(9,0,"CanonLon XPath",format)
								worksheet.write(9,1,CanonLon_XPath,entryFormat)
								worksheet.write(10,0,"LowerSuperOutputArea XPath",format)
								worksheet.write(10,1,LowerSuperOutputArea_XPath,entryFormat)
								worksheet.write(11,0,"Ruleid XPath",format)
								worksheet.write(11,1,Ruleid_XPath,entryFormat)
								column = 0
								for val in colHead:
										worksheet.write(13,column,val,format)
										column = column+1
								row = 14
								for key in sorted(dict_location.keys()):
										if key not in ["NA","na"]:
												listofdata=dict_location.get(key)
												if listofdata[0]: #if the city is present then write the data's to the sheet 
														worksheet.write(row,1,key,entryFormat) #Zipcode as key 
														Col = 0
														for colhead in range(len(colHead)):
																if(colhead == 0):
																		continue
																else:
																		cell = colhead - 1 #value inside the list
																		#print listofdata
																		if listofdata[cell] in ["NA","-999","na"]:
																				continue
																		else:
																				if(Col == 1):
																						Col = Col+1
																						
																				worksheet.write(row,Col,listofdata[cell],entryFormat)   
																Col=Col+1
												else:
														continue
												row = row+1
										else:
												continue
										
										
								
								
								
						elif( section == "CanonSkilltoSkillClusterMapping"):
								value=config[section]
								Source = value["Source"]
								source = str(Source)
								CanonSkill_Xpath = value["CanonSkill Xpath"]
								CanonSkillCluster_Xpath = value["CanonSkillCluster Xpath"]
								OutputColumnHeader = value["OutputColumnHeader"]
								ColHead = OutputColumnHeader.split(';')
								worksheet=workbook.add_worksheet(section)
								worksheet.write(0,0,"CanonSkill Xpath",format)
								worksheet.write(0,1,CanonSkill_Xpath,entryFormat)
								worksheet.write_url(0,9,  'internal:Summary!A1',hyperlink,"Back To Summary Sheet")
								if(source == "Based on Lookup and raw input document"):
										worksheet.write(0,13,"Note: This is the possible CanonSkilltoSkillClusterMapping values based on the lookup. In addition to this, there will be other CanonSkilltoSkillClusterMapping values emitted based on input raw tags",format1)
								worksheet.write(1,0,"CanonSkillCluster Xpath",format)
								worksheet.write(1,1,CanonSkillCluster_Xpath,entryFormat)
								worksheet.write(3,0,ColHead[0],format)
								worksheet.write(3,1,ColHead[1],format)

								row = 4
								for val in sorted(skillsMapping_dict.keys()):
										worksheet.write_string(row,0,val,entryFormat)
										trimval = skillsMapping_dict[val].rstrip('\n')
										worksheet.write_string(row,1,trimval,entryFormat)
										row = row + 1  
								
								
						elif(section == "CanonJobType"):#CanonJobType Module
								value=config[section]
								Source = value["Source"]
								source = str(Source)
								JobType_Xpath = value["JobType Xpath"]
								JobHours_Xpath = value["JobHours Xpath"]
								JobTaxTerm_Xpath = value["JobTaxTerm Xpath"]
								WorkFromHome_Xpath = value["WorkFromHome Xpath"]
								OutputColumnHeader = value["OutputColumnHeader"]
								colhead = OutputColumnHeader.split(';')
								worksheet=workbook.add_worksheet(section)
								
								worksheet.write(0,0,"JobType Xpath",format)
								worksheet.write(0,1,JobType_Xpath,entryFormat)
								worksheet.write_url(0,9,  'internal:Summary!A1',hyperlink,"Back To Summary Sheet")
								if(source == "Based on Lookup and raw input document"):
										worksheet.write(0,13,"Note: This is the possible CanonJobType values based on the lookup. In addition to this, there will be other CanonJobType values emitted based on input raw tags",format1)
								worksheet.write(2,0,colhead[0],format)
								JobType_Data = value["JobType Data"]
								Jobtype = JobType_Data.split(';')
								row = 3
								for val in Jobtype:
										worksheet.write(row,0,val,entryFormat)
										row = row + 1

								worksheet.write(8,0,"JobHours Xpath",format)
								worksheet.write(8,1,JobHours_Xpath,entryFormat)
								worksheet.write(10,0,colhead[1],format)
								JobHours_Data = value["JobHours Data"]
								JobHours = JobHours_Data.split(';')
								row = 11
								for val in JobHours:
										worksheet.write(row,0,val,entryFormat)
										row = row + 1

								worksheet.write(14,0,"JobTaxTerm Xpath",format)
								worksheet.write(14,1,JobTaxTerm_Xpath,entryFormat)
								worksheet.write(16,0,colhead[0],format)
								JobTaxTerm_Data = value["JobTaxTerm Data"]
								JobTaxTerm = JobTaxTerm_Data.split(';')
								row = 17
								for val in JobTaxTerm:
										worksheet.write(row,0,val,entryFormat)
										row = row + 1

								worksheet.write(21,0,"WorkFromHome Xpath",format)
								worksheet.write(21,1,WorkFromHome_Xpath,entryFormat)
								worksheet.write(23,0,colhead[0],format)
								WorkFromHome_Data = value["WorkFromHome Data"]
								WorkFromHome = WorkFromHome_Data.split(';')
								row = 24
								for val in WorkFromHome:
										worksheet.write(row,0,val,entryFormat)
										row = row + 1
						elif(section == "CanonYearsOfExperience"):#for CanonYearsOfExperience module
								value=config[section]
								Source = value["Source"]
								source = str(Source)
								CanonYearsOfExperience_Level_XPath = value["CanonYearsOfExperience-Level XPath"]
								CanonYearsOfExperience_CanonLevel_XPath = value["CanonYearsOfExperience-CanonLevel XPath"]
								OutputColumnHeader = value["OutputColumnHeader"]
								colHeader = OutputColumnHeader.split(';')
								LevelData = value["LevelData"]
								level= LevelData.split(';')
								level.sort()
								CanonLevelData = value["CanonLevelData"]
								canonlevel = CanonLevelData.split(';')
								canonlevel.sort()
								worksheet=workbook.add_worksheet(section)
								worksheet.write(0,0,"CanonYearsOfExperience-Level XPath",format)
								worksheet.write(0,1,CanonYearsOfExperience_Level_XPath,entryFormat)
								worksheet.write_url(0,9,  'internal:Summary!A1',hyperlink,"Back To Summary Sheet")
								if(source == "Based on Lookup and raw input document"):
										worksheet.write(0,13,"Note: This is the possible CanonYearsOfExperience values based on the lookup. In addition to this, there will be other CanonYearsOfExperience values emitted based on input raw tags",format1)
								worksheet.write(2,0,colHeader[0],format)
								row = 3
								for val in level:
										worksheet.write_string(row,0,val,entryFormat)
										row = row + 1
								worksheet.write(8,0,"CanonYearsOfExperience-CanonLevel Xpath",format)
								worksheet.write(8,1,CanonYearsOfExperience_CanonLevel_XPath,entryFormat)
								worksheet.write(10,0,colHeader[1],format)
								row = 11
								for val in canonlevel:
										worksheet.write_string(row,0,val,entryFormat)
										row = row + 1
						elif(section == "ConsolidatedOnetRank"):
								value=config[section]
								Source = value["Source"]
								source = str(Source)
								data = value["Data"]
								Data = data.split(';')
								XPath = value["XPath"]
								OutputColumnHeader = value["OutputColumnHeader"]
								worksheet=workbook.add_worksheet(section)

								worksheet.write(0,0,"ConsolidatedOnetRank XPath",format)
								worksheet.write(0,1,XPath,entryFormat)
								worksheet.write_url(0,9,  'internal:Summary!A1',hyperlink,"Back To Summary Sheet")
								if(source == "Based on Lookup and raw input document"):
										worksheet.write(0,13,"Note: This is the possible ConsolidatedOnetRank values based on the lookup. In addition to this, there will be other ConsolidatedOnetRank values emitted based on input raw tags",format1)
								worksheet.write(2,0,OutputColumnHeader,format)
								row = 3
								for value in Data:
										worksheet.write(row,0,value,entryFormat)
										row = row + 1
						elif(section == "SSOCRank"):
								value=config[section]
								Source = value["Source"]
								source = str(Source)
								data = value["Data"]
								Data = data.split(';')
								XPath = value["XPath"]
								OutputColumnHeader = value["OutputColumnHeader"]
								worksheet=workbook.add_worksheet(section)

								worksheet.write(0,0,"SSOCRank XPath",format)
								worksheet.write(0,1,XPath,entryFormat)
								worksheet.write_url(0,9,  'internal:Summary!A1',hyperlink,"Back To Summary Sheet")
								if(source == "Based on Lookup and raw input document"):
										worksheet.write(0,13,"Note: This is the possible SSOCRank values based on the lookup. In addition to this, there will be other SSOCRank values emitted based on input raw tags",format1)
								worksheet.write(2,0,OutputColumnHeader,format)
								row = 3
								for value in Data:
										worksheet.write(row,0,value,entryFormat)
										row = row + 1
						elif(section == "NumberOfOpenings"):
								value=config[section]
								Source = value["Source"]
								source = str(Source)
								data = value["Data"]
								Data = data.split(';')
								XPath = value["XPath"]
								OutputColumnHeader = value["OutputColumnHeader"]
								worksheet=workbook.add_worksheet(section)

								worksheet.write(0,0,"NumberOfOpenings XPath",format)
								worksheet.write(0,1,XPath,entryFormat)
								worksheet.write_url(0,9,  'internal:Summary!A1',hyperlink,"Back To Summary Sheet")
								if(source == "Based on Lookup and raw input document"):
										worksheet.write(0,13,"Note: This is the possible NumberOfOpenings values based on the lookup. In addition to this, there will be other NumberOfOpenings values emitted based on input raw tags",format1)
								worksheet.write(2,0,OutputColumnHeader,format)
								row = 3
								for value in Data:
										worksheet.write(row,0,value,entryFormat)
										row = row + 1
						elif(section == "CanonNumberOfOpenings"):
								value=config[section]
								Source = value["Source"]
								source = str(Source)
								data = value["Data"]
								Data = data.split(';')
								XPath = value["XPath"]
								OutputColumnHeader = value["OutputColumnHeader"]
								worksheet=workbook.add_worksheet(section)

								worksheet.write(0,0,"CanonNumberOfOpenings XPath",format)
								worksheet.write(0,1,XPath,entryFormat)
								worksheet.write_url(0,9,  'internal:Summary!A1',hyperlink,"Back To Summary Sheet")
								if(source == "Based on Lookup and raw input document"):
										worksheet.write(0,13,"Note: This is the possible CanonNumberOfOpenings values based on the lookup. In addition to this, there will be other CanonNumberOfOpenings values emitted based on input raw tags",format1)
								worksheet.write(2,0,OutputColumnHeader,format)
								row = 3
								for value in Data:
										worksheet.write(row,0,value,entryFormat)
										row = row + 1
						elif(section == "ANZSCORank"):
								value=config[section]
								Source = value["Source"]
								source = str(Source)
								data = value["Data"]
								Data = data.split(';')
								XPath = value["XPath"]
								OutputColumnHeader = value["OutputColumnHeader"]
								worksheet=workbook.add_worksheet(section)

								worksheet.write(0,0,"ANZSCORank XPath",format)
								worksheet.write(0,1,XPath,entryFormat)
								worksheet.write_url(0,9,  'internal:Summary!A1',hyperlink,"Back To Summary Sheet")
								if(source == "Based on Lookup and raw input document"):
										worksheet.write(0,13,"Note: This is the possible ANZSCORank values based on the lookup. In addition to this, there will be other ANZSCORank values emitted based on input raw tags",format1)
								worksheet.write(2,0,OutputColumnHeader,format)
								row = 3
								for value in Data:
										worksheet.write(row,0,value,entryFormat)
										row = row + 1
								
								
						elif(section == "CanonSalary"):#for CanonSalary module 
								value=config[section]
								Source = value["Source"]
								source = str(Source)
								ISO_4217_XPath = value["ISO_4217_XPath"]
								SalType_XPath = value["SalType_XPath"]
								PayPeriod_XPath = value["PayPeriod_XPath"]
								OutputColumnHeader = value["OutputColumnHeader"]
								colHeader = OutputColumnHeader.split(';')
								ISOData = value["ISOData"]
								isodata = ISOData.split(';')
								#isodata.sort()
								SalData = value["SalData"]
								saldata = SalData.split(';')
								#saldata.sort()
								PayData = value["PayData"]
								paydata = PayData.split(';')
								#paydata.sort()
								worksheet=workbook.add_worksheet(section)

								worksheet.write(0,0,"ISO-4217 XPath",format)
								worksheet.write(0,1,ISO_4217_XPath,entryFormat)
								worksheet.write_url(0,9,  'internal:Summary!A1',hyperlink,"Back To Summary Sheet")
								if(source == "Based on Lookup and raw input document"):
										worksheet.write(0,13,"Note: This is the possible CanonSalary values based on the lookup. In addition to this, there will be other CanonSalary values emitted based on input raw tags",format1)
								worksheet.write(2,0,colHeader[0],format)
								row = 3
								for val in isodata:
										worksheet.write_string(row,0,val,entryFormat)
										row = row + 1
											 
								worksheet.write(11,0,"SalType XPath",format)
								worksheet.write(11,1,SalType_XPath,entryFormat)
								worksheet.write(13,0,colHeader[1],format)
								row = 14
								for val in saldata:
										worksheet.write_string(row,0,val,entryFormat)
										row = row + 1

								worksheet.write(21,0,"PayPeriod XPath",format)
								worksheet.write(21,1,PayPeriod_XPath,entryFormat)
								worksheet.write(23,0,colHeader[2],format)
								row = 24
								for val in paydata:
										worksheet.write_string(row,0,val,entryFormat)
										row = row + 1
						elif(section == "ConsolidatedInferredUKSIC"):
								value=config[section]
								Source = value["Source"]
								source = str(Source)
								XPath=value["XPath"]
								OutputColumnHeader=value["OutputColumnHeader"]
								worksheet=workbook.add_worksheet(section)
								worksheet.write(0,0,section+" XPath",format)
								worksheet.write(0,1,XPath,entryFormat)
								worksheet.write_url(0,9,  'internal:Summary!A1',hyperlink,"Back To Summary Sheet")
								if(source == "Based on Lookup and raw input document"):
										worksheet.write(0,13,"Note: This is the possible ConsolidatedInferredUKSIC values based on the lookup. In addition to this, there will be other ConsolidatedInferredUKSIC values emitted based on input raw tags",format1)
								worksheet.write(2,0,OutputColumnHeader,format)
								row = 3
								result = list(unique_everseen(listofdataIndustry))#to make the list unique
								result.sort()
								for var in range(len(result)):
										value = result[var]
										if (len(value)<1):
												continue
										else:
												worksheet.write(row,0,value,entryFormat)
												row=row+1
						elif(section == "ConsolidatedInferredANZSIC"):
								value=config[section]
								Source = value["Source"]
								source = str(Source)
								XPath=value["XPath"]
								OutputColumnHeader=value["OutputColumnHeader"]
								worksheet=workbook.add_worksheet(section)
								worksheet.write(0,0,section+" XPath",format)
								worksheet.write(0,1,XPath,entryFormat)
								worksheet.write_url(0,9,  'internal:Summary!A1',hyperlink,"Back To Summary Sheet")
								if(source == "Based on Lookup and raw input document"):
										worksheet.write(0,13,"Note: This is the possible ConsolidatedInferredANZSIC values based on the lookup. In addition to this, there will be other ConsolidatedInferredANZSIC values emitted based on input raw tags",format1)
								worksheet.write(2,0,OutputColumnHeader,format)
								row = 3
								result = list(unique_everseen(anzlistofdataIndustry))#to make the list unique
								result.sort()
								for var in range(len(result)):
										value = result[var]
										if (len(value)<1):
												continue
										else:
												temp = value.rstrip('\n')
												worksheet.write(row,0,temp,entryFormat)
												row=row+1
						elif(section == "ConsolidatedInferredNAICS"):
								value=config[section]
								Source = value["Source"]
								source = str(Source)
								XPath=value["XPath"]
								OutputColumnHeader=value["OutputColumnHeader"]
								worksheet=workbook.add_worksheet(section)
								worksheet.write(0,0,section+" XPath",format)
								worksheet.write(0,1,XPath,entryFormat)
								worksheet.write_url(0,9,  'internal:Summary!A1',hyperlink,"Back To Summary Sheet")
								if(source == "Based on Lookup and raw input document"):
										worksheet.write(0,13,"Note: This is the possible ConsolidatedInferredNAICS values based on the lookup. In addition to this, there will be other ConsolidatedInferredNAICS values emitted based on input raw tags",format1)
								worksheet.write(2,0,OutputColumnHeader,format)
								row = 3
								#intlist = list()
								#for var in uslistofdataIndustry:
										#val = int(var,36)
										#intlist.append(val)
										#intlist.sort()
								#uniqlist = list(unique_everseen(intlist))
								result = list(unique_everseen(uslistofdataIndustry))#to make the list unique
								result.sort()
								for var in range(len(result)):
										#value = result[var].decode('utf-8')  #to rectify the unicode decode error ,we are decoding each value
										value = result[var] #Python 3.4 , the string is already decoded - KIK:06-APR-2018
										if (len(value)<1):
												continue
										else:
												temp = value.rstrip('\n')
												worksheet.write(row,0,temp,entryFormat)
												row=row+1
								
						elif(section == "ConsolidatedUKSICRuleid"):
								value=config[section]
								Source = value["Source"]
								source = str(Source)
								XPath=value["XPath"]
								OutputColumnHeader=value["OutputColumnHeader"]
								worksheet=workbook.add_worksheet(section)
								worksheet.write(0,0,section+" XPath",format)
								worksheet.write(0,1,XPath,entryFormat)
								worksheet.write_url(0,9,  'internal:Summary!A1',hyperlink,"Back To Summary Sheet")
								if(source == "Based on Lookup and raw input document"):
										worksheet.write(0,13,"Note: This is the possible ConsolidatedUKSICRuleid values based on the lookup. In addition to this, there will be other ConsolidatedUKSICRuleid values emitted based on input raw tags",format1)
								worksheet.write(2,0,OutputColumnHeader,format)
								row = 3
								result = list(unique_everseen(listIndustryRuleid)) #to make the list unique
								result.sort()
								for var in range(len(result)):
										value = result[var]
										if (len(value)<1):
												continue
										else:
												worksheet.write(row,0,value,entryFormat)
												row=row+1
						elif(section == "ConsolidatedNAICSRuleid"):
								value=config[section]
								Source = value["Source"]
								source = str(Source)
								XPath=value["XPath"]
								OutputColumnHeader=value["OutputColumnHeader"]
								worksheet=workbook.add_worksheet(section)
								worksheet.write(0,0,section+" XPath",format)
								worksheet.write(0,1,XPath,entryFormat)
								worksheet.write_url(0,9,  'internal:Summary!A1',hyperlink,"Back To Summary Sheet")
								if(source == "Based on Lookup and raw input document"):
										worksheet.write(0,13,"Note: This is the possible ConsolidatedNAICSRuleid values based on the lookup. In addition to this, there will be other ConsolidatedNAICSRuleid values emitted based on input raw tags",format1)
								worksheet.write(2,0,OutputColumnHeader,format)
								row = 3
								result = list(unique_everseen(uslistIndustryRuleid)) #to make the list unique
								result.sort()
								for var in range(len(result)):
										value = result[var]
										if (len(value)<1):
												continue
										else:
												worksheet.write(row,0,value,entryFormat)
												row=row+1
						elif(section == "StdMajorCIPCode"):#for StdMajor-CIPCode module
								value=config[section]
								Source = value["Source"]
								source = str(Source)
								StdMajor_XPath = value["StdMajor XPath"]
								CIPCode_XPath = value["CIPCode XPath"]
								OutputColumnHeader = value["OutputColumnHeader"]
								ColHead = OutputColumnHeader.split(';')
												
								worksheet=workbook.add_worksheet(section)
								worksheet.write(0,0,"StdMajor XPath",format)
								worksheet.write(0,1,StdMajor_XPath,entryFormat)
								worksheet.write_url(0,9,  'internal:Summary!A1',hyperlink,"Back To Summary Sheet")
								if(source == "Based on Lookup and raw input document"):
										worksheet.write(0,13,"Note: This is the possible StdMajorCIPCode values based on the lookup. In addition to this, there will be other StdMajorCIPCode values emitted based on input raw tags",format1)
								worksheet.write(1,0,"CIPCode XPath",format)
								worksheet.write(1,1,CIPCode_XPath,entryFormat)
								worksheet.write(3,0,ColHead[0],format)
								worksheet.write(3,1,ColHead[1],format)
								row = 4
								for val in sorted(uniq_dict_std.keys()):
										worksheet.write_string(row,0,val,entryFormat)
										worksheet.write_string(row,1,uniq_dict_std[val],entryFormat)
										row = row + 1
						for key in dict_of_list.keys():#for the sections which formed in a dictionary by the Extractmodule function
								if(key==section):
										if(section in ["Onet","BGTOcc","LocalGovt","ANZSCO","ANZBGTOcc","SSOC","BGTOccSGP"]):
												listofdata=dict_of_list.get(key)
												listofdata=[k for k in listofdata if k !="NA"]
												value=config[section]
												Source = value["Source"]
												source = str(Source)
												ColumnHeader=value["OutputColumnHeader"]
												xpath=value["XPath"]
												colhead=ColumnHeader.split(';')
												if(section == "BGTOccSGP"):
													worksheet=workbook.add_worksheet("BGTOcc")
												else:
													worksheet=workbook.add_worksheet(section)
												worksheet.write(0,0,key+"XPath",format)
												worksheet.write(0,1,xpath,entryFormat)
												worksheet.write_url(0,9,  'internal:Summary!A1',hyperlink,"Back To Summary Sheet")
												if(source == "Based on Lookup and raw input document"):
														worksheet.write(0,13,"Note: This is the possible %s values based on the lookup. In addition to this, there will be other %s values emitted based on input raw tags"%(section,section),format1)
												for i in range(len(colhead)):
														worksheet.write(2,i,colhead[i],format)
												row=3
												if section in ["Onet","BGTOcc","ANZBGTOcc","SSOC","BGTOccSGP"]:
														#print "pppp"
														for val in (unique_everseen(listofdata)):
																#print val
														
																if val in sorted(uniq_dict.keys()):
																		#print val
																		worksheet.write_string(row,0,val,entryFormat)
																		worksheet.write_string(row,1,uniq_dict[val],entryFormat)
																		row=row+1
																		
																else:
																		continue
												elif section in ["LocalGovt","ANZSCO"]:
														for val in (unique_everseen(listofdata)):
															 val1=val.rstrip('\n')
															 key="%s.0"%(val1)
															 if key in sorted(uniq_dict.keys()):
																	 worksheet.write_string(row,0,val1,entryFormat)
																	 worksheet.write_string(row,1,uniq_dict[key],entryFormat)
																	 row=row+1
															 else:
																	 continue
											 
										
										
										elif(section == "ConsolidatedDegree"):
												listofdata=dict_of_list.get(key)
												listofdata.sort()
												listofdata=[k for k in listofdata if k !="NA"]
												value=config[section]
												Source = value["Source"]
												source = str(Source)
												Degree_XPath = value["Degree XPath"]
												MaxDegree_XPath = value["MaxDegree XPath"]
												MinDegree_XPath = value["MinDegree XPath"]
												PreferredDegree_XPath = value["PreferredDegree XPath"]
												RequiredDegree_XPath = value["RequiredDegree XPath"]
												ColumnHeader=value["OutputColumnHeader"]
												worksheet=workbook.add_worksheet(key)
												worksheet.write(0,0,key+"Degree XPath",format)
												worksheet.write(0,1,Degree_XPath,entryFormat)
												worksheet.write_url(0,9,  'internal:Summary!A1',hyperlink,"Back To Summary Sheet")
												if(source == "Based on Lookup and raw input document"):
														worksheet.write(0,13,"Note: This is the possible ConsolidatedDegree values based on the lookup. In addition to this, there will be other ConsolidatedDegree values emitted based on input raw tags",format1)
												worksheet.write(1,0,key+"MaxDegree XPath",format)
												worksheet.write(1,1,MaxDegree_XPath,entryFormat)
												worksheet.write(2,0,key+"MinDegree XPath",format)
												worksheet.write(2,1,MinDegree_XPath,entryFormat)
												worksheet.write(3,0,key+"PreferredDegree XPath",format)
												worksheet.write(3,1,PreferredDegree_XPath,entryFormat)
												worksheet.write(4,0,key+"RequiredDegree XPath",format)
												worksheet.write(4,1,RequiredDegree_XPath,entryFormat)
												worksheet.write(6,0,ColumnHeader,format)
												row=7
												
												for var in list(unique_everseen(listofdata)):
														if (len(var)<1):
																continue
														else:
																worksheet.write_string(row,0,var,entryFormat)
																row=row+1
												
										elif(section == "ConsolidatedDegreeLevels"):
												listofdata=dict_of_list.get(key)
												listofdata.sort()
												listofdata=[k for k in listofdata if k !="NA"]
												value=config[section]
												Source = value["Source"]
												source = str(Source)
												XPath = value["XPath"]
												ColumnHeader=value["OutputColumnHeader"]
												worksheet=workbook.add_worksheet(key)
												worksheet.write(0,0,key+"XPath",format)
												worksheet.write(0,1,XPath,entryFormat)
												worksheet.write_url(0,9,  'internal:Summary!A1',hyperlink,"Back To Summary Sheet")
												if(source == "Based on Lookup and raw input document"):
														worksheet.write(0,13,"Note: This is the possible ConsolidatedDegree values based on the lookup. In addition to this, there will be other ConsolidatedDegree values emitted based on input raw tags",format1)
												worksheet.write(1,0,"MaxDegree XPath",format)
												worksheet.write(1,1,"//JobDoc/DataElementsRollup/MaxDegreeLevel",entryFormat)
												worksheet.write(2,0,"MinDegree XPath",format)
												worksheet.write(2,1,"//JobDoc/DataElementsRollup/MinDegreeLevel",entryFormat)
												worksheet.write(3,0,"PreferredDegree XPath",format)
												worksheet.write(3,1,"//JobDoc/DataElementsRollup/PreferredDegreeLevels",entryFormat)
												worksheet.write(4,0,"RequiredDegree XPath",format)
												worksheet.write(4,1,"//JobDoc/DataElementsRollup/RequiredDegreeLevels",entryFormat)
												worksheet.write(6,0,ColumnHeader,format)
												row=7
												
												for var in list(unique_everseen(listofdata)):
														if (len(var)<1):
																continue
														else:
																worksheet.write_string(row,0,var,entryFormat)
																row=row+1
										elif(section == "AQFLevel"):
												listofdata=dict_of_list.get(key)
												listofdata.sort()
												listofdata=[k for k in listofdata if k !="NA"]
												value=config[section]
												Source = value["Source"]
												source = str(Source)
												MinXPath = value["MinXPath"]
												MaxXPath = value["MaxXPath"]
												PreferredXPath = value["PreferredXPath"]
												RequiredXPath = value["RequiredXPath"]
												ColumnHeader=value["OutputColumnHeader"]
												worksheet=workbook.add_worksheet(key)
												worksheet.write(0,0,"MinAQFLevel",format)
												worksheet.write(0,1,MinXPath,entryFormat)
												worksheet.write_url(0,9,  'internal:Summary!A1',hyperlink,"Back To Summary Sheet")
												if(source == "Based on Lookup and raw input document"):
														worksheet.write(0,13,"Note: This is the possible ConsolidatedDegree values based on the lookup. In addition to this, there will be other ConsolidatedDegree values emitted based on input raw tags",format1)
												worksheet.write(1,0,"MaxAQFLevel",format)
												worksheet.write(1,1,MaxXPath,entryFormat)
												worksheet.write(2,0,"PreferredAQFLevel",format)
												worksheet.write(2,1,PreferredXPath,entryFormat)
												worksheet.write(3,0,"RequiredAQFLevel",format)
												worksheet.write(3,1,RequiredXPath,entryFormat)
												
												worksheet.write(5,0,ColumnHeader,format)
												row=6
												
												for var in list(unique_everseen(listofdata)):
														if (len(var)<1):
																continue
														else:
																worksheet.write_string(row,0,var,entryFormat)
																row=row+1
										
										
										else:
												listofdata=dict_of_list.get(key)
												intlist = list()
												if section in ["ConsolidatedInferredNAICS","SourceID"]:
														for var in listofdata:
																val = int(var)
																intlist.append(val)
																intlist.sort()
														uniqlist = list(unique_everseen(intlist))
														#print uniqlist
																
												else:
														listofdata.sort()
												#print listofdata
												listofdata=[k for k in listofdata if k !="NA"]
												value=config[section]
												Source = value["Source"]
												source = str(Source)
												xpath=value["XPath"]
												ColumnHeader=value["OutputColumnHeader"]
												worksheet=workbook.add_worksheet(key)
												worksheet.write(0,0,key+"XPath",format)
												worksheet.write(0,1,xpath,entryFormat)
												worksheet.write_url(0,9,  'internal:Summary!A1',hyperlink,"Back To Summary Sheet")
												if(source == "Based on Lookup and raw input document"):
														worksheet.write(0,13,"Note: This is the possible %s value based on the lookup. In addition to this, there will be other %s values emitted based on input raw tags"%(section,section),format1)
												worksheet.write(2,0,ColumnHeader,format)
												row=3
												
												result = set(listofdata)
												#result = [item for item in result if item.title() not in result]
												result = [item for item in result if item.istitle() or item.title() not in result]
												#final = list()
												#for item in result:
												#		final.append(item)
												
												result.sort()
												
												
												if section in ["ConsolidatedInferredNAICS","SourceID"]:
														for var in uniqlist:
																#print var
																worksheet.write(row,0,var,entryFormat)
																row=row+1
												elif(section == "StandardTitle"):
														for var in list(unique_everseen(result)):
																if (len(var)<1):
																		continue
																else:
																		value = var.rstrip('\n')
																		worksheet.write_string(row,0,value,entryFormat)
																		row=row+1
												elif section in ["ConsolidatedOnetRuleId","CertificationRuleid","CanonEmployerRuleID","ConsolidatedDegreeRuleid","DegreeExclusionRuleid","StdMajorRuleid","SkillsSkillid","SkillsRuleId","CanonIntermediaryRuleid","BGTOccRuleid","ConsolidatedSSOCUniqueID"]:
														for var in list(unique_everseen(listofdata)):
																if (len(var)<1):
																		continue
																else:
																		trimvar = var.rstrip('\n')
																		if(trimvar==''):
																			continue
																		else:
																			worksheet.write_string(row,0,trimvar,entryFormat)
																			row=row+1
												else:
														#for var in list(unique_everseen(listofdata)):
														new = list()
														for val in result:
																if val in new:
																		continue
																else:
																		new.append(val)
														finalresult =list(unique_everseen(new))
														
														for var in finalresult:
																if (len(var)<1):
																		continue
																else:
																		trimvar = var.rstrip('\n')
																		worksheet.write_string(row,0,trimvar,entryFormat)
																		row=row+1
				else:
						continue
#print sectionlist
workbook.close()
