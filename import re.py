from pyxll import xl_macro
import re
import win32api

@xl_macro("string fileURL: float[]")
def parseInFile(fileURL):
	inFile=open(fileURL)
	inFile.readline()
	arrOut=[[],[]]
	while not EOF:
		data=re.split('\s+', inFile.readline())
		iTu=len(data)-3
		iV=len(data)-2
		arrOut[0].append(data[iTu])
		arrOut[1].append(data[iV])
	return arrOut