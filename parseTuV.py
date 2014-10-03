from pyxll import xl_macro
from pyxll import xl_func
import re

@xl_macro("string fileURL: var")
def parseInFile(fileURL):
	inFile=open(fileURL)
	inFile.readline()
	#arrOut=[[],[]]
	arrout=[[]]
	while True:
		line = inFile.readline()
		if not line:
			break
		data=re.split('\s+', line)
		iTu=len(data)-3
		iV=len(data)-2
		# arrOut[0].append(data[iTu])
		# arrOut[1].append(data[iV])
		arrOut[0].append(data[iTu])
		arrOut[0].append(data[iV])
	return arrOut


@xl_func("string x: int")
def py_strlen(x):
    """returns the length of x"""
    return len(x)