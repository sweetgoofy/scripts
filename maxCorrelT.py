from pyxll import xl_func
import re
@xl_func("string path: float")
def maxCorrelT(path):
	inFile=open(path)
	maxTuple=()
	inFile.readline() #discard first line
	maxVal=-99
	while True:
		line=inFile.readline()
		if not line: #if EOF reached
			break
		t=line.split() #split line into tuple (t, v)
		v=t[1]
		if v>maxVal:
			maxVal=v
			maxTuple=t
	return float(maxTuple[0])




