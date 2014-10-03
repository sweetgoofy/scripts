from pyxll import xl_func
from pyxll import xl_macro
import re
@xl_func("string path: string",macro=True)
def parseFilePath(path):
	# C:\Users\uqgzhan3\Desktop\PhD_Exp_Data\Series_B2\20140314\steps8-9\x'=10cm\h=0.1395
	sTuple=path.split('\\')
	dateRE=re.compile('\d\d\d\d\d\d\d\d')
	hRE=re.compile('h=[0.\d*]')
	date=""
	h = 0
	for s in sTuple:
		if dateRE.match(s):
			date=s
		if hRE.match(s):
			h=float((s.split('='))[1])
	dcOverH=round(((((0.92+0.153*h/1.01)*((9.8*(2.0/3.0*h)**3))**(0.5))**2/9.8)**(1.0/3.0))/0.1,3)
	return "%s_h1=%s_dc/h=%s" % (date, str(h), str(dcOverH))


@xl_func("string path: float")
def Vc(path):
	"""given a preformatted string,
	calculate the critical velocity Vc"""
	result = re.findall(r'0\.\d*', path)
	h = float(result[0])
	#dcOverH=round(((((0.92+0.153*h/1.01)*((9.8*(2.0/3.0*h)**3))**(0.5))**2/9.8)**(1.0/3.0))/0.1,3)
	dc=(((0.92+0.153*h/1.01)*((9.8*(2.0/3.0*h)**3))**(0.5))**2/9.8)**(1.0/3.0)
	Vc = (9.8 * dc) ** 0.5
	return Vc


