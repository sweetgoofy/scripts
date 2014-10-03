from pyxll import xl_macro
from pyxll import xl_func
import re

"""
Pseudocode:
	def filterBubbles(data, constraint):
		parse(data)
		return [d in data if satisfies(d,constraint)]

*** will want to get stats for these bubbles ***
"""

class Bubble:
	"""
	each instance represents a bubble object.
	properties:
		- int id - identifier
		- double[] t - time
		- double[] x - x coords
		- double[] y - y coords
	"""
	def __init__(	self,
					idf = -1, 
					t = [], 
					x = [], 
					y = []):
		self.id = idf
		self.t = t
		self.x = x
		self.y = y
	
	def appendT(self, t):
		self.t.append(t)

	def appendX(self, x):
		self.x.append(x)

	def appendY(self, y):
		self.y.append(y)

	def getId(self):
		return self.id

	def getX(self):
		return self.x

	def getY(self):
		return self.y

	def getT(self):
		return self.t

def filterBubbles(data, xCons1 = None, xCons2 = None, yCons1 = None, yCons2 = None):
	"""
		Filters an array of bubbles according to the given criteria.
		
		@Parameters:
			data - excel range containing all bubbles in the form [id, [t], [x], [y]]
			xCons - constraint for x, in the form of "[<>=]+\d+"
			yCons - constraint for y, in the form of "[<>=]+\d+"
		@Return:
			an array of filtered bubble objects.
	"""
	bubbles = initBubbles(data)
	ret = []
	for b in bubbles:
		isValid = True
		if xCons1 and not(eval(str(b.x[0]) + xCons1)): isValid = False
		if xCons2 and not(eval(str(b.x[0]) + xCons2)): isValid = False
		if yCons1 and not(eval(str(b.y[0]) + yCons1)): isValid = False
		if yCons2 and not(eval(str(b.y[0]) + yCons2)): isValid = False
		if isValid: ret.append(b)
	return ret

def initBubbles(data):
	# data == Excel range object corresponding to list of lists
	curBubble = None
	bubbles = []
	for row in data: # row: [id, t, x, y]
		if row[0] > 0:
			b = Bubble(row[0], [], [], [])
			if curBubble: bubbles.append(curBubble)	# append the previous bubble each time we find a new bubble
			curBubble = b
		if curBubble:
			curBubble.appendT(row[1])
			curBubble.appendX(row[2])
			curBubble.appendY(row[3])
	if curBubble: bubbles.append(curBubble)	# append the last bubble
	
	return bubbles

def convertToRange(bubbles):
	""" Takes a list of bubbles and return an Excel range object"""
	arrOut = []
	for b in bubbles:
		for i in xrange(len(b.getX())):
			if i == 0:
				arrOut.append([b.getId(), b.getT()[i], b.getX()[i], b.getY()[i]])
			else:
				arrOut.append(["", b.getT()[i], b.getX()[i], b.getY()[i]])
	return arrOut

@xl_macro("float[] data, string xCons1, string xCons2, string yCons1, string yCons2: var")
def selectBubbles(	data, 
					xCons1, 
					xCons2,
					yCons1, 
					yCons2):

	p = re.compile("((<)|(>)|(==)|(>=)|(<=)){1}-*(\d.)*\d+")
	isValid = True
	
	if len(xCons1)>0:
		if not p.match(xCons1): isValid = False
	if len(xCons2)>0:
		if not p.match(xCons2): isValid = False
	if len(yCons1)>0:
		if not p.match(yCons1): isValid = False
	if len(yCons2)>0:
		if not p.match(yCons2): isValid = False
	
	return convertToRange(filterBubbles(data, xCons1, xCons2, yCons1, yCons2)) if isValid else "invalid constraints detected."
