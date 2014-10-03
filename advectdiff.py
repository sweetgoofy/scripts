from pyxll import xl_func
from pyxll import xl_macro
import math


def cache(fn):
	""" cache return values for all decorated functions,
		so that no function with same arguments are computed twice.
	"""
	cache.c = dict()
	def _fn(*args, **kwargs):
		key = fn.__name__ + str(args) + str(kwargs)
		try:
			ret = cache.c[key]
		except KeyError, e:
			ret = fn(*args, **kwargs)
			cache.c[key] = ret
		return ret
	return _fn


@cache
@xl_func("float[] inp: float")
def getY90(inp):
	""" take a [y, c] list, and returns Y90.
		if no precise value found then linearly interpolate for Y90.
		input array does not need to be sorted.
	"""
	inp = sorted(inp, key = lambda x: x[0])
	for i in xrange(len(inp)):
		e = inp[i]
		if e[1]>=0.9:
			second = e
			first = inp[i-1]
			break
	return first[0]+(0.9-first[1])/(second[1]-first[1])*(second[0]-first[0])

@cache
@xl_func("float[] inp: float")
def getClearWaterDepth(inp):
	""" return the equivalent clear water depth d using trapezoidal sums.
		if no precise value for C90 is found then linearly interpolate for C90.
		input array does not need to be sorted.
	"""
	y90 = getY90(inp)
	inp = sorted(inp, key = lambda x: x[0])
	s = 0
	if inp[0][0] > 0:
		s += (1-inp[0][1]/2.0) * inp[0][0]
	for i in xrange(1,len(inp)):
		prev = inp[i-1]
		cur = inp[i]
		if cur[0] > y90:
			top = 0.9
			base = prev[1]
			height = y90 - prev[0]
			s += (1 - (top + base)/2) * height
			break
		base = prev[1]
		top = cur[1]
		height = cur[0] - prev[0]
		s += (1 - (top + base)/2) * height
		
	return s

@cache
@xl_func("float[] inp: float")
def getCMean(inp):
	""" return the mean void fraction concentration 1 - d/Y90
		input array does not need to be sorted.
	"""
	inp = sorted(inp, key = lambda x: x[0])
	return 1 - getClearWaterDepth(inp) / getY90(inp)

@cache
@xl_func("float[] inp: float")
def getD0(inp):
	""" input array does not need to be sorted.
		-1/0.364 * ln(1.0434 - Cmean/0.7622)
	"""
	cmean = getCMean(inp)
	return (-1/3.614) * math.log(1.0434 - cmean/0.7622)

@cache
@xl_func("float[] inp: float")
def getK1(inp):
	""" input array does not need to be sorted.
		K' = 0.032745... + 1/(2 * D0) - 8/(81 * D0)
	"""
	d0 = getD0(inp)
	return 0.32745 + 1/(2 * d0) - 8/(81 * d0)

@cache
@xl_func("float[] inp: float")
def getLambda(inp):
	""" input array does not need to be sorted.
		lambda = 0.9 * (K' - Cmean)
	"""
	return 0.9/(getK1(inp) - getCMean(inp))

@cache
@xl_func("float[] inp: float")
def getK2(inp):
	""" input array does not need to be sorted.
		K'' = 0.9/(1 - exp(-lambda))
	"""
	return 0.9/(1-math.exp(-getLambda(inp)))

@cache
@xl_func("float y, float[] inp: float")
def estimateCt(y, inp):
	""" estimate C for location y, for a transition flow regime.
		input array does not need to be sorted.
		C(y) = K'' * (1 - exp(-lambda * y/Y90))
	"""
	return getK2(inp) * (1 - math.exp(-getLambda(inp) * y / getY90(inp)))

@cache
@xl_func("float y, float[] inp: float")
def estimateCs(y, inp):
	""" estimate C for a location y, for a skimming flow regime.
		input array does not need to be sorted.
		C(y) = 1 - tanh^2(K' - (y/Y90)/(2*D0) + (y/Y90 - 1/3)^3/(3*D0))
	"""
	return 1 -(math.tanh(getK1(inp) - (y/getY90(inp)) / (2 * getD0(inp)) + (y/getY90(inp) - 1/3.0)**3 / (3 * getD0(inp))))**2

@cache
@xl_func("float y, float[] inp: float")
def estimateCsKnownY90(y, inp):
	""" estimate C for a location y' = y/Y90.
		input array does not need to be sorted.
		C(y') = 1 - tanh^2(K' - y'/(2*D0) + (y' - 1/3)^3/(3*D0))
	"""
	return 1 -(math.tanh(getK1(inp) - y / (2 * getD0(inp)) + (y - 1/3.0)**3 / (3 * getD0(inp))))**2

@cache
@xl_func("float target, float[] inp, int idx: var")
def vLookUpInterp(target, inp, idx):
	"""
		imitate VLOOKUP in excel.
		if target not found then interpolate between the nearest pair.
	"""
	if len(inp) < 2: return "Error! Invalid input"
	inp = sorted(inp, key = lambda x: x[0])
	for i in xrange(1, len(inp)):
		cur = inp[i]
		prev = inp[i-1]
		if cur[0] >= target:
			try:
				return prev[idx] + (target - prev[0]) / (cur[0] - prev[0]) * (cur[idx] - prev[idx])
			except:
				return "Error! Check index"
	# not found
	return "All values less than target, check input"

@xl_func("string s: float")
def getDc(s):
	"""
		return dc from a description string
	"""
	return float(s.split("=")[-1])*0.1

@xl_func("float h: float")
def qFelder(h):
	"""
		calculate q = Q/W, given upstream head h = H1
	"""
	return (0.92 + 0.153 * h/1.01) * math.sqrt(9.8 * (2/3.0 * h)**3)

@xl_func("float x, float y: float")
def transformY(x, y):
	c = 0.894427
	s = 0.447214
	xt = x/c
	# return (y + xt * s - 100) / c
	return -100*c + x * s + y * c


@cache
@xl_func("float[] pt, float[] points, float size: float")
def findAdjacentPoints(pt, points, size):
	# px = sorted(points, key = lambda x: x[0])
	# py = sorted(points, key = lambda y: y[0])
	# xmin, xmax, ymin, ymax = px[0], px[-1], py[0], py[-1]
	
	# deltaX = (xmax - xmin)/resolution
	# deltaY = (ymax - ymin)/resolution

	def containsPt(rect, pt):
		# rect(xmin, xmax, ymin, ymax)
		xmin, xmax, ymin, ymax = rect[0], rect[1], rect[2], rect[3]
		return (xmin <= pt[0] and pt[0] <= xmax) and (ymin <= pt[1] and pt[1] <= ymax)

	def rect(pt, size):
		return (pt[0] - size/2.0, pt[0] + size/2.0, pt[1] - size/2.0, pt[1] + size/2.0)

	r = rect(pt[0], size)
	count = 0
	for p in points:
		if containsPt(r, p): 
			count += 1

	return count
	# for x in xrange(xmin, xmax, deltaX):
	# 	for y in xrange(ymin, ymax, deltaY):
	# 		rect = (x - size/2.0, x + size/2.0, y - size/2.0, y + size/2.0)


@xl_func("float q, float[] yc, float[] v: var")
def alphaCorrection(q, yc, v):
	""" calculate kinetic energy correction coefficient alpha
		for a variable density fluid.
		input arrays must be sorted.
		elevation must be in mm.
	"""
	y90 = getClearWaterDepth(yc)
	yArr = map(lambda x: x[0], yc)
	cArr = map(lambda x: x[1], yc)
	vArr = map(lambda x: x[0], v)

	assert len(yc) == len(v)

	s = 0
	if yArr[0] > 0:
		s += yArr[0] / 1000 * (2 - cArr[0]) / 2 * (vArr[0] / 2) ** 3

	for i in xrange(1, len(yArr)):
		prevY = yArr[i-1]
		prevC = cArr[i-1]
		prevV = vArr[i-1]

		curY = yArr[i]
		curC = cArr[i]
		curV = vArr[i]

		if curC > 0.9:
			# reached Y90
			v90 = prevV + (0.9 - prevC) / (curC - prevC) * (curV - prevV)
			avgC = (0.9 + prevC) / 2
			avgV = (v90 + prevV) / 2
			s += (1 - avgC) * avgV ** 3 * (getY90(yc) - prevY) / 1000
			break

		avgV = (curV + prevV) / 2
		avgC = (curC + prevC) / 2
		s += (1 - avgC) * avgV ** 3 * (curY - prevY) / 1000
	return s * (getClearWaterDepth(yc) / 1000) ** 2 / q ** 3