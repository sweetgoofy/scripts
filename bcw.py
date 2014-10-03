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


@xl_func("float[] y, float[] v, float offset: float")
def q(y, v, offset=0):
	""" take an Excel y', V array and integrate to find q
		data - [y' (mm), V (m/s)] array
		offset - offset in y' such that y' + offset = y
	"""
	assert len(y) == len(v)
	data = []
	for i in xrange(len(y)):
		data.append([y[i][0], v[i][0]])
	data = sorted(data, key = lambda x: x[0]) # sort data wrt y'
	# assert isinstance(data[0], list)
	adjustedData = [[entry[0] + offset, entry[1]] for entry in data] # correct elevation with offset
	ret = 0

	
	cur = adjustedData[0]
	ret += (cur[1] + 0) / 2.0 * cur[0] / 1000

	for i in xrange(1, len(adjustedData)):
		cur = adjustedData[i]
		prev = adjustedData[i-1]
		d0, d1, v0, v1 = prev[0], cur[0], prev[1], cur[1]
		ret += (v0 + v1) / 2.0 * (d1 - d0) / 1000
	return ret

@xl_func("float h: float")
def qBCW(h):
	return math.sqrt(9.8 * (2 / 3.0 * h) ** 3)

@xl_func("float[] y, float[] h, float offset: float")
def HBar(y, h, offset):
	""" take an Excel y', H array and integrate to find Hbar
		data - [y' (mm), H] array
		offset - offset in y' such that y' + offset = y
	"""
	assert len(y) == len(h)
	data = []
	for i in xrange(len(y)):
		data.append([y[i][0], h[i][0]])
	data = sorted(data, key = lambda x: x[0]) # sort data wrt y'

	adjustedData = [[entry[0] + offset, entry[1]] for entry in data] # correct elevation with offset
	ret = 0


	for i in xrange(1, len(adjustedData)):
		cur = adjustedData[i]
		prev = adjustedData[i-1]
		d0, d1, v0, v1 = prev[0], cur[0], prev[1], cur[1]
		ret += (v0 + v1) / 2.0 * (d1 - d0)
	return ret / (adjustedData[-1][0] - adjustedData[0][0])

@xl_func("float target, float[] inp, int idx, int interpIdx: var")
def vLookUpInterp2(target, inp, idx, interpIdx = 0):
	"""
		imitate VLOOKUP in excel.
		if target not found then interpolate between the nearest pair.

	"""
	if len(inp) < 2: return "Error! Invalid input"
	#inp = sorted(inp, key = lambda x: x[interpIdx])
	for i in xrange(1, len(inp)):
		cur = inp[i]
		prev = inp[i-1]
		if cur[interpIdx] >= target:
			try:
				return prev[idx] + (target - prev[interpIdx]) / (cur[interpIdx] - prev[interpIdx]) * (cur[idx] - prev[idx])
			except:
				return "Error! Check index"
	# not found
	return "All values less than target, check input"

@xl_func("float[] y, float[] v, float v0, float delta, float offset: float")
def delta1(y, v, v0, delta, offset = 0):
	"""
		given arrays y, v and parameters v0, delta,
		return the displacement thickness delta1
		offset is optional.
	"""
	y = [e[0] + offset for e in y]
	v = [e[0] for e in v]
	assert y[0] >= 0
	ret = 0
	if y[0] > 0:
		ret += (1 - v[0] / 2.0 / v0) * (y[0] - 0)
	for i in xrange(1, len(y)):
		if y[i] > delta:
			ret += (1 - (0.99 * v0 + v[i-1]) / 2.0 / v0) * (delta - y[i-1])
			break
		ret += (1 - (v[i] + v[i-1]) / 2.0 / v0) * (y[i] - y[i-1])
	return ret

@xl_func("float[] y, float[] v, float v0, float delta, float offset: float")
def delta2(y, v, v0, delta, offset = 0):
	"""
		given arrays y, v and parameters v0, delta,
		return the momentum thickness delta2
		offset is optional.
	"""
	y = [e[0] + offset for e in y]
	v = [e[0] for e in v]
	assert y[0] >= 0
	ret = 0
	if y[0] > 0:
		ret += v[0] / 2.0 / v0 * (1 - v[0] / 2.0 / v0) * (y[0] - 0)
	for i in xrange(1, len(y)):
		if y[i] > delta:
			ret += (0.99 * v0 + v[i-1]) / 2.0 / v0 * (1 - (0.99 * v0 + v[i-1]) / 2.0 / v0) * (delta - y[i-1])
			break
		ret += (v[i] + v[i-1]) / 2.0 / v0 * (1 - (v[i] + v[i-1]) / 2.0 / v0) * (y[i] - y[i-1])
	return ret

@xl_func("float[] y, float[] v, float v0, float delta, float offset: float")
def delta3(y, v, v0, delta, offset = 0):
	"""
		given arrays y, v and parameters v0, delta,
		return the energy thickness delta3
		offset is optional.
	"""
	y = [e[0] + offset for e in y]
	v = [e[0] for e in v]
	assert y[0] >= 0
	ret = 0
	if y[0] > 0:
		ret += v[0] / 2.0 / v0 * (1 - (v[0] / 2.0 / v0) ** 2) * (y[0] - 0)
	for i in xrange(1, len(y)):
		if y[i] > delta:
			ret += (0.99 * v0 + v[i-1]) / 2.0 / v0 * (1 - ((0.99 * v0 + v[i-1]) / 2.0 / v0) ** 2) * (delta - y[i-1])
			break
		ret += (v[i] + v[i-1]) / 2.0 / v0 * (1 - ((v[i] + v[i-1]) / 2.0 / v0) ** 2) * (y[i] - y[i-1])
	return ret

@xl_func("float lookUpValue, float[] arr: string")
def index2(lookUpValue, arr):
	rowNumber = 1
	for row in arr:
		for i in xrange(len(row)):
			if row[i] == lookUpValue:
				colNumber = i+1;
				return "%d,%d" % (rowNumber, colNumber)
		rowNumber += 1
	return	"cannot locate value"