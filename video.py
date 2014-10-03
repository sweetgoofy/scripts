from __future__ import division
from pyxll import xl_func
from pyxll import xl_macro
import re
@xl_func("float dist: float")
def dScale(dist):
	"""Given measured distance, output scale as actual dist / measured dist (cm/cm)"""
	return (1/dist)

@xl_func("string t1, string t2, int fps: float")
def dt(t1, t2, fps):
	"""Given two time strings and fps, calculate the actual time from playback time.
		t1 has to be less than t2.
	"""
	time1 = t1.split(":")
	h1 = int(time1[0])
	m1 = int(time1[1])
	s1 = int(time1[2])
	ms1 = int(time1[3])

	time2 = t2.split(":")
	h2 = int(time2[0])
	m2 = int(time2[1])
	s2 = int(time2[2])
	ms2 = int(time2[3])

	return ((ms2-ms1)+(s2-s1)*30+(m2-m1)*1800+(h2-h1)*108000)/fps