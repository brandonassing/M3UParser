import sys
import xlwt
import os

#media objects
class media():
	def __init__(self, info, path):
		self.info = info
		self.path = path

#file parser
def parseM3U(file):
	try:
		assert(type(file) == '_io.TextIOWrapper')
	except AssertionError:
		file = open(file,'r')

		line = file.readline()
	if not line.startswith('#EXTM3U'):
		return
	list = []
	mediaObj = media(None, None)
	for line in file:
		line = line.strip()
		if line.startswith('#EXTINF:-1,'):
			info = line.replace("#EXTINF:-1,","")
			mediaObj=media(info.encode('ascii', 'ignore'),None)
		elif len(line) > 0:
			mediaObj.path = line
			list.append(mediaObj)
			mediaObj=media(None, None)
	file.close()
	return list

#MAIN
m3ufile=sys.argv[1]

try:
	os.remove('StreamInfo.xls')
except FileNotFoundError:
	print('StreamInfo.xls not found')
wb = xlwt.Workbook()
ws = wb.add_sheet('Stream Info')
styleTitle = xlwt.easyxf('font: name Times New Roman, color-index black, bold on')
ws.write(0,0,"Info", styleTitle)
ws.write(0,1,"Path", styleTitle)

file = open("StreamInfo.txt", "w")
file.truncate()

list = parseM3U(m3ufile)
for i,mediaObj in enumerate(list):
	print (mediaObj.info, mediaObj.path)

	ws.write(i + 1,0, mediaObj.info.decode("utf-8"))
	ws.write(i + 1,1, mediaObj.path)

	file.write(mediaObj.info.decode("utf-8"))
	file.write('\n')
	file.write(mediaObj.path)
	file.write('\n')
file.close()
wb.save('StreamInfo.xls')

