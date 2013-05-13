# coding: utf-8
#!/usr/bin/python
# -*- coding: utf-8 -*-
# Copyright 2013 Jason Blanks. All Rights Reserved.
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
#    http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.
#
#*************Requires the redemption library to be installed!*******************************
#
#

import sys, stat, os, xlwt, re
import win32com.client

def RecurseFolderCount(folder, pst_path, TotalCount, rdoSession, msgCount):
  print "current folder: "+str(folder)
	print "Child Folders: "+str(folder.Folders)
	print "Count Here: "+str(TotalCount)
	print "Item Count here: "+str(folder.Items.Count)+"\n"
	TotalCount = TotalCount + folder.Items.Count   
	for subFolder in folder.Folders:
		TotalCount = RecurseFolderCount(subFolder, pst_path, TotalCount, rdoSession, msgCount)
	return TotalCount

msgCount = 0
directory = "."
extension = ".pst"
TTemp = 0
OverAllSize=0
TotalCount = 0
OverAllCount = 0
mbSize = 0
gbSize=0
pstpath = '.'
wbk = xlwt.Workbook()
sheet = wbk.add_sheet('sheet 1')
s = 0



for (path, dirs, files) in os.walk(pstpath):
	print "working.."
	filelist=[file for file in files if file.lower().endswith(extension)]
	
	for d in filelist:
		try:
			size = os.path.getsize(path + "\\" + d)
			#size /= 1024*1024.0
			#gsize /= 1024*1024*1024.0
			OverAllSize = size + OverAllSize
			rdoSession = win32com.client.Dispatch("Redemption.RDOSession")
			rdoSession.LogonPstStore(path + "\\" + d)
			pstStore = rdoSession.Stores.DefaultStore
		except Exception as e:
			s = s + 1
			sheet.write(s,0,path + "\\" +d)
			sheet.write(s,1,str("Unable to open"))
			sheet.write(s,2,str("Unable to open"))
			continue
			print e
		rootFolder = pstStore.IPMRootFolder
		#print "working on " + d
		#print "RootCount: " + str(rootFolder.Items.Count)
		TotalCount = 0
		TotalCount = RecurseFolderCount(rootFolder, pstStore.PstPath, TotalCount, rdoSession, msgCount)
			
		OverAllCount = OverAllCount + TotalCount
		gbSize = OverAllSize / (1024*1024*1024.0)
		mbSize = OverAllSize / (1024*1024.0)

		print "End Count: " + str(TotalCount)
		s = s + 1
		filegbSize = size / (1024*1024*1024.0)
		filembSize = size / (1024*1024.0)
		sheet.write(s,0,path + "\\" +d)
		sheet.write(s,1,str(TotalCount))
		sheet.write(s,2,str(size))
		sheet.write(s,3,str(filembSize))
		sheet.write(s,4,str(filegbSize))
		#out.write(path + "\\" + d +"\t"+str(TotalCount)+"\t"+str(size)+"\n")
sheet.write(0,0,"File")
sheet.write(0,1,"MSG Count")
sheet.write(0,2,"Size in Bytes")
sheet.write(0,3,"Size in MB")
sheet.write(0,4,"Size in GB")


sheet.write(s+3,1,"Total MSG Count")
sheet.write(s+4,1,str(OverAllCount))
sheet.write(s+3,2,"Total Size Bytes")
sheet.write(s+4,2,str(OverAllSize))
sheet.write(s+6,2,"Total Size MB")
sheet.write(s+7,2,str(mbSize))
sheet.write(s+9,2,"Total Size GB")
sheet.write(s+10,2,str(gbSize))
#out.write("\nTotal msg count: "+str(OverAllCount)+"\tTotal Data Size: "+str(OverAllSize)+"bytes\t"+str(mbSize)+"mb\t"+str(gbSize)+"gb")
wbk.save('report.xls')
