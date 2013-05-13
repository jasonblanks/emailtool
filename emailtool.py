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

"""This file contains a parser for PST files"""

import win32com.client
import time
import sys
import argparse
import getopt
import re, os
import time
#import datetime
from datetime import date



def MSG2PST(folder, msgpath, rdoSession, pstStore):
  def RecurseFolder(CurrentDir, pstCurrent):
		print "I am here."
		pstParrent = pstCurrent
		ParentDir = CurrentDir
		os.chdir(CurrentDir)
		files = [f for f in os.listdir('.') if os.path.isfile(f)]
		#files = [f for f in os.listdir(Dir) if f.lower().endswith('.msg')]
		check()
		try:
			pstCurrent = pstCurrent.Folders.add(CurrentDir)
			print "Current Folder: "+CurrentDir
		except Exception as e:
			pstCurrent = pstCurrent.Folders(CurrentDir)
			print "Current Folder: "+CurrentDir
			#print "made it to exception"
			#print e
			pass
		#pstCurrent = pstCurrent.Folders.add(CurrentDir)
		for file in files:
			try:
				pstRoot = pstStore.IPMRootFolder
				#pstCurrent = pstRoot
				oItem = pstCurrent.Items.Add(0)
				oItem.Import(file,3)
				oItem.CopyTo(pstCurrent)
				oItem.Save()
			except Exception as e:
				print "Error copy file: "+str(file)
				f = open('c:\\\\errorlog.txt','a')
				f.write(str(e)+"\n"+str(file)+"\n") # python will convert \n to os.linesep
				f.close()
				pass
		subFolders = [o for o in os.listdir('.') if os.path.isdir(o)]
		for subFolder in subFolders:
			RecurseFolder(subFolder, pstCurrent)
		os.chdir('..')
		
	os.chdir(msgpath)
	Dirs = [o for o in os.listdir(msgpath) if os.path.isdir(o)]
	files = [f for f in os.listdir('.') if os.path.isfile(f)]
	for file in files:
		try:
			print "yup"
			pstRoot = pstStore.IPMRootFolder
			pstCurrent = pstRoot
			oItem = pstCurrent.Items.Add(0)
			oItem.Import(file,3)
			oItem.CopyTo(pstCurrent)
			oItem.Save()
		except Exception as e:
			print "nope"
			f = open('errorlog.txt','a')
			f.write(e+"\n"+file+"\n") # python will convert \n to os.linesep
			f.close()
			pass
		#pstRoot = pstStore.IPMRootFolder
		#pstCurrent = pstRoot
		#oItem = pstCurrent.Items.Add(0)
		#oItem.Import(file,3)
		#oItem.CopyTo(pstCurrent)
		#oItem.Save()
	for Dir in Dirs:
		pstRoot = pstStore.IPMRootFolder
		pstCurrent = pstRoot
		print "oh yeah"
		RecurseFolder(Dir, pstCurrent)

def PST2MSG(folder, pst_path, rdoSession, thepstStore):
	def nestedPST2MSG(folder, pst_path, rdoSession, thepstStore):
		ExportPath = "psttool"
		check()
		print "entering folder: " + folder.Name
		Fname = str(folder.Name)
		ensure_dir(Fname, ExportPath)
		for item in folder.Items:
			print "2"
			ensure_dir(Fname, ExportPath)
			print "3"
			item.SaveAs(ExportPath+"\\"+ Fname +"\\"+item.EntryID+".msg",3)
			print "4"
		
		for subFolder in folder.Folders:
			nestedPST2MSG(subFolder, pst_path, rdoSession, thepstStore)		
	#print folder.Folders
	for subFolder in folder.Folders:
		if str(subFolder) == "Top of Personal Folders":
			#print subFolder
			nestedPST2MSG(subFolder, pst_path, rdoSession, thepstStore)

def PST2PST(pstStore, folder, NewFolder, TotalCount, rdoSession, rdoNewPSTSession, MsgIds, QC):
	TotalCount = TotalCount + folder.Items.Count
	check()
	for m in MsgIds:
		try:
			new = pstStore.GetMessageFromID(m)
			new.CopyTo(NewFolder)
			QC.append(m)
		except Exception as e:
			file_error = open('file_error.txt', 'a')
			file_error.write(m+"\n")
			file_error.close()
			print e
			print m
	return TotalCount, QC


def ensure_dir(Fname, ExportPath):
	path = ExportPath
	FullPath = os.path.join(path, Fname)
	if not os.path.exists(FullPath):
		os.makedirs(FullPath)

def GetMsgIds(msgIdFile):
	f=open(msgIdFile)
	line = f.readlines()
	f.close()
	return f
def check():
	today = date.today()
	Max = date(2013,10,13)
	Min = date(2013,02,13)
	if Max <= today or today <= Min:
		print "Error"
		sys.exit()

def RecurseFolder(folder, pst_path, TotalCount, TestCount, rdoSession, MsgIds):
	msgCount = 0
	mItr = 0
	TotalCount = TotalCount + folder.Items.Count
	ExportPath = "psttool"
	
	print "entering folder: " + folder.Name
	print "item count = " + str(folder.Items.Count)
	print "folder count = " + str(folder.Folders.Count)
	Fname = str(folder.Name)
	ensure_dir(Fname, ExportPath)
	for item in folder.Items:
		msgCount += 1
		TestCount = TestCount + 1
		#print item.EntryID
		for m in MsgIds:
			if item.EntryID == m:
				print item.EntryID
				ensure_dir(Fname, ExportPath)
				item.SaveAs(ExportPath+"\\"+ Fname +"\\"+item.EntryID+".msg",)
				del MsgIds[mItr]
				mItr = mItr + 1
				break
		mItr = mItr+1
		
	for subFolder in folder.Folders:
		TotalCount, TestCount = RecurseFolder(subFolder, pst_path, TotalCount, TestCount, rdoSession, MsgIds)
	print "Extracted: ", msgCount
	print "testcount: ", TestCount
	return TotalCount, TestCount
	
def RecurseFolderCount(folder, pst_path, TotalCount, rdoSession):
	msgCount = 0
	TotalCount = TotalCount + folder.Items.Count
	
	print "Folder: " + folder.Name +" MSG Count:" + str(folder.Items.Count)
	print "folder count = " + str(folder.Folders.Count)
	for item in folder.Items:
		msgCount += 1
				   
	for subFolder in folder.Folders:
		TotalCount = RecurseFolderCount(subFolder, pst_path, TotalCount, rdoSession)
	return TotalCount

def main(argv):

	dirpath = ''
	ExportArg= ''
	filepath = ''
	exportpath = ''
	msgIdFile = ''
	filters = ''
	search = ''
	pstpath = ''
	TotalCount = 0
	TestCount = 0
	CountArg = 0
	QCArg = ''
	folder = ''
	QC = []
	
	
	#USAGE
	#python email.py -p "D:\Email\Correspondence.pst" -m "MessageIdFile.txt"
	
	## TO DO: Add better error handeling
	try:
		opts, args = getopt.getopt(argv,"hp:m:ce:",["pstpath=","msgid=","count","export="])
	except getopt.GetoptError:
		print
		print 'copyright Jason Blanks, David Nides 2013'
		print 'for questions, comments, etc contact jason.blanks@gmail.com'
		print 'Requries redemption Developer version:\nhttp://www.dimastr.com/redemption/download.htm\n'
		print 'usage:'
		print 'emailtool.py -p <pstpath> -m <MessageID file path> -e pst2pst'
		sys.exit(3)
	for opt, arg in opts:
		if opt == '-h':
			print
			print 'copyright Jason Blanks 2013'
			print 'for questions, comments, etc contact jason.blanks@gmail.com'
			print 'Requries redemption Developer version:\nhttp://www.dimastr.com/redemption/download.htm\n'
			print 'usage:'
			print 'email.py -p <pstpath> -m <MessageID file path> -e pst2pst'
			print '\n\nMandatory arguments: '
			print '	-p, --path		PST file path.'
			print '	-m, --msgid		Message ID file list file.'
			print '	-e, --export		Export type, options are: '
			print '				pst2pst, pst2msg, msg2pst'
			print '	-s, --split		Not implemented yet'

			print 
			print '\n\nOptional arguments: '
			print '	-r, --report		Not implemented yet'
			print '	-qc, --check		Not implemented yet'
			sys.exit()
		elif opt in ("-p", "--pstpath"):
			pstpath = str(arg)
		elif opt in ("-m", "--msgid"):
			msgIdFile = arg
		elif opt in ("-c", "--count"):
			CountArg = True
		elif opt in ("-e", "--export"):
			ExportArg = arg
		#elif opt in ("-QC", "--check"):
		#	QCArg = str(arg)

 
	## RdoSession object	
	## MSGID File
	#try:
	#	print 'opening msg id file'
	#	MsgIds = [l.strip() for l in open('mid.txt')]
	#except Exception as e:
	#	print e

	#try:
	#	Count = RecurseFolderCount(folder, pst_path, TotalCount, rdoSession)

	#except Exception as e:
	#	print e
	#	return
	#
	#rootFolder = pstStore.IPMRootFolder
	#NewRootFolder = newPstStore.IPMRootFolder
	print "parsing"
	

	if msgIdFile:
		MsgIds = [l.strip() for l in open('mid.txt')]
		#TotalCount, TestCount = RecurseFolder(rootFolder, pstStore.PstPath, TotalCount, TestCount, rdoSession, MsgIds)
		#TotalCount = ExportPST(rootFolder, NewRootFolder, TotalCount, rdoSession, rdoNewPSTSession, MsgIds, QC)
		#ExportPST(rootFolder, pstStore.PstPath, TotalCount, rdoSession, MsgIds)
		
	if CountArg:
		TotalCount = RecurseFolderCount(rootFolder, pstStore.PstPath, TotalCount, rdoSession)
		
	if ExportArg:
		if ExportArg == 'pst2pst':
			try:
				try:
					rdoSession = win32com.client.Dispatch("Redemption.RDOSession")
					print "rdoSession object created."
					rdoSession.LogonPstStore(pstpath)
					print "Connected to PST store"
					pstStore = rdoSession.Stores.DefaultStore
					rootFolder = pstStore.RootFolder
				except Exception as e:
					print e
				MsgIds = [l.strip() for l in open('mid.txt')]
				rdoNewPSTSession = win32com.client.Dispatch("Redemption.RDOSession")
				print "Please Enter a File Name to Give Your New PST: \n-> "
				fileName = sys.stdin.readline().strip()
				os.system('cls')
				rdoNewPSTSession.LogonPstStore(str(fileName))
				#rdoNewPSTSession.LogonPstStore("c:\\"+str(fileName))
				newPstStore = rdoNewPSTSession.Stores.DefaultStore
				#rdoNewPSTSession.Logon
				#newPstStore = rdoNewPSTSession.Stores.AddPstStore("c:\\newPSTTest.pst")
			except Exception as e:
				print e
			
			#NewRootFolder = newPstStore.IPMRootFolder
			#cd ..
			os.system('cls')
			print "Please Enter a Folder Name to export MSGs to in the new PST: \n->"
			FName = sys.stdin.readline().strip()
			os.system('cls')
			print "working. . . "
			#call RDOSession/RDOStore.GetMessageFromID
			NewRootFolder = newPstStore.IPMRootFolder.Folders.Add(FName)
			TotalCount, QC = PST2PST(pstStore, rootFolder, NewRootFolder, TotalCount, rdoSession, rdoNewPSTSession, MsgIds, QC)
			
			#OLD QC FUNCTION
			#file_diff = open('QC_Report.txt', 'w')
			#file_diff.write("QC Length: "+str(len(QC))+" MsgIDs Length: "+str(len(MsgIds))+"\n")
			#for linecheck in MsgIds:
			#	if linecheck not in QC:
			#		#print "true"
			#		file_diff.write(linecheck+"\n")
			#file_diff.close()
			#print "QC Length: "+str(len(QC))
			
		elif ExportArg == 'pst2msg':
			print pstpath
			try:
				rdoSession = win32com.client.Dispatch("Redemption.RDOSession")
				print "rdoSession object created."
				rdoSession.LogonPstStore(pstpath)
				print "Connected to PST store"
				thepstStore = rdoSession.Stores.DefaultStore
				rootFolder = thepstStore.RootFolder
				PST2MSG(rootFolder, thepstStore.PstPath, rdoSession, thepstStore)
			except Exception as e:
				print e	

		elif ExportArg == 'msg2pst':
			try:
				rdoSession = win32com.client.Dispatch("Redemption.RDOSession")
				print "Please Enter a File Name to Give Your New PST: \n-> "
				fileName = sys.stdin.readline().strip()
				os.system('cls')
				rdoSession.LogonPstStore(str(fileName))
				#rdoSession.LogonPstStore("c:\\"+str(fileName))
				pstStore = rdoSession.Stores.DefaultStore
				rootFolder = pstStore.RootFolder
				print"."
				MSG2PST(rootFolder, pstpath, rdoSession, pstStore)
			except Exception as e:
				print e	

	
	print ""
	#print "PST Store Name = " + pstStore.Name
	#print "PST Path = " + pstStore.PstPath
	#print "Total Item Count = " + str(TotalCount)
	#print "All done."
	
if __name__ == "__main__":
	# parse command line options
	main(sys.argv[1:])
