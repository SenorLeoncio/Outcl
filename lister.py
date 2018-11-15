#pip install pywin32

# Make sure that cache is generated with c:\python36\Lib\site-packages\win32com\client
# ( http://www.icodeguru.com/WebServer/Python-Programming-on-Win32/ch12.htm )


import win32com.client

import datetime
import os
import re

class myLog(object) :
	prev_dt = datetime.datetime.now()
	curr_dt = datetime.datetime.now()
	i = 0

	def click(self) :
		self.prev_dt = self.curr_dt
		self.curr_dt = datetime.datetime.now()
		self.i += 1
		print( '-- myLog entry #{} elapsed: {}sec --'. format(self.i, (self.curr_dt-self.prev_dt).total_seconds()) )



def fnOpenOutlook() :
	ns = win32com.client.Dispatch("Outlook.Application").GetNameSpace("MAPI")
	return ns



nsOutlook	= fnOpenOutlook()

inboxFolder = nsOutlook.GetDefaultFolder(6)
fromSystemsFolder = inboxFolder.Folders('Auto').Folders('from systems')
messages = fromSystemsFolder.Items
messages = messages.Restrict("[From] = 'interact.connect.support@responsys.com'")

dtCutoff = datetime.datetime.now() - datetime.timedelta(days=1)
strReceivedTimeCutoff = "[ReceivedTime] >= '{}/{}/{} 00:00'".format(dtCutoff.month, dtCutoff.day, dtCutoff.year)
print( 'Applying filter: {}'. format(strReceivedTimeCutoff) )
messages = messages.Restrict(strReceivedTimeCutoff)

msgCount = len(messages)
print( 'Number of emails qualifying the above filter: {}'. format(msgCount) )

maxCount = min(msgCount,30000)
print( 'Final number of emails to be processed: {}'. format(maxCount) )

dMessage = {}
mylog = myLog()

for i in range(0,maxCount)  :
	(strNotification, strConnectJob, strAccount, strResult) = messages[i].Subject.split(' - ')
	strBody = messages[i].Body

	#print(strBody)
	s = re.search('(Description:\xa0)(.*)\r',strBody)
	strDesc = ''
	if s is not None:
		strDesc = s.group(2).replace('\xa0',' ')

	#print('{}\t{}' .format(strConnectJob, strDesc))

#		t = ( strResult, messages[i].ReceivedTime.strftime('%Y-%m-%d %H:%M:%S') )
	#t = ( strResult, messages[i].ReceivedTime.replace(tzinfo=None) )

	if strConnectJob not in dMessage :
		dMessage[ strConnectJob ] = strDesc
	#	dMessage[strConnectJob].append( t )
	#else :
	#	dMessage[ strConnectJob ] = [ t ]

	print( i, end=' ', flush=True )

for (k,v) in dMessage.items() :
	print('{}\t{}' .format(k,v))

#print( '\n{} distinct connect jobs parsed from emails'. format(len(dMessage)) )
#mylog.click()

#print( 'Now sorting dMessage...' )
#for d in dMessage :
#	dMessage[d].sort( key=lambda x: x[1], reverse=True )

#mylog.click()
#return dMessage


#	Open Excel and read parameters first
#wb			= fnOpenExcelWorkbook()
#dConfig		= fnLoadConfig(wb)

#( dDatesColumns, dNewDates ) = fnLoadDates(wb, dConfig)

#dMessage	= fnLoadMessages(nsOutlook)

#dCJ			= fnParseMessages(dMessage, dConfig)

#dCJParameters = fnLoadCJParameters(wb)
