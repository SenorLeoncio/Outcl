#pip install pywin32

# Make sure that cache is generated with c:\python36\Lib\site-packages\win32com\client
# ( http://www.icodeguru.com/WebServer/Python-Programming-on-Win32/ch12.htm )

# Useful links:
#	https://docs.microsoft.com/en-us/office/vba/api/Outlook.MailItem
#	https://docs.microsoft.com/en-us/office/vba/api/Outlook.Items.Restrict


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



class Reporter :

	def __init__( self, strOutlookFolder='from systems', strExcelFilename='Dsh1.xlsx' ) :

		nsOutlook = win32com.client.Dispatch("Outlook.Application").GetNameSpace("MAPI")

		self.folderFromSystems = nsOutlook.GetDefaultFolder(6).Folders('Auto').Folders(strOutlookFolder)

		appExcel = win32com.client.Dispatch('Excel.Application')
		appExcel.Visible = True

		self.workbookExcel = appExcel.Workbooks.Open( os.getcwd() + '\\' + strExcelFilename )


	def LoadConfig(self) :

		self.dConfig = {}

		ws = self.workbookExcel.Worksheets('Config')
		row = 2
		while True :
			strName  = ws.Cells(row,1).Value
			strValue = ws.Cells(row,2).Value
			if strName == None :
				break
			self.dConfig[strName] = strValue
			row += 1


	def LoadDates(self) :

		self.dDatesColumns	= {}

		col				= 2
		ws				= self.workbookExcel.Worksheets('Dashboard')
		while True :
			dtDT = ws.Cells(2,col).Value
			if dtDT == None :
				break
			strDT = dtDT.strftime('%Y-%m-%d') # because Excel draws '2018-07-12' but returns .Value = '2018-07-12 00:00:00+00:00' (datetime)
			self.dDatesColumns[strDT] = col
		#	ws.Columns(col).Hidden = True
			col += 1

		self.dDatesToProcess = {}
		for i in range( 0, int(self.dConfig['ProcessLastNDays']) ) : # When ProcessLastNDays == 1 - it means only today, this will make only iteration i = 0
			dtDate = datetime.datetime.now() - datetime.timedelta(days=i)
			self.dDatesToProcess[dtDate.strftime('%Y-%m-%d')] = dtDate

		'''
		dDatesToProcess :
		{	'2018-10-07': datetime.datetime(2018, 10, 7, 13, 40, 6, 401248),
 			'2018-10-08': datetime.datetime(2018, 10, 8, 13, 40, 6, 401248)
 		}
		'''

		for (k,v) in sorted(self.dDatesToProcess.items()) :
			if k in self.dDatesColumns :
				continue
			self.dDatesColumns[k] = col
			col += 1


	def ReadMessages(self) :

		messages = self.folderFromSystems.Items
		messages = messages.Restrict("[From] = 'interact.connect.support@responsys.com'")

		dtCutoff = datetime.datetime.now() - datetime.timedelta(days=self.dConfig['ProcessLastNDays']-1)
		strReceivedTimeCutoff = "[ReceivedTime] >= '{}/{}/{} 00:00'".format(dtCutoff.month, dtCutoff.day, dtCutoff.year)
		print( 'Applying filter: {}'. format(strReceivedTimeCutoff) )
		messages = messages.Restrict(strReceivedTimeCutoff)

		msgCount = len(messages)
		print( 'Number of emails qualifying the above filter: {}'. format(msgCount) )

		maxCount = min(msgCount,int(self.dConfig['MaxNumberOfEmails']))
		print( 'Final number of emails to be processed: {}'. format(maxCount) )

		#self.msg = messages
		self.lMessage = []

		reCompile = re.compile( '^Error(.+)$', flags=re.MULTILINE )

		i = 0
		for m in list (messages)[ 0:maxCount ]  :
			(strNotification, strConnectJob, strAccount, strResult) = m.Subject.split(' - ')

			strResult = 'S' if strResult == 'Task Successful' else 'F' if strResult == 'Task Failed' else 'U'

			try:
				strErrorText = reCompile.findall(m.Body)[-1]
			except:
				strErrorText = None

			dtReceivedDT = m.ReceivedTime.replace(tzinfo=None)
			strReceivedDTShort = dtReceivedDT.strftime('%Y-%m-%d')

			t = ( strConnectJob, strReceivedDTShort, dtReceivedDT, strResult, strErrorText )

			self.lMessage.append( t )

			print( i, end=' ', flush=True )
			i += 1

		#print( '\n{} distinct connect jobs parsed from emails'. format(len(self.lMessage)) )

		print( '\nNow sorting lMessage...' )
		self.lMessage.sort( key=lambda x: x[2] )


	def	ParseMessagesIntoCJ(self) :

		self.dCJ = {}
		for ( strConnectJob, strReceivedDTShort, dtReceivedDT, strResult, strErrorText ) in self.lMessage :

			# lMessage:
			'''	( strConnectJob, strReceivedDTShort, dtReceivedDT, strResult, strErrorText )
			[
				('DOI_TrueFit_CL_UL',	'2018-10-18', '2018-10-18 19:13:49', 'S', None),
				('SI_MCOM_POS_PET_UL',	'2018-10-18', '2018-10-18 18:58:38', 'S', None),
				('BS_POS_CL_UL',		'2018-10-18', '2018-10-18 18:56:23', 'F', ':JobId.RunId:... error msg'),
			  	...
			]
			'''

			k = (strConnectJob, strReceivedDTShort)
			if k not in self.dCJ :
				self.dCJ[k] = [ strResult, strErrorText, 1 ]
			else :
				self.dCJ[k][0] += strResult
				
				if self.dCJ[k][1] != None and strErrorText != None :
					self.dCJ[k][1] += '\n'
					self.dCJ[k][1] += strErrorText
				elif self.dCJ[k][1] == None and strErrorText != None :
					self.dCJ[k][1] = strErrorText

			# dCJ:
			'''
			{
			 	('SI_PET_UL', '2018-10-17') : ['SSSSSSSSSSSSS', None],
			 	('SI_PET_UL', '2018-10-18') : ['SSSSSSSSSSSS', None],
			 	('SI_MCOM_POS_PET_UL', '2018-10-18') : ['FFFFFFFFFSSSSSSSSSSS', ':JobId.RunId:\xa071688.3064995....\n....']
			}
			'''


	def LoadCJParameters(self) :

		self.dCJParameters	= {}
		ws	= self.workbookExcel.Worksheets('Inventory')

		row = 2
		while True :
			strCJ = ws.Cells(row,1).Value
			#print( 'Reading CJ: {}'. format(strCJ) )
			if strCJ == None :
				break
			#print('Reading parameters for CJ = {}'. format(strCJ))

			self.dCJParameters[strCJ] = {}
			col = 2
			while True :
				col += 1
				#print('Reading parameters for the col = {}'. format(col))

				strParamDesc = ws.Cells(1,col).Value
				if strParamDesc == None :
				# if we reached the empty column without values...
					break

				strParam = ws.Cells(row,col).Value
				if strParam == None :
				# if the parameter value is not specified...
					continue

				if col == 3 :
					self.dCJParameters[strCJ]['C'] = strParam

				if col == 4 :
					self.dCJParameters[strCJ]['D'] = strParam

				if col == 5 :
					self.dCJParameters[strCJ]['E'] = strParam

			row += 1

		#dCJParameters:
		'''
		{	'DOI_TrueFit_CL_UL': {'C': '1', 'D': 'rgbAquamarine'},
			'Product_Rec_MEN_PET_Import': {}
		}
		e.g. :
			dCJParameters['DOI_TrueFit_CL_UL']['C'] == '1'
		'''


	def PopulateMissingCJToExcel(self) :

		ws = self.workbookExcel.Worksheets('Dashboard')
		row = 3

		lCJinExcel = []

		while True :
			strCJ = ws.Cells(row,1).Value
			if strCJ == None :
				break

			lCJinExcel.append(strCJ)

			row += 1

		lCJParsedFromOutlook = [ k[0] for k in self.dCJ ]

		self.lCJMissing = set(lCJParsedFromOutlook) - set(lCJinExcel)

		strComment = 'Added ' + datetime.datetime.now().strftime('%Y-%m-%d')

		for strCJ in self.lCJMissing :

			ws.Cells(row,1).Value = strCJ
			ws.Cells(row,1).AddComment( strComment )
			ws.Cells(row,1).Comment.Shape.TextFrame.AutoSize = True

			row += 1


	def PopulateExcel(self) :

		ws = self.workbookExcel.Worksheets('Dashboard')
		ws.Select

		#	Populate column headers for new dates:
		for (k,v) in sorted(self.dDatesToProcess.items()) :
			col = self.dDatesColumns[k]
			ws.Cells(1,col).Value = v.strftime('%A') # Monday, Tuesday, etc...
			ws.Cells(2,col).Value = k


		row = 3
		# iterate through all rows(connect jobs)
		while True :
			strCJ = ws.Cells(row,1).Value
			if strCJ == None :
				break
			
			#col = 2

			for strDT in self.dDatesToProcess :
				col = self.dDatesColumns[strDT]

				if (strCJ, strDT) in self.dCJ :
					# if we have something to populate into this Excel cell:

					self.dCJ[ strCJ, strDT ][2] = 0

					# Clear value and reset highlight
					ws.Cells(row,col).Value = None
					ws.Cells(row,col).Interior.Pattern = -4142
					ws.Cells(row,col).ClearComments()

					ws.Cells(row,col).Value = self.dCJ[ strCJ, strDT ][0]

					# display comments (connect job error message):
					if self.dCJ[ strCJ, strDT ][1] != None :
						ws.Cells(row,col).AddComment( self.dCJ[ strCJ, strDT ][1] )
						ws.Cells(row,col).Comment.Shape.TextFrame.AutoSize = True
						#if col == max(self.dDatesColumns.values()) :
						#	ws.Cells(row,col).Comment.Visible = True

					if strCJ in self.dCJParameters :
						dParameter = self.dCJParameters[strCJ]

						if 'C' in dParameter and dParameter['C'] == 'Y' :
							if 'S' not in self.dCJ[ strCJ, strDT ][0] :
								ws.Cells(row,col).Interior.Color = self.dConfig['HighlightFailure']
							
						if 'D' in dParameter :
							ws.Cells(row,col).Interior.Color = dParameter['D']

				else :
					# if we do not have anything to populate into this cell... we may still want to highlight the cell if respective param is configured

					if strCJ in self.dCJParameters :
						dParameter = self.dCJParameters[strCJ]

						if 'E' in dParameter and dParameter['E'] == 'Y' :
								ws.Cells(row,col).Interior.Color = self.dConfig['HighlightFailure']

			row += 1

		for (k,v) in self.dDatesToProcess.items() :
			ws.Columns(self.dDatesColumns[k]).AutoFit()


	def SaveExcelWorkbook(self) :
		self.workbookExcel.Save()



#if __name__ == '__main__' :

r = Reporter()

r.LoadConfig()

r.LoadDates()

r.ReadMessages()

r.ParseMessagesIntoCJ()

r.LoadCJParameters()

r.PopulateMissingCJToExcel()

r.PopulateExcel()

r.SaveExcelWorkbook()
