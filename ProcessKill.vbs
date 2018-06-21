Dim Pgm
Dim ArrPgm() 
Dim ArrIndex
Dim Wmi

Dim Processes
Dim Process
Dim sQuery

ReDim ArrPgm(10) 

ArrPgm(0) = "OUTLOOK.EXE"
ArrPgm(1) = "iexplore.exe"
ArrPgm(2) = "saplogon.exe"				'SAP
ArrPgm(3) = "ASRS.exe"					'IMS
ArrPgm(4) = "HTC.MD.UI.Main.exe"
ArrPgm(5) = "EXCEL.EXE"
ArrPgm(6) = "FilmTestEvaluation.exe"	
ArrPgm(7) = "notepad.exe"
ArrPgm(8) = "chrome.exe"
ArrPgm(9) = "JBrowser5.exe"             '수지 SCM
ArrPgm(10) = "WerFault.exe"				'J.Browser5 

'====================  Process Kill  ========================

on error resume next

for ArrIndex = LBound(ArrPgm) to UBound(ArrPgm)
	Pgm = ArrPgm(ArrIndex) 	
	
	Set Wmi = GetObject("winmgmts:")
	sQuery = "Select * From win32_process Where name = '" & Pgm & "'"
	Set Processes = Wmi.execquery(sQuery)

		For Each Process In Processes
			Process.Terminate
		Next

	Set Wmi = Nothing
next  

Wscript.Quit
