Option Explicit

Dim objShell
Dim sapApplication
Dim SapPath
Dim FilePath
Dim LogPath
Dim FileName
Dim SapGuiAuto
Dim Connection
Dim Session
Dim SapConnectionName 
Dim SapUsr
Dim User
Dim Pwd
Dim CurrentDate
Dim FirstDate
Dim LastDate
Dim ReceiverAddress
Dim ObjPassword
Dim Logged

Logged = False
SapPath = """C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"""
SapUsr = ""'##################################
Pwd = ""'################################
FileName = "export.xlsx"
SapConnectionName = "01 - ECC - Production - EP0"	
ReceiverAddress = ""
User =  InputBox("Digite a matrícula do usuário logado no computador: ")'<<<<<<<<<<<<<<<<<< 
FilePath = "C:\Users\" & User & "\Documents"
LogPath = "C:\Users\" & User & "\Documents\logs" 
FirstDate = InputBox("Digite a primeira data: (Formato 00/00/0000) ")
LastDate = InputBox("Digite a segunda data: (Formato 00/00/0000) ")

If Len(FirstDate) = 0 Or Len(User) = 0 Or Len(LastDate) = 0 Then
	MsgBox("Dados Inválidos!")
	WScript.Quit
End If

LastDate = GetFormatDate(LastDate)
FirstDate = GetFormatDate(FirstDate)

Call Main()

Sub Main()
	OpenSap
	If Logged = True Then
		GetEvents
		CloseSap
	End If	
End Sub

Sub OpenSap()
	'Open the sap logon screen
	On Error Resume Next
	If Not IsObject(SapApplication) Then
		Set ObjShell = CreateObject("WScript.Shell")
		ObjShell.Run SapPath 	
		WScript.sleep 4000	
		Set ObjShell = Nothing	
	End If	
	
	If Err.Number <> 0 Then
		Log Now() & " Erro ao abrir o SAP (" & Err.Number & " ) : " & Err.Description 
		WScript.Quit
	End If 
	
	If Not IsObject(SapApplication) Then	
		Set SapGuiAuto = GetObject("SAPGUI") 'Get the SAP GUI Scripting object (like if you was typing in init menu from windows)	
		Set SapApplication = SapGuiAuto.GetScriptingEngine() 'Get the currently running SAP GUI	
	End If		
	
	If Not IsObject(Connection) Then
	   Set Connection = SapApplication.OpenConnection(SapConnectionName, True) 'Get the first system that is currently connected	
	End If

	If Not IsObject(Session) Then
	   Set Session = Connection.Children(0) 'Get the first Session (window) on that connection	
	End If

	If IsObject(WScript) Then
	   WScript.ConnectObject Session,     "on"
	   WScript.ConnectObject SapApplication, "on"
	End If
	Session.findById("wnd[0]").maximize 
	
	If Len(SapUsr) > 0 And Len(Pwd) > 0 Then		
		Session.findById("wnd[0]/usr/txtRSYST-BNAME").text = SapUsr
		Session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = Pwd		
		Session.findById("wnd[0]").sendVKey 0	
		
		If(Session.ActiveWindow.Name = "wnd[1]") Then					
			If(InStr(session.findById("wnd[1]/usr/txtMULTI_LOGON_TEXT").text, SapUsr) > 0 And InStr(session.findById("wnd[1]/usr/txtMULTI_LOGON_TEXT").text, "logon") > 0) Then
				MsgBox("O usuario " & SapUsr & " já estava logado! Tente novamente mais tarde.")		
				session.findById("wnd[1]").close()
				session.findById("wnd[0]").close()				
				Logged = False
			Else
				WScript.Quit
			End If
		Else
			Logged = True
		End If
	Else 
		Log "Error on length of User or Password!"
		MsgBox("Erro no usuario ou senha!")
		WScript.Quit
	End If
	
End Sub

Sub CloseSap
	On Error Resume Next
	If (Logged = True) Then
		Session.findById("wnd[0]").close() 'close connection screen	
		Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press		
	End If
	Err.Clear
	
	Session = Empty
	Connection = Empty
	SapApplication = Empty
	SapGuiAuto = Empty
	ObjShell = Empty
	Logged = False
End Sub

'Get all events from month and export to a spreadsheet  in FilePath Variable
Sub GetEvents()		
	
	'Open the screen for event consultants
	Session.findById("wnd[0]").maximize
	Session.findById("wnd[0]/tbar[0]/okcd").text = "iw29"	'/0 new window
	Session.findById("wnd[0]").sendVKey 0
	Session.findById("wnd[0]").sendVKey 17 'shift + f5 equivalent
	
	'Pop up 
	Session.findById("wnd[1]/usr/txtV-LOW").text = "notasccmee"
	Session.findById("wnd[1]/usr/txtENAME-LOW").text = "" 'delete matricule
	Session.findById("wnd[1]/usr/txtENAME-LOW").setFocus
	Session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 0
	Session.findById("wnd[1]/tbar[0]/btn[8]").press
	
	'Back in the iw29 screen
	
	Session.findById("wnd[0]/usr/ctxtDATUV").text = FirstDate 'init data
	Session.findById("wnd[0]/usr/ctxtDATUB").text = LastDate'final data
	Session.findById("wnd[0]/usr/ctxtDATUB").setFocus
	Session.findById("wnd[0]/usr/ctxtDATUB").caretPosition = 10	
	Session.findById("wnd[0]").sendVKey 8 'f8 equivalent to make the search
	
	If(InStr(Session.findById("wnd[0]/sbar").text, "Nenhum objeto selecionado") = 0) Then	
	
		'Select All cells
		Session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell -1,""
		Session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectAll
		Session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu
		'Choose export to XLS
		Session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItem "&XXL"	
		Session.findById("wnd[1]/tbar[0]/btn[0]").press	
		Session.findById("wnd[1]/usr/ctxtDY_PATH").text = FilePath
		Session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = FileName	
		Session.findById("wnd[1]/tbar[0]/btn[11]").press 'substitute file
		Session.findById("wnd[0]/tbar[0]/btn[15]").press 'close screen of excell export	
		
		'Close instance of Excel openned by SAP
		ClosePlan(FileName)
		CloseExcel()
		'put detailed description in spreadsheet
	End If
	
	Session.findById("wnd[0]/tbar[0]/btn[15]").press 'close	iw29 screen	
	
	If Err.Number <> 0 Then
		Log "Erro ao obter notas :(" & Err.Number & ") : " & Err.Description
		CloseSap
		WScript.Quit
	Else	
		On Error Resume Next
		GetDetailedDescription()
		If Err.Number = 0 Then
			Requisition(FilePath & "\" & FileName)
		Else
			Log "Erro ao obter notas :(" & Err.Number & ") : " & Err.Description
			CloseSap
			WScript.Quit
		End If
	End If

	'msgbox "Finalized!"	
End Sub

'get detailed description for each note in spreadsheet in variable FileName
Sub GetDetailedDescription()
		
	Dim objExcell
	Dim objWorkBook	
	Dim objSheet
	Dim count 
	Dim i
	Dim lastRowCount
	Dim noteColumn
	Dim completeDescriptionColumn	
	Dim field	
			
	Session.findById("wnd[0]/tbar[0]/okcd").text = "iw23" 'open iw23 screen
	Session.findById("wnd[0]").sendVKey 0
	
	If Err.Number <> 0 Then
		Log "Erro: Could Not open the iw23 screen (" & Err.Number & ") : " & Err.Description
		CloseSap
		WScript.Quit
	End If
	
	'open spreadsheet and get detailed description for each note
	If Not IsObject(objExcell) Then
		Set objExcell = CreateObject("Excel.Application")
	End If
	
	If Err.Number <> 0 Then
		Log "Erro: Could Not open excel (" & Err.Number & ") : " & Err.Description
		CloseSap
		WScript.Quit
	End If
	
	If Not IsObject(objWorkBook) Then
		Set objWorkBook = objExcell.Workbooks.Open(FilePath & "\" & FileName)
	End If
	
	If Err.Number <> 0 Then
		Log "Erro: Could Not open the file " & (FilePath & "\" & FileName) & " ==> (" & Err.Number & ") : " & Err.Description
		CloseSap
		WScript.Quit
	End If
	
	objExcell.Application.DisplayAlerts = False
	objExcell.Application.Visible = True	
	
	If Not IsObject(objSheet) Then
		Set objSheet = objWorkBook.sheets(1)
	End If
	
	count = 2 ' First Line of lecture, line 1 is the worksheet header
	lastRowCount = objSheet.UsedRange.Rows.Count 'get total number of rows
	noteColumn = 6 'number of column that contains note	
	completeDescriptionColumn = 25 'number of column to put complete description
	objExcell.Cells(1, completeDescriptionColumn).Value = "Descrição Completa"
	
	While count <= lastRowCount
		Dim nota
		Dim description	
		Dim qtdLines
						
		nota = (objExcell.Cells(count, noteColumn).Value)
		Session.findById("wnd[0]/usr/ctxtRIWO00-QMNUM").text = nota 'put note in text box
		Session.findById("wnd[0]").sendVKey 0 'enter
		description = ""
		Err.Clear				
		
		On Error Resume Next
		set field = Session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_2:SAPLIQS0:7715/cntlTEXT/shellcont/shell")		
		
		If Err.Number = 0 Then
			qtdLines = field.LineCount
			
			For i = 1 To qtdLines
				description = description & field.GetLineText(i) & Chr(10)
			Next			
			objExcell.Cells(count, completeDescriptionColumn).Value = description
			Session.findById("wnd[0]/tbar[0]/btn[3]").press 'get back to search description for next note			
		Else			
			objExcell.Cells(count, completeDescriptionColumn).Value = "---"			
			Err.Clear
		End If	
			
		count = count + 1
	Wend	
	
	objExcell.ActiveWorkBook.SaveAs(FilePath & "\" & FileName) 
	objExcell.ActiveWorkBook.Close()	
	objExcell.Application.Quit
	objSheet = Empty
	objWorkBook = Empty
	objExcell = Empty	
	
End Sub

Sub CloseExcel()
	Dim objWMIService, objProcess, colProcess
	
	Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\.\root\cimv2")
	Set colProcess = objWMIService.ExecQuery("Select * from Win32_Process Where Name = " & "'EXCEL.EXE'")
	
	For Each objProcess in colProcess
		objProcess.Terminate()
	Next
	
End Sub

Sub ClosePlan(name)
	Dim LvWorkbook
	Set LvWorkbook = GetObject("C:\Users\" & User & "\Documents\" & name)
	LvWorkbook.Close
End Sub


Function GetCurrentDate()
	Dim lvDt 
	lvDt = Date 'Current date
	
	GetCurrentDate = lvDt
End Function

Function GetFormatDate(dt)
	Dim lvDay, lvMonth, lvYear
	
	If CInt(Day(dt)) < 10 Then
		lvDay = "0" & Day(dt)
	Else
		lvDay = Day(dt)
	End If
	
	If CInt(Month(dt)) < 10 Then
		lvMonth = "0" & Month(dt)
	Else
		lvMonth = Month(dt)
	End If
	
	lvYear = Year(dt)
	
	GetFormatDate = lvDay & "." & lvMonth & "." & lvYear
End Function

Function GetFirstDayOfCurrentMonth(dt)
	Dim lvDay, lvMonth, lvYear
	
	lvDay = "01"
	
	If CInt(Month(dt)) < 10 Then
		lvMonth = "0" & Month(dt)
	Else
		lvMonth = Month(dt)
	End If
	
	lvYear = Year(dt)
	
	GetFirstDayOfCurrentMonth = lvDay & "." & lvMonth & "." & lvYear
End Function

Function GetLogFormatDate(dt)
	Dim lvDay, lvMonth, lvYear, lvHour
	
	If CInt(Day(dt)) < 10 Then
		lvDay = "0" & Day(dt)
	Else
		lvDay = Day(dt)
	End If
	
	If CInt(Month(dt)) < 10 Then
		lvMonth = "0" & Month(dt)
	Else
		lvMonth = Month(dt)
	End If
	
	If CInt(Hour(dt)) < 10 Then
		lvHour = "0" & Hour(dt)
	Else
		lvHour = Hour(dt)
	End If
		
	lvYear = Year(dt)
	
	GetLogFormatDate = lvDay & "_" & lvMonth & "_" & lvYear & "_" & lvHour
End Function

Sub Log(text)
	Dim objFSO
	Dim objFile
	
	Set objFSO = CreateObject("Scripting.FileSystemObject")	
	Set objFile = objFSO.OpenTextFile(LogPath & "\log_" & GetLogFormatDate(Now()) & ".txt", 8, True)
	objFile.WriteLine(Now() & " ===> " & text)
	
	Set objFile = Nothing
	set objFSO = Nothing
End Sub

'send the file in datapath to controller defined in ReceiverAddress variable
Sub Requisition(dataPath)
	
	On Error Resume Next
	Dim objectStream	
	'Getting file
	Set objectStream = CreateObject("ADODB.Stream")	
	objectStream.Type = 1	
	objectStream.Mode = 3
	objectStream.Open
	objectStream.LoadFromFile dataPath		
	
	Dim xmlHttp
	
	Set xmlHttp = CreateObject("MSXML2.XMLHTTP")
	xmlHttp.Open "POST", ReceiverAddress, False
	xmlHttp.setRequestHeader "Content-Type", "multipart/form-data; boundary=---------------------------19951207"	
	xmlHttp.send objectStream.Read()
	objectStream.Close
	
	Set objectStream = Nothing
	Set xmlHttp = Nothing
	
	If Err.Number <> 0 Then
		Log "Erro: Could Not do the post requisition " & (FilePath & "\" & FileName) & " ==> (" & Err.Number & ") : " & Err.Description
		WScript.Quit
	End If	

End Sub