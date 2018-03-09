Option Explicit

Dim ObjShell
Dim SapApplication
Dim SapPath
Dim FilePath
Dim LogPath
Dim FileName
Dim SapGuiAuto
Dim Connection
Dim Session
Dim SapConnectionName 
Dim User
Dim SapUsr
Dim Pwd
Dim FirstDate
Dim LastDate
Dim ReceiverAddress
Dim TimeInMinutes
Dim Logged
Dim MonthTarget 
Dim MonthCount

SapPath = """C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"""
'SapPath = """C:\Program Files\SAP\FrontEnd\SAPgui\saplogon.exe""" 'old computers 
SapUsr = ""'################################
Pwd = ""'################################
FileName = "export.xlsx"
SapConnectionName = "01 - ECC - Production - EP0"	
ReceiverAddress = ""
TimeInMinutes = InputBox("Digite o intervalo de execucao do Script: (Valor em Minutos)")
User = InputBox("Digite a matricula do usuario logado no computador (com 01): ")
Logged = False

If Len(User) > 0  And Len(TimeInMinutes) > 0 Then
	Call Main()
Else
	MsgBox("Verifique os Dados!")
	WScript.Quit
End If

Sub Main()
	MonthTarget = 20
	MonthCount = 1
	Do While True
		If MonthCount <> MonthTarget Then
			FirstDate = (GetFormatDate(GetCurrentDate()))
			LastDate = (GetFormatDate(GetCurrentDate()))
			MonthCount = MonthCount + 1
		Else 
			FirstDate = GetFirstDayOfCurrentMonth(GetCurrentDate())
			LastDate = (GetFormatDate(GetCurrentDate()))
			MonthCount = 1
		End If
		
		OpenSap
		If (Logged = True) Then	
			On Error Resume Next		
			GetEvents			
			Err.Clear
		End If
		CloseSap
		WScript.sleep (CInt(TimeInMinutes) * 60000)
	Loop
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
	   Set Session = Connection.Sessions(0) 'Get the first Session (window) on that connection	
	End If

	If IsObject(WScript) Then
	   WScript.ConnectObject Session,     "on"
	   WScript.ConnectObject SapApplication, "on"
	End If
	Session.findById("wnd[0]").maximize 
	
	If Len(User) > 0 And Len(Pwd) > 0 Then
		FilePath = "C:\Users\" & User & "\Documents"
		LogPath = "C:\Users\" & User & "\Documents\logs"
		Session.findById("wnd[0]/usr/txtRSYST-BNAME").text = SapUsr
		Session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = Pwd		
		Session.findById("wnd[0]").sendVKey 0	
		
		If(Session.ActiveWindow.Name = "wnd[1]") Then					
			If(InStr(session.findById("wnd[1]/usr/txtMULTI_LOGON_TEXT").text, SapUsr) > 0 And InStr(session.findById("wnd[1]/usr/txtMULTI_LOGON_TEXT").text, "logon") > 0) Then
				Log "O usuario " & SapUsr & " já estava logado, rodando novamente em " & TimeInMinutes	& " minutos"			
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

Sub GetEvents()		
		
	If Err.Number <> 0 Then
		Log "Erro ao obter notas :(" & Err.Number & ") : " & Err.Description & "(on get events init)"
		WScript.Quit
	End If	
	
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
		On Error Resume Next
		Session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell -1,""
		
		'if is no error, the routine is currently in iw29 and the search returned data
		If Err.Number = 0 Then
			Session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectAll
			Session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu
			'Choose export to XLS
			Session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItem "&XXL"	
			Session.findById("wnd[1]/tbar[0]/btn[0]").press	
			Session.findById("wnd[1]/usr/ctxtDY_PATH").text = FilePath
			Session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = FileName
			Session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 11
			Session.findById("wnd[1]/tbar[0]/btn[11]").press 'substitute file
			Session.findById("wnd[0]/tbar[0]/btn[15]").press 'close screen of excell export			
			
			ClosePlan(FileName)
			CloseExcel()
			'Close instance of Excel openned by SAP
			
		Else	
			Log "Erro ao obter notas :(" & Err.Number & ") : " & Err.Description & "(selectAll)"
			Err.Clear
		End If
	End If			
	
	'in any way, i need close iw29 screen, to enter in transaction for get detailed description of older or new events
	Session.findById("wnd[0]/tbar[0]/btn[15]").press 'close	iw29 screen	
	
	On Error Resume Next
	GetDetailedDescription()		
	
	If Err.Number = 0 Then
		Requisition(FilePath & "\" & FileName)
	Else
		Log "Erro ao obter notas :(" & Err.Number & ") : " & Err.Description & "(GetDetailedDescription)"
		CloseSap
		'WScript.Quit
	End If	
End Sub

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
	
	session.findById("wnd[0]/tbar[0]/okcd").text = "iw23" 'open iw23 screen
	session.findById("wnd[0]").sendVKey 0
	
	If Err.Number <> 0 Then
		Log "Erro: Could Not open the iw23 screen (" & Err.Number & ") : " & Err.Description
		WScript.Quit
	End If
	
	'open spreadsheet and get detailed description for each note
	If Not IsObject(objExcell) Then
		Set objExcell = CreateObject("Excel.Application")
	End If
	
	If Err.Number <> 0 Then
		Log "Erro: Could Not open excel (" & Err.Number & ") : " & Err.Description
		WScript.Quit
	End If
	
	If Not IsObject(objWorkBook) Then
		Set objWorkBook = objExcell.Workbooks.Open(FilePath & "\" & FileName)
	End If
	
	If Err.Number <> 0 Then
		Log "Erro: Could Not open the file " & (FilePath & "\" & FileName) & " ==> (" & Err.Number & ") : " & Err.Description
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
		session.findById("wnd[0]/usr/ctxtRIWO00-QMNUM").text = nota 'put note in text box
		session.findById("wnd[0]").sendVKey 0 'enter
		description = ""
		Err.Clear				
		
		On Error Resume Next
		set field = session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_2:SAPLIQS0:7715/cntlTEXT/shellcont/shell")
		
		If Err.Number = 0 Then	
			qtdLines = field.LineCount
			
			For i = 1 To qtdLines
				description = description & field.GetLineText(i) & Chr(10)
			Next		
		
			objExcell.Cells(count, completeDescriptionColumn).Value = description
			session.findById("wnd[0]/tbar[0]/btn[3]").press 'get back to search description for next note			
		Else
			objExcell.Cells(count, completeDescriptionColumn).Value = "---"
			Err.Clear
		End If		
		count = count + 1
	Wend	
	
	objExcell.ActiveWorkBook.SaveAs(FilePath & "\" & FileName) 
	objExcell.ActiveWorkBook.Close()	
	
	objSheet = Empty
	objWorkBook = Empty
	objExcell = Empty
	
End Sub

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
		CloseSap
		'WScript.Quit
	End If	

End Sub

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