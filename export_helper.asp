<%
dim	tmpemp()
dim	tmpcon()
dim	tmpDates()
dim	tmpmile()
dim	mpmilecap()
dim tmpempADP()
dim tmpmileADP()
dim tmpID3()
dim tmpHrs(), tmpHrsExt()
dim	tmpemp2()
dim	tmpcon2()
dim	tmpDates2()
dim	tmpmile2()
dim	mpmilecap2()
dim myDay(6), tbl(6), myUnits(6), myAC(6), mysun(), rsHrs(6), sundatex(2), myacode()
Function Z_GetDate(sundate, dayname)
		Select Case Ucase(dayname)
      Case "SUN"
          Z_GetDate = sundate
      Case "MON"
          Z_GetDate = DateAdd("d", 1, sundate)
      Case "TUE"
          Z_GetDate = DateAdd("d", 2, sundate)
      Case "WED"
          Z_GetDate = DateAdd("d", 3, sundate)
      Case "THU"
          Z_GetDate = DateAdd("d", 4, sundate)
      Case "FRI"
          Z_GetDate = DateAdd("d", 5, sundate)
      Case "SAT"
          Z_GetDate = DateAdd("d", 6, sundate)
    End Select
End Function

function ACdesc(xxx)
	If xxx = "" Then exit function
	
	Set rsAC = Server.CreateObject("ADODB.RecordSet")
	sqlAC = "SELECT [desc] FROM activity_T WHERE code = " & xxx'strAC(ctr)
	rsAC.Open sqlAC, g_strCONN, 3, 1
	If Not rsAC.EOF Then
		ACdesc = rsAC("desc") 
		'response.write "Xxx: " & xxx & " = " & rsAC("desc")  & "<br>"
	End If
	rsAC.Close
	Set rsAC = Nothing
End function

Function SearchArraysacode(acode)
On Error Resume Next
	SearchArraysacode = -1
	lngMax = UBound(myacode)
	If Err.Number <> 0 Then Exit Function
	
	For lngI2 = 0 to lngMax
		If myacode(lngI2) = acode Then Exit For
	Next
	If lngI2 > lngMax Then Exit Function
	SearchArraysacode = lngI2
End Function

Function GetWorkName(xxx)
	GetWorkName = "N/A"
	Set rsName = Server.CreateObject("ADODB.RecordSet")
	sqlName = "SELECT * FROM Worker_T WHERE [index] = " & xxx
	rsName.Open sqlName, g_strCONN, 3, 1
	If Not rsName.EOF Then
		GetWorkName = rsName("Lname") & """,""" & rsName("fname")
	End If
	rsName.Close
	Set rsName = Nothing
End Function

Function GetCMAgencyID(xxx)
	GetCMAgencyID = 0
	Set rsName = Server.CreateObject("ADODB.RecordSet")
	sqlName = "SELECT cmcid FROM Case_Manager_T WHERE [index] = " & xxx
	rsName.Open sqlName, g_strCONN, 3, 1
	If Not rsName.EOF Then
		GetCMAgencyID = Z_CZero(rsName("cmcid"))
	End If
	rsName.Close
	Set rsName = Nothing
End Function

Function GetCMAgency(xxx)
	GetCMAgency = "N/A"
	Set rsName = Server.CreateObject("ADODB.RecordSet")
	sqlName = "SELECT CMCname FROM CaseMngmt_T WHERE [CMCid] = " & xxx
	rsName.Open sqlName, g_strCONN, 3, 1
	If Not rsName.EOF Then
		GetCMAgency = rsName("CMCname")
	End If
	rsName.Close
	Set rsName = Nothing
End Function

Function GetCMName(xxx)
	GetCMName = "N/A"
	If xxx = "" Then Exit Function
	Set rsCode = Server.CreateObject("ADODB.RecordSet")
	sqlCode = "SELECT lname, fname From Case_Manager_t WHERE [index] = " & xxx
	rsCode.Open sqlCode, g_strCONN, 3, 1
	If Not rsCode.EOF Then GetCMName = rsCode("lname") & """,""" & rsCode("fname")
	rsCode.Close
	Set rsCode = Nothing
End Function

Function GetCMAdr(xxx)
	GetCMAdr = "N/A"
	If xxx = "" Then Exit Function
	Set rsCode = Server.CreateObject("ADODB.RecordSet")
	sqlCode = "SELECT CMCaddr, CMCcity, CMCstate, CMCzip From CaseMngmt_T WHERE [CMCid] = " & xxx
	rsCode.Open sqlCode, g_strCONN, 3, 1
	If Not rsCode.EOF Then GetCMAdr = rsCode("CMCaddr") & ", " & rsCode("CMCcity") & ", " & rsCode("CMCstate") & ", " & rsCode("CMCzip")
	rsCode.Close
	Set rsCode = Nothing
End Function

Function GetCM(xxx)
	GetCM = "N/A" & """,""" & "N/A"
	Set rsName = Server.CreateObject("ADODB.RecordSet")
	sqlName = "SELECT * FROM Proj_Man_T WHERE [ID] = " & xxx
	rsName.Open sqlName, g_strCONN, 3, 1
	If Not rsName.EOF Then
		GetCM = rsName("lname") & """,""" & rsName("fname")
	End If
	rsName.Close
	Set rsName = Nothing
End Function

Function GetUser(xxx)
	GetUser = xxx
	If Z_Czero(xxx) = 0 Then Exit Function
	Set rsRate = Server.CreateObject("ADODB.RecordSet")
	sqlRate = "SELECT * FROM Input_t WHERE [index] = " & xxx 
	rsRate.Open sqlRate, g_strCONN, 1, 3
	If Not rsRate.EOF Then
		GetUser = rsRate("lname") & ", " & rsRate("fname")
	End If
	rsRate.Close
	Set rsRate = Nothing
End Function

Function GetNameWork(zzz)
	Set rsName = Server.CreateObject("ADODB.RecordSet")
	sqlName = "SELECT * FROM Worker_T WHERE Social_Security_Number = '" & zzz & "' "
	rsName.Open sqlName, g_strCONN, 3, 1
	If Not rsName.EOF Then
		GetNameWork = rsName("Lname") & """,""" & rsName("fname")
	End If
	rsName.Close
	Set rsName = Nothing
End Function

Function GetHoliday(tmpRdate, tmpRhrs, tmpDay)
	If tmpRhrs = 0 Then
	  GetHoliday = 0
	  Exit Function
	End If
	'GET DATE FOR COMPARISON
	Select Case tmpDay 
		Case "SUN" 
			tmpComDate = tmpRdate
		Case "MON" 
			tmpComDate = DateAdd("d", 1, tmpRdate)
		Case "TUE" 
			tmpComDate = DateAdd("d", 2, tmpRdate)
		Case "WED" 
			tmpComDate = DateAdd("d", 3, tmpRdate)
		Case "THU" 
			tmpComDate = DateAdd("d", 4, tmpRdate)
		Case "FRI" 
			tmpComDate = DateAdd("d", 5, tmpRdate)
		Case "SAT" 
			tmpComDate = DateAdd("d", 6, tmpRdate)
	End Select
	
	'CHECK REG HOLIDAY
	Set rsReg = Server.CreateOBject("ADODB.RecordSet")
	sqlReg = "SELECT * FROM RegHoliday_T WHERE month = " & Cint(Month(tmpComDate)) & " AND Day = " & Cint(Day(tmpComDate))
	rsReg.Open sqlReg, g_StrCONN, 1 ,3
	If rsReg.EOF Then
		GetHoliday = 0
	Else
		GetHoliday = tmpRhrs
	End If
	rsReg.Close
	Set rsReg = Nothing
	If GetHoliday = 0 Then
		'CHECK SPEC HOLIDAY
		Set rsSpec = Server.CreateOBject("ADODB.RecordSet")
		sqlSpec = "SELECT * FROM SpecHoliday_T WHERE month = " & Month(tmpComDate) & " AND Day = " & Day(tmpComDate) & " AND Year = " & Year(tmpComDate)
		rsSpec.Open sqlSpec, g_StrCONN, 1 ,3
		If rsSpec.EOF Then
			GetHoliday = 0
		Else
			GetHoliday = tmpRhrs
		End If
		rsSpec.Close
		Set rsSpec = Nothing
	End If
	
	'If Cbool(Instr(1,HolidayDate, tmpHoliday)> 0)Then
	'	GetHoliday = tmpRhrs
	'Else
	'	GetHoliday = 0
	'End If
End Function
	
Function GetPRate2(xxx)
	GetPRate2 = 0
	Set rsRate = Server.CreateObject("ADODB.RecordSet")
	sqlRate = "SELECT * FROM Consumer_T WHERE medicaid_number = '" & xxx & "'"
	rsRate.Open sqlRate, g_strCONN, 1, 3
	If Not rsRate.EOF Then
		GetPRate2 = rsRate("rate")
	End If
	rsRate.Close
	Set rsRate = Nothing
End Function

Function GetName(zzz)
	GetName = "N/A"
	Set rsName = Server.CreateObject("ADODB.RecordSet")
	sqlName = "SELECT * FROM worker_t WHERE Social_Security_Number = '" & zzz & "' "
	rsName.Open sqlName, g_strCONN, 3, 1
	If Not rsName.EOF Then
		GetName = rsName("Lname") & ", " & rsName("fname")
	End If
	rsName.Close
	Set rsName = Nothing
End Function

Function GetName2(zzz)
	GetName2 = "N/A"
	Set rsName = Server.CreateObject("ADODB.RecordSet")
	sqlName = "SELECT * FROM consumer_T WHERE medicaid_number = '" & zzz & "' "
	rsName.Open sqlName, g_strCONN, 3, 1
	If Not rsName.EOF Then
		GetName2 = rsName("Lname") & ", " & rsName("fname")
	End If
	rsName.Close
	Set rsName = Nothing
End Function

Function GetCode(xxx)
	If xxx = "" Then Exit Function
	Set rsCode = Server.CreateObject("ADODB.RecordSet")
	sqlCode = "SELECT [code] From Consumer_T WHERE medicaid_number = '" & xxx & "'"
	rsCode.Open sqlCode, g_strCONN, 3, 1
	If Not rsCode.EOF Then
		GetCode = rsCode("code")
	Else
		GetCode = "M"
	End If
	rsCode.Close
	Set rsCode = Nothing
End Function

Function ValidDate(dt1, dt2, xdate, xhrs, xday)
    If xhrs = 0 Then
        ValidDate = 0
        Exit Function
    End If
    Select Case xday
        Case "SUN"
            tmpComDate = xdate
        Case "MON"
            tmpComDate = DateAdd("d", 1, xdate)
        Case "TUE"
            tmpComDate = DateAdd("d", 2, xdate)
        Case "WED"
            tmpComDate = DateAdd("d", 3, xdate)
        Case "THU"
            tmpComDate = DateAdd("d", 4, xdate)
        Case "FRI"
            tmpComDate = DateAdd("d", 5, xdate)
        Case "SAT"
            tmpComDate = DateAdd("d", 6, xdate)
    End Select
    ValidDate = 0
    If dt1 = "" And dt2 = "" Then
        ValidDate = xhrs
    ElseIf dt1 <> "" And dt2 = "" Then
        If tmpComDate >= Z_CDate(dt1) Then ValidDate = xhrs
    ElseIf dt1 = "" And dt2 <> "" Then
        If tmpComDate <= Z_CDate(dt2) Then ValidDate = xhrs
    ElseIf dt1 <> "" And dt2 <> "" Then
        If tmpComDate >= Z_CDate(dt1) And tmpComDate <= Z_CDate(dt2) Then ValidDate = xhrs
    End If
End Function

Function SearchArrays60(strID, tmpID3)
	On Error Resume Next
	SearchArrays60 = -1
	lngMax = UBound(tmpID3)
	If Err.Number <> 0 Then Exit Function
	
	For lngI = 0 to lngMax
		If tmpID3(lngI) = strID Then Exit For
	Next
	If lngI > lngMax Then Exit Function
	SearchArrays60 = lngI
End Function

Function GetPM1(zzz)
	GetPM1 = 0
	Set rsPM1 = Server.CreateObject("ADODB.RecordSet")
	sqlPM1= "SELECT * FROM worker_T WHERE Social_Security_Number = '" & zzz & "' "
	rsPM1.Open sqlPM1, g_strCONN, 3, 1
	If Not rsPM1.EOF Then
		GetPM1 = Z_Czero(rsPM1("PM1"))
	End If
	rsPM1.Close
	Set rsPM1 = Nothing
End Function

Function GetPM2(zzz)
	GetPM2 = 0
	Set rsPM2 = Server.CreateObject("ADODB.RecordSet")
	sqlPM2= "SELECT * FROM worker_T WHERE Social_Security_Number = '" & zzz & "' "
	rsPM2.Open sqlPM2, g_strCONN, 3, 1
	If Not rsPM2.EOF Then
		GetPM2 = Z_Czero(rsPM2("PM2"))
	End If
	rsPM2.Close
	Set rsPM2 = Nothing
End Function

Function GetName3(zzz)
	GetName3 = "N/A"
	If zzz <> "" Then
		Set rsName = Server.CreateObject("ADODB.RecordSet")
		sqlName = "SELECT * FROM Proj_Man_T WHERE ID = " & zzz 
		rsName.Open sqlName, g_strCONN, 3, 1
		If Not rsName.EOF Then
			GetName3 = rsName("Lname") & ", " & rsName("fname")
		End If
		rsName.Close
		Set rsName = Nothing
	End If
End Function

Function SearchArrays3(xxx)
	SearchArrays3 = -1
	On Error Resume Next	
	lngMax = UBound(tmpEmp)
	If Err.Number <> 0 Then Exit Function
	
	For lngI = 0 to lngMax
		If tmpemp(lngI) = xxx  Then Exit For
	Next
	If lngI > lngMax Then Exit Function
	SearchArrays3 = lngI
End Function

Function SearchArrays32(xxx)
	SearchArrays32 = -1
	On Error Resume Next	
	lngMax = UBound(tmpEmp2)
	If Err.Number <> 0 Then Exit Function
	
	For lngI = 0 to lngMax
		If tmpemp2(lngI) = xxx  Then Exit For
	Next
	If lngI > lngMax Then Exit Function
	SearchArrays32 = lngI
End Function

Function SearchArrays4(xxx)
	SearchArrays4 = -1
	On Error Resume Next	
	lngMax = UBound(tmpempADP)
	If Err.Number <> 0 Then Exit Function
	
	For lngI = 0 to lngMax
		If tmpempADP(lngI) = xxx  Then Exit For
	Next
	If lngI > lngMax Then Exit Function
	SearchArrays4 = lngI
End Function

Function GetFileNum(xxx)
	Set rsName = Server.CreateObject("ADODB.RecordSet")
	sqlName = "SELECT * FROM worker_t WHERE Social_Security_Number = '" & xxx & "' "
	rsName.Open sqlName, g_strCONN, 3, 1
	If Not rsName.EOF Then
		GetFileNum = rsName("FileNum")
	Else
		GetFileNum = "N/A"
	End If
	rsName.Close
	Set rsName = Nothing
End Function

Function GetAPM2(xxx)
	GetAPM2 = "N/A"
	sqlRel = "SELECT * FROM Consumer_T WHERE medicaid_number = '" & xxx & "'"
	Set rsRel = Server.CreateObject("ADODB.RecordSet")
	rsRel.Open sqlRel, g_strCONN, 3, 1
	If Not rsRel.EOF Then
		GetAPM2 = GetName3(rsRel("PMID"))
	End If
	rsRel.Close
	Set rsRel = Nothing
End Function

Function GetName(zzz)
	GetName = "N/A"
	Set rsName = Server.CreateObject("ADODB.RecordSet")
	sqlName = "SELECT * FROM worker_t WHERE Social_Security_Number = '" & zzz & "' "
	rsName.Open sqlName, g_strCONN, 3, 1
	If Not rsName.EOF Then
		GetName = rsName("Lname") & ", " & rsName("fname")
	End If
	rsName.Close
	Set rsName = Nothing
End Function

Function GetName3(zzz)
	GetName3 = "N/A"
	Set rsName = Server.CreateObject("ADODB.RecordSet")
	sqlName = "SELECT * FROM Proj_Man_T WHERE ID = " & zzz 
	rsName.Open sqlName, g_strCONN, 3, 1
	If Not rsName.EOF Then
		GetName3 = rsName("Lname") & ", " & rsName("fname")
	End If
	rsName.Close
	Set rsName = Nothing
End Function

Function ACdesc(xxx)
	If xxx = "" Then exit function
	
	Set rsAC = Server.CreateObject("ADODB.RecordSet")
	sqlAC = "SELECT [desc] FROM activity_T WHERE code = " & xxx'strAC(ctr)
	rsAC.Open sqlAC, g_strCONN, 3, 1
	If Not rsAC.EOF Then
		ACdesc = rsAC("desc") 
		'response.write "Xxx: " & xxx & " = " & rsAC("desc")  & "<br>"
	End If
	rsAC.Close
	Set rsAC = Nothing
End function
%>