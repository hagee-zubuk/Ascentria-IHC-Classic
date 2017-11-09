<%Language=VBScript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_security.asp" -->
<%
	DIM tmp(), xtmp()
	Dim tmpDates(), tmpWorID(), tmpIDs(),  tmpHrs(), tmpPMID(), tmpEmpID(), tmpPMID2()
	dim tmpemp(), tmpcon(), tmpmile(), tmpmilecap(), tmpHrsExt(), tmpID()
	
		
Server.ScriptTimeout = 360000

Function ApproveNum(strPhone, strCID)
	ApproveNum = False
	Set rsWID = Server.CreateObject("ADODB.RecordSet")
	sqlWID = "SELECT Aphone1, Aphone2, Aphone3, Aphone4, Aphone5 FROM Consumer_T WHERE Medicaid_Number = '" & strCID & "'"
	rsWID.Open sqlWID, g_strCONN, 3, 1
	If Not rsWID.EOF Then
		ApproveFon = "," & rsWID("Aphone1") & "," & rsWID("Aphone2") & "," & rsWID("Aphone3") & "," & rsWID("Aphone4") & "," & rsWID("Aphone5") & "," 
		myPhone = "," & strPhone & ","
		If Instr(ApproveFon, myPhone) > 0 Then ApproveNum = True
	End If
	rsWID.Close
	Set rsWID = Nothing
End Function
	Function GetUser(xxx)
	GetUser = xxx
	If Z_Czero(xxx) = 0 Then Exit Function
	Set rsRate = Server.CreateObject("ADODB.RecordSet")
	sqlRate = "SELECT * FROM Input_t WHERE index = " & xxx 
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
		GetNameWork = rsName("Lname") & ", " & rsName("fname")
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
	Function GetPRate(hrs, PID, mytype)
	GetPRate = 0
	Set rsRate = Server.CreateObject("ADODB.RecordSet")
	sqlRate = "SELECT * FROM Consumer_T WHERE medicaid_number = '" & PID & "'"
	rsRate.Open sqlRate, g_strCONN, 1, 3
	If Not rsRate.EOF Then
		If mytype = 0 Then
			GetPRate = rsRate("rate") * hrs
		Else
			GetPRate = rsRate("rate") * hrs * 1.5
		End If
	End If
	rsRate.Close
	Set rsRate = Nothing
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
	Function GetMRate(tmpCDate, tmpTMile)
	If Z_CZero(tmpTMile) = 0 Then
	  GetMRate = 0
	  Exit Function
	End If
	Set rsRateX = Server.CreateObject("ADODB.RecordSet")
	sqlRateX = "SELECT * FROM MileRate_T WHERE miledate <= #" & tmpCDate & "# ORDER BY miledate DESC"
	rsRateX.Open sqlRateX, g_strCONN, 3, 1
	If Not rsRateX.EOF Then
		tmpCurrMRate = rsRateX("milerate")
		tmpCurrMDate = rsRateX("miledate")
	Else
		tmpCurrMRate = 0
	End If
	rsRateX.Close
	Set rsRateX = Nothing
	'COMPUTE FOR Mileage Rate
	GetMRate = tmpTMile * tmpCurrMRate
End Function
	Function GetNextSat(xxx,yyy)
	difwk = DateDiff("ww", yyy, PDate)
		myDATE = xxx
    If difwk >= 0 Then
        wknum = difwk + 1
        If Z_IsOdd2(wknum) Then
            If WeekdayName(Weekday(myDATE), True) = "Sun" Then
                sunDATE = myDATE
                satDATE = DateAdd("d", 13, sunDATE)
            Else
                difDATE = DatePart("w", myDATE)
                sunDATE = DateAdd("d", -CInt(difDATE - 1), myDATE)
                satDATE = DateAdd("d", 13, sunDATE)
            End If
        Else
            If WeekdayName(Weekday(myDATE), True) = "Sun" Then
                sunDATE = DateAdd("d", -7, myDATE)
                satDATE = DateAdd("d", 13, sunDATE)
            Else
                difDATE = DatePart("w", myDATE)
                sunDATE = DateAdd("d", -CInt(difDATE + 6), myDATE)
                satDATE = DateAdd("d", 13, sunDATE)
            End If
        End If
    Else
        wknum = difwk
        If Not Z_IsOdd2(wknum) Then
            If WeekdayName(Weekday(myDATE), True) = "Sun" Then
                sunDATE = myDATE
                satDATE = DateAdd("d", 13, sunDATE)
            Else
                difDATE = DatePart("w", myDATE)
                sunDATE = DateAdd("d", -CInt(difDATE - 1), myDATE)
                satDATE = DateAdd("d", 13, sunDATE)
            End If
        Else
            If WeekdayName(Weekday(myDATE), True) = "Sun" Then
                sunDATE = DateAdd("d", -7, myDATE)
                satDATE = DateAdd("d", 13, sunDATE)
            Else
                difDATE = DatePart("w", myDATE)
                sunDATE = DateAdd("d", -CInt(difDATE + 6), myDATE)
                satDATE = DateAdd("d", 13, sunDATE)
            End If
        End If
    End If
    GetNextSat = satDATE
End Function
	Function GetNameCSV(zzz)
	GetNameCSV = "N/A"
	Set rsName = Server.CreateObject("ADODB.RecordSet")
	sqlName = "SELECT * FROM Consumer_t WHERE Medicaid_number = '" & zzz & "' "
	rsName.Open sqlName, g_strCONN, 3, 1
	If Not rsName.EOF Then
		GetNameCSV = rsName("Lname") & """,""" & rsName("fname")
	End If
	rsName.Close
	Set rsName = Nothing
End Function
Function GetNameWorkCSV(zzz)
	Set rsName = Server.CreateObject("ADODB.RecordSet")
	sqlName = "SELECT * FROM Worker_T WHERE Social_Security_Number = '" & zzz & "' "
	rsName.Open sqlName, g_strCONN, 3, 1
	If Not rsName.EOF Then
		GetNameWorkCSV = rsName("Lname") & """,""" & rsName("fname")
	End If
	rsName.Close
	Set rsName = Nothing
End Function

	Function GetRate(tmpRdate, tmpRhrs, tmpDay)
	If tmpRhrs = 0 Then
	  GetRate = 0
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
	'GET RATE FOR DATE
	Set rsRateX = Server.CreateObject("ADODB.RecordSet")
	sqlRateX = "SELECT * FROM Rate_T WHERE rDate <= #" & tmpComDate & "# ORDER BY rDate DESC"
	rsRateX.Open sqlRateX, g_strCONN, 3, 1
	If Not rsRateX.EOF Then
		tmpCurrRate = rsRateX("rate")
		tmpCurrDate = rsRateX("rDate")
	End If
	rsRateX.Close
	Set rsRateX = Nothing
	'COMPUTE FOR UNIT
	tmpUnit = tmpRhrs * 4
	'COMPUTE FOR RATE
	GetRate = tmpUnit * tmpCurrRate
End Function
	Function GetCMAgency(xxx)
		GetCMAgency = "N/A"
		Set rsName = Server.CreateObject("ADODB.RecordSet")
		sqlName = "SELECT * FROM Case_Manager_T WHERE [index] = " & xxx
		rsName.Open sqlName, g_strCONN, 3, 1
		If Not rsName.EOF Then
			GetCMAgency = rsName("Agency")
		End If
		rsName.Close
		Set rsName = Nothing
	End Function
	Function GetWorkName(xxx)
		GetWorkName = "N/A"
		Set rsName = Server.CreateObject("ADODB.RecordSet")
		sqlName = "SELECT * FROM Worker_T WHERE [index] = " & xxx
		rsName.Open sqlName, g_strCONN, 3, 1
		If Not rsName.EOF Then
			GetWorkName = rsName("Lname") & ", " & rsName("fname")
		End If
		rsName.Close
		Set rsName = Nothing
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
	Function GetDate(xdate, xday)
		Select Case xday
        Case "SUN"
            GetDate = xdate
        Case "MON"
            GetDate = DateAdd("d", 1, xdate)
        Case "TUE"
            GetDate = DateAdd("d", 2, xdate)
        Case "WED"
            GetDate = DateAdd("d", 3, xdate)
        Case "THU"
            GetDate = DateAdd("d", 4, xdate)
        Case "FRI"
            GetDate = DateAdd("d", 5, xdate)
        Case "SAT"
            GetDate = DateAdd("d", 6, xdate)
    End Select
	End Function
	Function GetAPM(xxx)
		GetAPM = "N/A"
		sqlRel = "SELECT * FROM ConWork_T, consumer_T WHERE WID = '" & xxx & "' AND CID = Medicaid_Number"
		Set rsRel = Server.CreateObject("ADODB.RecordSet")
		rsRel.Open sqlRel, g_strCONN, 3, 1
		If Not rsRel.EOF Then
			GetAPM = GetName3(rsRel("PMID"))
		End If
		rsRel.Close
		Set rsRel = Nothing
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
	Function GetWarn(zzz)
		If zzz = 0 Then GetWarn = "Verbal Warning"
		If zzz = 1 Then GetWarn = "Written Warning"
		If zzz = 2 Then GetWarn = "Final Warning"
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
	Function GetAllwdHrs(zzz)
		GetAllwdHrs = "N/A"
		Set rsName = Server.CreateObject("ADODB.RecordSet")
		sqlName = "SELECT * FROM consumer_T WHERE medicaid_number = '" & zzz & "'" 
		rsName.Open sqlName, g_strCONN, 3, 1
		If Not rsName.EOF Then
			GetAllwdHrs = rsName("maxhrs") 
		End If
		rsName.Close
		Set rsName = Nothing
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
	Function ValidDate(dt1, dt2, xdate, xhrs, xday)
    If xhrs = 0 Then
        ValidDate = 0
        Exit Function
    End If
    Select Case Ucase(xday)
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
	Function Z_Find2WkPeriod(dtDate)
	' Returns a string with the start & end dates 
	' for the 2-week period dtDate belongs to.
	'
	' Format of the returned string is: mm/dd-mm/dd
	' dtDate must be a valid date
	
	DIM	myDt, rawWeek, wkDay, dtStartSun, dtEndSat, lngWk, strTmp, dtRef
	
	' START: validity check
	Z_Find2WkPeriod = ""
	If Not IsDate(dtDate) Then Exit Function
	
	myDt = Z_CDate(dtDate)
	rawWeek = DatePart("ww", myDt, 1, 1)	' 1- first day of week is Sunday, 1- week count starts when January 1 occurs
	wkDay = Weekday(myDt, 1) - 1			' 1- first day of week is Sunday
	' find the preceeding Sunday, if not a Sunday itself...
	If wkDay > 0 Then
		dtStartSun = DateAdd("d", wkDay * (-1), myDt)
	Else
		dtStartSun = myDt
	End If
	
	dtRef = Z_CDate(wk1)
	lngWk = DateDiff("ww", dtStartSun, dtRef)
	
	If Z_IsOdd2(lngWk) Then dtStartSun = DateAdd("d", -7, dtStartSun)
	dtEndSat = DateAdd("d", 13, dtStartSun)
	
	' dtStartSun & dtEndSat now contains start and end dates for the period, respectively, now make it a string
	strTmp = DatePart("m", dtStartSun) & "/" & DatePart("d", dtStartSun) & "/" & DatePart("yyyy", dtStartSun)
	strTmp = strTmp & " - " & _
			DatePart("m", dtEndSat) & "/" & DatePart("d", dtEndSat) & "/" & DatePart("yyyy", dtEndSat)
	Z_Find2WkPeriod = strTmp
End Function
Function SearchArrays60(strID, tmpID)
	On Error Resume Next
	SearchArrays60 = -1
	lngMax = UBound(tmpID)
	If Err.Number <> 0 Then Exit Function
	
	For lngI = 0 to lngMax
		If tmpID(lngI) = strID Then Exit For
	Next
	If lngI > lngMax Then Exit Function
	SearchArrays60 = lngI
End Function
Function SearchArrays4(strwk, strEID, strWID, tmpDates, tmpEmpID, tmpWorID)
		DIM	lngMax, lngI
	
	' START: validity check
	SearchArrays4 = -1
	On Error Resume Next	
	lngMax = UBound(tmpDates)
	If Err.Number <> 0 Then Exit Function
	
	For lngI = 0 to lngMax
		If tmpDates(lngI) = strWk And tmpWorID(lngI) = strWID And tmpEmpID(lngI) = strEID Then Exit For
	Next
	If lngI > lngMax Then Exit Function
	SearchArrays4 = lngI
End Function
Function SearchArrays(strwk, strEID, strWID, tmpDates, tmpEmpID, tmpWorID)
		DIM	lngMax, lngI
	
	' START: validity check
	SearchArrays = -1
	On Error Resume Next	
	lngMax = UBound(tmpDates)
	If Err.Number <> 0 Then Exit Function
	
	For lngI = 0 to lngMax
		If tmpDates(lngI) = strWk And tmpWorID(lngI) = strWID And tmpEmpID(lngI) = strEID Then Exit For
	Next
	If lngI > lngMax Then Exit Function
	SearchArrays = lngI
End Function
Function SearchArrays2(strwk, strEID, strWID, tmpDates, tmpEmpID, tmpWorID)
		DIM	lngMax, lngI
	
	' START: validity check
	SearchArrays2 = -1
	On Error Resume Next	
	lngMax = UBound(tmpDates)
	If Err.Number <> 0 Then Exit Function
	
	For lngI = 0 to lngMax
		If tmpDates(lngI) = strWk And tmpWorID(lngI) = strWID Then Exit For
	Next
	If lngI > lngMax Then Exit Function
	SearchArrays2 = lngI
End Function
Function GetPM(xxx)
	GetPM = "N/A"
	Set rsPM = Server.CreateObject("ADODB.RecordSet")
	sqlPM = "SELECT * FROM Proj_Man_T WHERE ID = " & xxx
	rsPM.Open sqlPM, g_strCONN, 3, 1
	If Not rsPM.EOF Then
		GetPM = rsPM("Lname") & ", " & rsPM("Fname")
	End If
	rsPM.Close
	Set rsPM = Nothing
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
	Session("PrintPrev") = ""
	Session("PrintPrevRep") = ""
	Session("PrintPrevPRoc") = ""
		If Request("sunDATE") <> "" and Request("satDATE") <> "" and Request("chkr") = 1 then 
			sunDATE = Request("sunDATE")
			satDATE = Request("satDATE")
			If Request("SelRep") = 27  Then 
				if Request("seltype") = 1 then selPay = "SELECTED"
				if Request("seltype") = 2 then selMed = "SELECTED"
				if Request("seltype") = 3 then selOthers = "SELECTED"
				sel27 = "SELECTED"
			ElseIf Request("SelRep") = 39 Then
				sel39 = "SELECTED"
			ElseIf Request("SelRep") = 44 Then
				sel44 = "SELECTED"
			ElseIf Request("SelRep") = 55 Then
				sel55 = "SELECTED"
				op1 = ""
				op2 = ""
				op3 = ""
				myOpt = Request("selopt")
				If myOpt = 0 Then op1 = "Selected"
				If myOpt = 1 Then op2 = "Selected"
				If myOpt = 2 Then op3 = "Selected"
				uri1 = ""
				uri2 = ""
				myUri = Request("seluri")
				If myUri = 0 Then uri1 = "Selected"
				If myUri = 1 Then uri2 = "Selected"
			ElseIf Request("SelRep") = 51 Then
				sel51 = "SELECTED"
			ElseIf Request("SelRep") = 64 Then
				sel64 = "SELECTED"
			ElseIf Request("SelRep") = 66 Then
				sel66 = "SELECTED"
				if Request("seltype42") = 2 then selMed = "SELECTED"
				if Request("seltype42") = 3 then selOthers = "SELECTED"
			ElseIf Request("SelRep") = 69 Then
				sel69 = "SELECTED"
			End If
		End If
		If Request("err") = 31 Then 
			Sel31= "Selected"
			Session("MSG") = "Invalid 'From:' and/or 'To:' date."
		End If	
		If Request("err") = 32 Then 
			Sel32= "Selected"
			Session("MSG") = "Invalid 'From:' and/or 'To:' date."
		End If	
		If Request("err") = 24 Then 
			Sel24= "Selected"
			Session("MSG") = "Invalid 'From:' and/or 'To:' date."
		End If	
		If Request("err") = 41 Then 
			Sel41= "Selected"
			Session("MSG") = "Invalid 'From:' and/or 'To:' date."
		End If	
		If Request("err") = 47 Then 
			Sel47= "Selected"
			Session("MSG") = "Invalid 'From:' and/or 'To:' date."
		End If	
		If Request("err") = 55 Then 
			Sel55= "Selected"
			Session("MSG") = "Invalid 'From:' and/or 'To:' date."
		End If	
		If Request("err") = 59 Then 
			Sel59= "Selected"
			Session("MSG") = "Invalid 'From:' and/or 'To:' date."
		End If	
		If Request("err") = 60 Then 
			Sel60= "Selected"
			Session("MSG") = "Invalid 'From:' and/or 'To:' date."
		End If	
		If Request("err") = 66 Then 
			Sel66= "Selected"
			Session("MSG") = "Invalid 'To:' date."
		End If
		'GET CONSUMERS
		Set rsCon = Server.CreateObject("ADODB.RecordSet")
		sqlCon = "SELECT * FROM Consumer_T, C_Status_T WHERE consumer_T.Medicaid_Number = C_Status_T.Medicaid_Number AND Active = true ORDER BY Lname, Fname"
		rsCon.Open sqlCon, g_strCONN, 3, 1
		Do Until rsCon.EOF
			Consel =""
			if Request("selcon") <> "0" Then 
				If rsCON("Consumer_t.Medicaid_Number") = Request("selcon") Then Consel = "SELECTED"
			End If
			strCON = strCON & "<option value='" & rsCON("Consumer_t.Medicaid_Number") & "' " & Consel & " >" & GetName2(rsCON("Consumer_t.Medicaid_Number")) & "</option>" & vbCrLf
			rsCon.MoveNExt
		Loop
		rsCon.Close
		Set rsCon = Nothing
		'GET WORKERS
		Set rsWor = Server.CreateObject("ADODB.RecordSet")
		sqlWor = "SELECT * FROM Worker_T  WHERE status = 'Active' ORDER BY Lname, Fname"
		rsWor.Open sqlWor, g_strCONN, 3, 1
		Do Until rsWor.EOF
			strWOR = strWOR & "<option value='" & rsWor("Social_Security_Number") & "'>" & GetName(rsWor("Social_Security_Number")) & "</option>" & vbCrLf
			rsWor.MoveNExt
		Loop
		rsWor.Close
		Set rsWor = Nothing
		
		If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
			server.scripttimeout = 360000
			Sel1 = ""
			Sel2 = ""
			Sel3 = ""
			Sel4 = ""
			Sel5 = ""
			Sel6 = ""
			Sel7 = ""
			Sel8 = ""
			Sel9 = ""
			Sel10 = ""
			Sel11 = ""
			Sel12 = ""
			Sel13 = ""
			Sel14 = ""
			Sel15 = ""
			Sel16 = ""
			Sel17 = ""
			Sel18 = ""
			Sel19 = ""
			Sel20 = ""
			Sel21 = ""
			Sel22 = ""
			Sel23 = ""
			Sel24 = ""
			Sel25 = ""
			Sel26 = ""
			Sel27 = ""
			Sel28 = ""
			Sel29 = ""
			Sel30 = ""
			Sel31 = ""
			Sel32 = ""
			Sel33 = ""
			Sel34 = ""
			Sel35 = ""
			Sel36 = ""
			Sel37 = ""
			Sel38 = ""
			Sel39 = ""
			Sel40 = ""
			Sel41 = ""
			Sel42 = ""
			Sel43 = ""
			Sel44 = ""
			Sel45 = ""
			Sel46 = ""
			Sel47 = ""
			Sel48 = ""
			Sel49 = ""
			Sel50 = ""
			Sel51 = ""
			Sel52 = ""
			Sel53 = ""
			Sel54 = ""
			Sel55 = ""
			Sel56 = ""
			Sel57 = ""
			Sel58 = ""
			Sel59 = ""
			Sel60 = ""
			Sel61 = ""
			Sel62 = ""
			Sel63 = ""
			Sel64 = ""
			Sel65 = ""
			Sel66 = ""
			Sel67 = ""	
			Sel68 = ""
			Sel69 = ""
								
			If Request("SelRep") = 1 Then Sel1 = "Selected"
			If Request("SelRep") = 2 Then Sel2 = "Selected"
			If Request("SelRep") = 3 Then Sel3 = "Selected"
			If Request("SelRep") = 4 Then Sel4 = "Selected"
			If Request("SelRep") = 5 Then Sel5 = "Selected"
			If Request("SelRep") = 6 Then Sel6 = "Selected"
			If Request("SelRep") = 7 Then Sel7 = "Selected"
			If Request("SelRep") = 8 Then Sel8 = "Selected"
			If Request("SelRep") = 9 Then Sel9 = "Selected"
			If Request("SelRep") = 10 Then Sel10 = "Selected"
			If Request("SelRep") = 11 Then Sel11 = "Selected"
			If Request("SelRep") = 12 Then Sel12 = "Selected"
			If Request("SelRep") = 13 Then Sel13 = "Selected"
			If Request("SelRep") = 14 Then Sel14 = "Selected"
			If Request("SelRep") = 15 Then Sel15 = "Selected"
			If Request("SelRep") = 16 Then Sel16 = "Selected"
			If Request("SelRep") = 17 Then Sel17 = "Selected"
			If Request("SelRep") = 18 Then Sel18 = "Selected"
			If Request("SelRep") = 19 Then Sel19 = "Selected"
			If Request("SelRep") = 20 Then Sel20 = "Selected"
			If Request("SelRep") = 21 Then Sel21 = "Selected"
			If Request("SelRep") = 22 Then Sel22 = "Selected"
			If Request("SelRep") = 23 Then Sel23 = "Selected"
			If Request("SelRep") = 24 Then Sel24 = "Selected"
			If Request("SelRep") = 25 Then Sel25 = "Selected"
			If Request("SelRep") = 26 Then Sel26= "Selected"
			If Request("SelRep") = 27 Then Sel27= "Selected"
			If Request("SelRep") = 28 Then Sel28= "Selected"
			If Request("SelRep") = 29 Then Sel29= "Selected"
			If Request("SelRep") = 30 Then Sel30= "Selected"
			If Request("SelRep") = 31 Then Sel31= "Selected"
			If Request("SelRep") = 32 Then Sel32= "Selected"
			If Request("SelRep") = 33 Then Sel33= "Selected"
			If Request("SelRep") = 34 Then Sel34= "Selected"
			If Request("SelRep") = 35 Then Sel35= "Selected"
			If Request("SelRep") = 36 Then Sel36= "Selected"
			If Request("SelRep") = 37 Then Sel37= "Selected"
			If Request("SelRep") = 38 Then Sel38= "Selected"
			If Request("SelRep") = 39 Then Sel39= "Selected"
			If Request("SelRep") = 40 Then Sel40= "Selected"
			If Request("SelRep") = 41 Then Sel41= "Selected"
			If Request("SelRep") = 42 Then Sel42= "Selected"
			If Request("SelRep") = 43 Then Sel43= "Selected"	
			If Request("SelRep") = 44 Then Sel44= "Selected"		
			If Request("SelRep") = 45 Then Sel45= "Selected"	
			If Request("SelRep") = 46 Then Sel46= "Selected"	
			If Request("SelRep") = 47 Then Sel47= "Selected"
			If Request("SelRep") = 48 Then Sel48= "Selected"
			If Request("SelRep") = 49 Then Sel49= "Selected"
			If Request("SelRep") = 50 Then Sel50= "Selected"
			If Request("SelRep") = 51 Then Sel51= "Selected"
			If Request("SelRep") = 52 Then Sel52= "Selected"
			If Request("SelRep") = 53 Then Sel53= "Selected"	
			If Request("SelRep") = 54 Then Sel54= "Selected"
			If Request("SelRep") = 55 Then Sel55= "Selected"
			If Request("SelRep") = 56 Then Sel56= "Selected"
			If Request("SelRep") = 57 Then Sel57= "Selected"
			If Request("SelRep") = 58 Then Sel58= "Selected"
			If Request("SelRep") = 59 Then Sel59= "Selected"
			If Request("SelRep") = 60 Then Sel60= "Selected"
			If Request("SelRep") = 61 Then Sel61= "Selected"
			If Request("SelRep") = 62 Then Sel62= "Selected"
			If Request("SelRep") = 63 Then Sel63= "Selected"
			If Request("SelRep") = 64 Then Sel64= "Selected"
			If Request("SelRep") = 65 Then Sel65= "Selected"
			If Request("SelRep") = 66 Then Sel66= "Selected"
			If Request("SelRep") = 67 Then Sel67= "Selected"
			If Request("SelRep") = 68 Then Sel68= "Selected"
			If Request("SelRep") = 69 Then Sel69= "Selected"
								
			If Request("SelRep") = 13 Then '1
				Session("MSG") = "Site Visit for Consumers report."
				typ = Request("SelRep")
				Srt = 1
				If Request("Sort") = 1 then Srt = 0
				strHEAD = "<tr bgcolor='#040C8B'><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Site Visit Date" & _
					"</b></font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Name</b></font></td><td align='center'>" & _
					"<font size='1' face='trebuchet ms' color='white'><b>Address</b></font></td><td align='center'><font size='1' " & _
					"face='trebuchet ms' color='white'><b>Phone No.</b></font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Project " & _
					" Manager</b></font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Comment</b></font>" & _
					"</td></tr>"
				
				Set rsTBL = Server.CreateObject("ADODB.RecordSet")
		
				sqlTBL = "SELECT * FROM Consumer_t, C_Status_t, C_Site_Visit_Dates_t , Proj_Man_T WHERE " & _
					"PMID = Proj_Man_T.ID AND Consumer_t.Medicaid_number = C_Status_t.Medicaid_number AND Consumer_t.Medicaid_number =" & _
					" C_Site_Visit_Dates_t.Medicaid_number AND Active = true AND NOT IsNull(Site_V_Date) " & _
					"ORDER BY Proj_Man_T.Lname, Proj_Man_T.Fname ASC, Consumer_t.Medicaid_number, Site_V_Date DESC"
				
				rsTBL.Open sqlTBL, g_strCONN, 1, 3
				
				tmpIDx = ""
				x = 0
				
				Do Until rsTBL.EOF
					If rsTBL("Consumer_t.Medicaid_number") <> tmpIDx then 
						ReDim Preserve tmp(x)
						tmp(x) = rsTBL("C_Site_Visit_Dates_t.index") & "|" & rsTBL("Site_V_Date") & "|" & rsTBL("PMID")
						x = x + 1
					End If
					tmpIDx = rsTBL("Consumer_t.Medicaid_number")
					rsTBL.MoveNext
				Loop
				If Request("Sort") <> 1 Then 
					For i = x - 2 to 0 Step - 1
						For j = 0 To i
							tmpj = split(tmp(j),"|")
							tmpj1 = split(tmp(j+1),"|")
							If Cdate(tmpj(1)) < Cdate(tmpj1(1)) AND tmpj(2) = tmpj1(2) Then
								intTemp = tmp(j + 1)
	              			tmp(j + 1) = tmp(j)
	              			tmp(j) = intTemp
							End If
						Next 
					Next 
				End If
				rsTBL.Close
				Set rsTBL = Nothing
				
				Set rsTBL2 = Server.CreateObject("ADODB.RecordSet")
				zzz = 0
				Do Until zzz = x 
					
					tmp2 = split(tmp(zzz),"|")	
					
					sqlTBL2 = "SELECT * FROM Consumer_t, C_Status_t, C_Site_Visit_Dates_t WHERE " & _
					"Consumer_t.Medicaid_number = C_Status_t.Medicaid_number AND Consumer_t.Medicaid_number =" & _
					" C_Site_Visit_Dates_t.Medicaid_number AND Active = true AND NOT IsNull(Site_V_Date) " & _
					"AND C_Site_Visit_Dates_t.index = " & tmp2(0)
					
					rsTBL2.Open sqlTBL2, g_strCONN, 1, 3
					
						Set rsPM = Server.CreateObject("ADODB.RecordSet")
						sqlPM = "SELECT * FROM Proj_Man_T WHERE ID = " & rsTBL2("PMID")
						rsPM.Open sqlPM, g_strCONN, 1, 3
						If Not rsPM.EOF Then
							PMname = rsPM("lname") & ", " & rsPM("fname")
						Else
							PMname = rsPM("ID")
						End If
						rsPM.Close
						Set rsPM = Nothing
						newComment = Replace(rsTBL2("Comments"), "|",  " ")
						strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL2("Site_V_Date") & _
							"</td><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL2("lname") & ", " & _
							rsTBL2("fname") & "</td><td align='center'><font size='1' face='trebuchet ms'>" & _
							rsTBL2("Address") & ", " & rsTBL2("City") & ", " & rsTBL2("State") & ", " & rsTBL2("zip") & "</font></td>" & _
							"<td align='center'><font size='1' face='trebuchet ms'>&nbsp;" & rsTBL2("PhoneNo") & _
							"<td align='center'><font size='1' face='trebuchet ms'>&nbsp;" & _
							PMname & "</font></td></font></td><td align='center' width='300px'><font size='1' face='trebuchet ms'>" & _
							newComment & "</font></td></tr>"
					rsTBL2.Close
					zzz =zzz + 1
				Loop
				Set rsTBL2 = Nothing
			ElseIf Request("SelRep") = 11 Then '2
				Session("MSG") = "Phone Contact for Consumers report."
				typ = Request("SelRep")
				Srt = 1
				If Request("Sort") = 1 then Srt = 0
				strHEAD = "<tr bgcolor='#040C8B'><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Phone Call Date</b>" & _
					"</font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Name</b></font></td><td align='center'>" & _
					"<font size='1' face='trebuchet ms' color='white'><b>Address</b></font></td><td align='center'><font size='1' face='trebuchet ms' color='white'>" & _
					"<b>Phone No.</b></font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Project " & _
					" Manager</b></font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Comment</b></font>" & _
					"</td></tr>"
				Set rsTBL = Server.CreateObject("ADODB.RecordSet")
		
				sqlTBL = "SELECT Consumer_t.*, C_Site_Visit_Dates_t.*, C_Status_t.Active FROM Consumer_t, C_Status_t, " & _
						"C_Site_Visit_Dates_t, Proj_Man_T WHERE PMID = Proj_Man_T.ID AND Consumer_t.Medicaid_number = C_Status_t.Medicaid_number AND " & _
						"Consumer_t.Medicaid_number = C_Site_Visit_Dates_t.Medicaid_number AND C_Status_t.Active = true " & _
						"AND Active = true AND NOT IsNull(phoneCon_last) ORDER BY Proj_Man_T.Lname, Proj_Man_T.Fname ASC, C_Site_Visit_Dates_t.Medicaid_number, phoneCon_last DESC"
				
				rsTBL.Open sqlTBL, g_strCONN, 1, 3
				
				tmpIDx = ""
				x = 0
				
				Do Until rsTBL.EOF
					If rsTBL("Consumer_t.Medicaid_number") <> tmpIDx then 
						ReDim Preserve tmp(x)
						tmp(x) = rsTBL("C_Site_Visit_Dates_t.index") & "|" & rsTBL("phoneCon_last") & "|" & rsTBL("PMID")
						x = x + 1
					End If
					tmpIDx = rsTBL("Consumer_t.Medicaid_number")
					rsTBL.MoveNext
				Loop
				
				If Request("Sort") <> 1 Then 
					For i = x - 2 to 0 Step - 1
						For j = 0 To i
							tmpj = split(tmp(j),"|")
							tmpj1 = split(tmp(j+1),"|")
							If Cdate(tmpj(1)) < Cdate(tmpj1(1)) AND tmpj(2) = tmpj1(2) Then
								intTemp = tmp(j + 1)
				              tmp(j + 1) = tmp(j)
				              tmp(j) = intTemp
							End If
						Next 
					Next 
				End If
				
				rsTBL.Close
				Set rsTBL = Nothing
				
				Set rsTBL2 = Server.CreateObject("ADODB.RecordSet")
				zzz = 0
				Do Until zzz = x 
					
					tmp2 = split(tmp(zzz),"|")	
					
					sqlTBL2 = "SELECT * FROM Consumer_t, C_Status_t, C_Site_Visit_Dates_t WHERE " & _
					"Consumer_t.Medicaid_number = C_Status_t.Medicaid_number AND Consumer_t.Medicaid_number =" & _
					" C_Site_Visit_Dates_t.Medicaid_number AND Active = true AND NOT IsNull(phoneCon_last) " & _
					"AND C_Site_Visit_Dates_t.index = " & tmp2(0)
					
					rsTBL2.Open sqlTBL2, g_strCONN, 1, 3
						Set rsPM = Server.CreateObject("ADODB.RecordSet")
						sqlPM = "SELECT * FROM Proj_Man_T WHERE ID = " & rsTBL2("PMID")
						rsPM.Open sqlPM, g_strCONN, 1, 3
						If Not rsPM.EOF Then
							PMname = rsPM("lname") & ", " & rsPM("fname")
						Else
							PMname = rsPM("ID")
						End If
						rsPM.Close
						Set rsPM = Nothing
						strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL2("phoneCon_last") & _
							"</td><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL2("lname") & ", " & _
							rsTBL2("fname") & "</td><td align='center'><font size='1' face='trebuchet ms'>" & _
							rsTBL2("Address") & ", " & rsTBL2("City") & ", " & rsTBL2("State") & ", " & rsTBL2("zip") & "</font></td>" & _
							"<td align='center'><font size='1' face='trebuchet ms'>&nbsp;" & rsTBL2("PhoneNo") & _
							"<td align='center'><font size='1' face='trebuchet ms'>&nbsp;" & _
							PMname & "</font></td></font></td><td align='center' width='300px'><font size='1' face='trebuchet ms'>" & _
							rsTBL2("pcom") & "</font></td></tr>"
					rsTBL2.Close
					zzz = zzz + 1
				Loop
				Set rsTBL2 = Nothing
			ElseIf Request("SelRep") = 8 Then '3
				Session("MSG") = "PCSP Worker Date of Hire report."
				typ = Request("SelRep")
				Srt = 1
				If Request("Sort") = 1 then Srt = 0
				strHEAD = "<tr bgcolor='#040C8B'><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Name</b>" & _
					"</font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Address</b>" & _
					"</font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Phone No.</b>" & _
					"</font></td><td align='center'>" & _
					"<a style='text-decoration: none;' href='JavaScript: SVSort();'><font size='1' face='trebuchet ms' color='white'><u>Date Of Hire</u></font></a></td>" & _
					"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Salary</b>" & _
					"</font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>RIHCC</b></font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>RIHCC2</b></font></td></tr>"
					'"</font></td><td align='center'><font size='1' face='trebuchet ms'>Consumer ID" & _
					'"</font></td></tr>"
					
					'<a style='text-decoration: none;' href='JavaScript: SVSort();'><font size='1' face='trebuchet ms' color='blue'>Date Of Hire</font>
				Set rsTBL = Server.CreateObject("ADODB.RecordSet")
				sqlTBL = "SELECT * FROM worker_t,Proj_Man_T WHERE status = 'Active' AND pm1 = proj_man_T.id ORDER BY Proj_Man_T.Lname, Proj_Man_T.Fname, month(date_hired), day(date_hired), worker_T.lname, worker_T.fname"
				If Request("Sort") = 1 Then 
					sqlTBL = "SELECT * FROM worker_t,Proj_Man_T WHERE status = 'Active' AND pm1 = proj_man_T.id ORDER BY month(date_hired), day(date_hired), Proj_Man_T.Lname, Proj_Man_T.Fname, worker_T.lname, worker_T.fname"
				End If
				If Request("Sort") = 1 Then sqlTBL = sqlTBL & " DESC"
				rsTBL.Open sqlTBL, g_strCONN, 1, 3
				Do Until rsTBL.EOF
					strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("worker_t.lname") & _
						", " & rsTBL("worker_t.fname") & "</font></td><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("Address") & _
						", " & rsTBL("City") & ", " & rsTBL("State") & ", " & rsTBL("Zip") & "</font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms'>&nbsp;" & rsTBL("PhoneNo") & _
						"</font></td><td align='center'><font size='1' face='trebuchet ms'>&nbsp;" & _
						rsTBL("Date_Hired") & "</td><td align='center'><font size='1' face='trebuchet ms'>&nbsp;" & Z_FormatNumber(rsTBL("Salary"), 2) & _
						"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & GetName3(rsTBL("PM1")) & "</font></td><td align='center'><font size='1' face='trebuchet ms'>" & GetName3(rsTBL("PM2")) & "</font></td></tr>"
						'<td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("CID") & "</font></td></tr>"
					rsTBL.MoveNext
				Loop
				rsTBL.Close
				Set rsTBL = Nothing
			ElseIf Request("SelRep") = 6 Then '4
				Session("MSG") = "Consumer by Town report."
				strHEAD = "<tr bgcolor='#040C8B'><td align='center'><font size='1' face='trebuchet ms' color='white'><b>RIHCC</b></font>" & _
					"</td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Name</b>" & _
					"</font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Town</b></font></td></tr>"
				Set rsTBL = Server.CreateObject("ADODB.RecordSet")
				sqlTBL = "SELECT * FROM Proj_Man_T, C_Status_t, Consumer_t WHERE " & _
					"PMID = Proj_Man_T.ID AND Consumer_t.Medicaid_number = C_Status_t.Medicaid_number AND Active = True " & _
					"ORDER BY Proj_Man_T.Lname, Proj_Man_T.Fname, City, Consumer_t.lname, Consumer_t.fname"
				rsTBL.Open sqlTBL, g_strCONN, 1, 3
				Do Until rsTBL.EOF
					Set rsPM = Server.CreateObject("ADODB.RecordSet")
					sqlPM = "SELECT * FROM Proj_Man_T WHERE ID = " & rsTBL("PMID")
					rsPM.Open sqlPM, g_strCONN, 1, 3
					If Not rsPM.EOF Then
						PMname = rsPM("lname") & ", " & rsPM("fname")
					Else
						PMname = rsPM("ID")
					End If
					rsPM.Close
					Set rsPM = Nothing
					
					strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & PMname & _
						"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("Consumer_t.lname") & _
						", " & rsTBL("Consumer_t.fname") & "</font></td><td align='center'><font size='1' face='trebuchet ms'>&nbsp;" & _
						rsTBL("City") & "</td></tr>"
					
					rsTBL.MoveNext
				Loop
				rsTBL.Close
				Set rsTBL = Nothing
			ElseIf Request("SelRep") = 7 Then '5
				Session("MSG") = "PCSP Worker by Town report."
				strHEAD = "<tr bgcolor='#040C8B'><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Town</b>" & _
					"</font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Name</b></font></td></tr>"
				Set rsTBL = Server.CreateObject("ADODB.RecordSet")
				sqlTBL = "SELECT lname, fname, city, state FROM Worker_t WHERE Status = 'Active' ORDER BY City, lname, fname"
				rsTBL.Open sqlTBL, g_strCONN, 1, 3
				Do Until rsTBL.EOF
					strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>&nbsp;" & _
						rsTBL("City") & "</td><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("lname") & _
						", " & rsTBL("fname") & "</font></td></tr>"
					rsTBL.MoveNext
				Loop
				rsTBL.Close
				Set rsTBL = Nothing
			ElseIf Request("SelRep") = 2 Then '6
				Session("MSG") = "All Active Consumer report."
				strHEAD = "<tr bgcolor='#040C8B'></font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Name</b> " &_
					"</font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Mailing Address</b></font></td><td align='center'>" & _
					"<font size='1' face='trebuchet ms' color='white'><b>Phone No.</b></font></td><td align='center'><font size='1' " & _
					"face='trebuchet ms' color='white'><b>DOB</b></font></td><td align='center'><font size='1' " & _
					"face='trebuchet ms' color='white'><b>Medicaid Number</b></font></td><td align='center'><font size='1' " & _
					"face='trebuchet ms' color='white'><b>Gender</b></font></td></tr>"
				Set rsTBL = Server.CreateObject("ADODB.RecordSet")
				sqlTBL = "SELECT lname, fname, maddress, mcity, mstate, mzip, DOB, PhoneNo, Consumer_t.Medicaid_Number, gender FROM Consumer_t, C_Status_t " & _
					"WHERE C_Status_t.Medicaid_number = Consumer_t.Medicaid_number AND Active = True ORDER BY lname, fname"
				rsTBL.Open sqlTBL, g_strCONN, 1, 3
				Do Until rsTBL.EOF
					strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("lname") & ", " & _
						rsTBL("fname") & "</font></td><td align='center'><font size='1' face='trebuchet ms'>" & _
						rsTBL("mAddress") & ", " & rsTBL("mCity") & ", " & rsTBL("mState") & ", " & rsTBL("mZip") & "</font></td><td align='center'>" & _
						"<font size='1' face='trebuchet ms'>&nbsp;" & rsTBL("PhoneNo") & "</td><td align='center'><font size='1' " & _
						"face='trebuchet ms'>&nbsp;" & rsTBL("DOB") & "</font></td>" 
						If Session("lngType") = 1 Or Session("lngType") = 2 Then 
							strBODY = strBODY & "<td align='center'><font size='1' " & _
								"face='trebuchet ms'>&nbsp;" & rsTBL("Medicaid_Number") & "</font></td>"
						Else
							strBODY = strBODY & "<Td>&nbsp;</td>"
						End If 
						strBODY = strBODY & "<td align='center'><font size='1' " & _
							"face='trebuchet ms'>&nbsp;" & rsTBL("Gender") & "</font></td></tr>"
					rsTBL.MoveNext
				Loop
				rsTBL.Close
				Set rsTBL = Nothing
			ElseIf Request("SelRep") = 3 Then '7
				Session("MSG") = "All Active PCSP Worker report."
				strHEAD = "<tr bgcolor='#040C8B'></font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Name</b> " &_
					"</font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Mailing Address</b></font></td><td align='center'>" & _
					"<font size='1' face='trebuchet ms' color='white'><b>Phone No.</b></font></td><td align='center'><font size='1' " & _
					"face='trebuchet ms' color='white'><b>DOB</b></font></td><td align='center'><font size='1' " & _
					"face='trebuchet ms' color='white'><b>File Number</b></font></td></tr>"
				Set rsTBL = Server.CreateObject("ADODB.RecordSet")
				sqlTBL = "SELECT lname, fname, maddress, mcity, mstate, mzip, DOB, PhoneNo, FileNum  FROM Worker_t  WHERE status = 'Active' " & _
					"ORDER BY lname, fname"
				rsTBL.Open sqlTBL, g_strCONN, 1, 3
				Do Until rsTBL.EOF
					strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("lname") & ", " & _
						rsTBL("fname") & "</font></td><td align='center'><font size='1' face='trebuchet ms'>" & _
						rsTBL("mAddress") & ", " & rsTBL("mCity") & ", " & rsTBL("mState") & ", " & rsTBL("mZip") & "</font></td><td align='center'>" & _
						"<font size='1' face='trebuchet ms'>" & rsTBL("PhoneNo") & "</td><td align='center'><font size='1' " & _
						"face='trebuchet ms'>" & rsTBL("DOB") & "</font></td><td align='center'><font size='1' " & _
						"face='trebuchet ms'>" & rsTBL("FileNum") & "</font></td></tr>"
					rsTBL.MoveNext
				Loop
				rsTBL.Close
				Set rsTBL = Nothing
			ElseIf Request("SelRep") = 12 Then '10
				Session("MSG") = "Consumer Referrals report."
				strHEAD = "<tr bgcolor='#040C8B'><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Referral Date</b>" & _
					"</font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Name</b></font></td>" & _
					"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Town</b></font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>RIHCC</b></font></td></tr>"
				Set rsTBL = Server.CreateObject("ADODB.RecordSet")
				sqlTBL = "SELECT lname, fname, city, Referral_Date, C_Status_t.Active, PMID FROM Consumer_t, C_Status_t " & _
					"WHERE Consumer_t.Medicaid_number = C_Status_t.Medicaid_number AND Active = True ORDER BY Referral_Date DESC, lname, fname"
				rsTBL.Open sqlTBL, g_strCONN, 1, 3
				Do Until rsTBL.EOF
					strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>&nbsp;" & _
						rsTBL("Referral_Date") & "</td><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("lname") & _
						", " & rsTBL("fname") & "</font></td><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("city") & _
						"</td><td align='center'><font size='1' face='trebuchet ms'>" & GetPM(rsTBL("PMID")) & "</td></tr>"
					rsTBL.MoveNext
				Loop
				rsTBL.Close
				Set rsTBL = Nothing
			ElseIf Request("SelRep") = 9 Then '12
				Session("MSG") = "PCSP Worker Interested in More Consumer report."
				strHEAD = "<tr bgcolor='#040C8B'><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Town</b>" & _
					"</font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Name</b></font></td></tr>"
				Set rsTBL = Server.CreateObject("ADODB.RecordSet")
				sqlTBL = "SELECT lname, fname, Towns FROM Worker_t, W_Towns_t WHERE Social_Security_Number = SSN " & _
					"AND Status = 'Active' ORDER BY Towns, lname, fname"
				rsTBL.Open sqlTBL, g_strCONN, 1, 3
				Do Until rsTBL.EOF
					strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>&nbsp;" & _
						rsTBL("Towns") & "</td><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("lname") & _
						", " & rsTBL("fname") & "</font></td></tr>"
					rsTBL.MoveNext
				Loop
				rsTBL.Close
				Set rsTBL = Nothing
			ElseIf Request("SelRep") = 1 Then '8
				Session("MSG") = "All Active Consumer with PCSP Worker and Hours report."
				strHEAD = "<tr bgcolor='#040C8B'><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Consumer Name</b>" & _
						"</font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Phone No.</b></font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Approved Hours</b></font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>PCSP Worker Name</b></font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Phone No.</b></font></td></tr>"
				Set rsLink = Server.CreateObject("ADODB.RecordSet")
				sqlLink = "SELECT * FROM ConWork_t, consumer_t WHERE CID = medicaid_number ORDER BY lname, fname"
				rsLink.Open sqlLink, g_strCONN, 1, 3
				Do Until rsLink.EOF
					Set rsCon = Server.CreateObject("ADODB.RecordSet")
					sqlCon = "SELECT * FROM Consumer_t, C_Status_t WHERE Consumer_t.medicaid_number = C_Status_t.medicaid_number " & _
					"AND Active = True AND Consumer_t.medicaid_number = '" & rsLink("CID") & "' "	
					rsCon.Open sqlCon, g_strCONN, 1, 3
					If Not rsCon.EOF Then
						strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & _
							rsCon("lname") & ", " & rsCon("fname") & "</font></td><td align='center'><font size='1' " & _
							"face='trebuchet ms'>&nbsp;" & rsCon("PhoneNo") & "</td><td align='center'><font size='1' " & _
							"face='trebuchet ms'>" & rsCon("MaxHrs") & "</td>" 
							Set rsWor = Server.CreateObject("ADODB.RecordSet")
							sqlWor = "SELECT * FROM Worker_t WHERE index = " & rsLink("WID")
							rsWor.Open sqlWor, g_strCONN, 1, 3
							If Not rsWor.EOF Then
								strBODY = strBODY & "<td align='center'><font size='1' face='trebuchet ms'>&nbsp;" & rsWor("lname") & ", " & _
									rsWor("fname") & "&nbsp;</td><td align='center'><font size='1' " & _
									"face='trebuchet ms'>&nbsp;" & rsWor("PhoneNo") & "</td>" 
							End If
							rsWor.Close
							Set rsWor = Nothing
							strBODY = strBODY & "</tr>"
					End If
					rsLink.MoveNext
				Loop
				rsLink.Close
				Set rsLink = Nothing
				rsCon.Close
				Set rsCon = Nothing	
			ElseIf Request("SelRep") = 4 Then '9
				Session("MSG") = "All Inactive Consumer with PCSP Worker and Hours report."
				strHEAD = "<tr bgcolor='#040C8B'><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Consumer Name</b>" & _
						"</font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Phone No.</b></font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Approved Hours</b></font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Inactive Date</b></font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Reason</b></font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>PCSP Worker Name</b></font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Phone No.</font></b></td></tr>"
				Set rsLink = Server.CreateObject("ADODB.RecordSet")
				sqlLink = "SELECT * FROM ConWork_t, consumer_t Where CID = medicaid_number ORDER BY lname, fname"
				rsLink.Open sqlLink, g_strCONN, 1, 3
				Do Until rsLink.EOF
					Set rsCon = Server.CreateObject("ADODB.RecordSet")
					sqlCon = "SELECT * FROM Consumer_t, C_Status_t WHERE Consumer_t.medicaid_number = C_Status_t.medicaid_number " & _
					"AND Active = False AND Consumer_t.medicaid_number = '" & rsLink("CID") & "' "	
					rsCon.Open sqlCon, g_strCONN, 1, 3
					If Not rsCon.EOF Then
						strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & _
							rsCon("lname") & ", " & rsCon("fname") & "</font></td><td align='center'><font size='1' " & _
							"face='trebuchet ms'>&nbsp;" & rsCon("PhoneNo") & "</td><td align='center'><font size='1' " & _
							"face='trebuchet ms'>" & rsCon("MaxHrs") & "</td><td align='center'><font size='1' face='trebuchet ms'>" & _
							rsCon("Inactive_date") & "</td><td align='center'><font size='1' face='trebuchet ms'>"
							
						'GET REASON
						tmpReas = ""
						If rsCon("Enter_Nursing_Home") = True Then tmpReas = tmpReas & "Enter Nursing Home or Other Setting<br>"
						If rsCon("Unable_Self_Direct") = True Then tmpReas = tmpReas & "Unable to Self-Direct<br>"
						If rsCon("Death") = True Then tmpReas = tmpReas & "Death<br>"
						If rsCon("Unable_Suitable_Worker") = True Then tmpReas = tmpReas & "Unable to Find Suitable Worker<br>"
						tmpReas = tmpReas & rsCon("A_Other")
						
						strBODY = strBODY & tmpReas & "</td>" 
						
							Set rsWor = Server.CreateObject("ADODB.RecordSet")
							sqlWor = "SELECT * FROM Worker_t WHERE index = " & rsLink("WID")
							rsWor.Open sqlWor, g_strCONN, 1, 3
							If Not rsWor.EOF Then
								strBODY = strBODY & "<td align='center'><font size='1' face='trebuchet ms'>" & rsWor("lname") & ", " & _
									rsWor("fname") & "</td><td align='center'><font size='1' " & _
									"face='trebuchet ms'>&nbsp;" & rsWor("PhoneNo") & "</td>" 
							End If
							rsWor.Close
							Set rsWor = Nothing
							strBODY = strBODY & "</tr>"
					End If
					rsLink.MoveNext
				Loop
				rsLink.Close
				Set rsLink = Nothing
				rsCon.Close
				Set rsCon = Nothing	
			ElseIf Request("SelRep") = 10 Then '11
				Session("MSG") = "PCSP Worker with No Consumer report."
				strHEAD = "<tr bgcolor='#040C8B'><td align='center'><font size='1' face='trebuchet ms' color='white'><b>SSN</b>" & _
						"</font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Name</b></font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Address</b></font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Phone No.</b></font></td></tr>"
				Set rsWor = Server.CreateObject("ADODB.RecordSet")
			sqlWor = "SELECT * FROM Worker_t WHERE [status] = 'Active' ORDER BY Lname, fname"
				rsWor.Open sqlWor, g_strCONN, 1, 3
				Do Until rsWor.EOF
					Set rsLink = Server.CreateObject("ADODB.RecordSet")
					sqlLink = "SELECT * FROM ConWork_t WHERE WID = '" & rsWor("index") & "' "
					rsLink.Open sqlLink, g_strCONN, 1, 3
					If rsLink.EOF Then
						strBODY = strBODY & "<tr>"
							If Session("lngType") = 1 Or Session("lngType") = 2 Then 
								strBODY = strBODY & "<td align='left'><font size='1' face='trebuchet ms'>&nbsp;" & _
								rsWor("Social_Security_Number") & "</td>"
							Else
							strBODY = strBODY & "<Td>&nbsp;</td>"
						End If 
						strBODY = strBODY & "<td align='center'><font size='1' face='trebuchet ms'>" & _
							rsWor("lname") & ", " & rsWor("fname") & "</font></td><td align='center'><font size='1' face='trebuchet ms'>" & _
							rsWor("Address") & ", " & rsWor("City") & ", " & rsWor("State") & ", " & rsWor("Zip") & "</font></td>" & _
							"<td align='center'><font size='1' face='trebuchet ms'>" &  rsWor("PhoneNo") & "</td></tr>"
					End If
					rsWor.MoveNext
				Loop
				rsLink.Close
				Set rsLink = Nothing
				rsWor.Close
				Set rsWor = Nothing
			ElseIf Request("SelRep") = 5 Then '13
				''''''A'''''''
				Set rsAct = Server.CreateObject("ADODB.RecordSet")
				sqlAct = "SELECT COUNT(C_Status_t.Medicaid_number) AS TActive FROM C_Status_t, Consumer_t WHERE Active = True" & _
					" AND C_Status_t.Medicaid_number = Consumer_t.Medicaid_number AND Year(Start_date) < Year(DateValue(Now)) - 1 "
				rsAct.Open sqlAct, g_strCONN, 1, 3
				TActive = rsAct("TActive")
				rsAct.Close
				Set rsAct = Nothing
				'''''B'''''''	
				Set rsIA = Server.CreateObject("ADODB.RecordSet")
				sqlIA = "SELECT COUNT(Medicaid_number) AS TIA FROM C_Status_t WHERE Active = False AND year(Inactive_date) = year(DateValue(now)) - 1"
				rsIA.Open sqlIA, g_strCONN, 1, 3
				TIA = rsIA("TIA")
				rsIA.Close
				Set rsIA = Nothing
				'''''C''''''
				Set rsNew = Server.CreateObject("ADODB.RecordSet")
				sqlNew = "SELECT COUNT(Medicaid_number) AS NEWC FROM Consumer_t WHERE Year(start_date) = Year(DateValue(Now)) - 1"
				rsNew.Open sqlNew, g_strCONN, 1, 3
				NEWC = rsNew("NEWC")
				rsNew.Close
				Set rsNew = Nothing
				'''''E''''''
				TCon = TActive + NEWC
				'''''D''''''
				TACT = TCon - TIA
				'''''F''''''
				Set rsTS = Server.CreateObject("ADODB.RecordSet")
				sqlTS = "SELECT * FROM Tsheets_t WHERE Year(date) = Year(DateValue(Now)) - 1"
				rsTS.Open sqlTS, g_strCONN, 1, 3
				ctr = 0
				Do Until rsTS.EOF
					If rsTS("mon") <> 0 Then ctr = ctr + 1
					If rsTS("tue") <> 0 Then ctr = ctr + 1
					If rsTS("wed") <> 0 Then ctr = ctr + 1
					If rsTS("thu") <> 0 Then ctr = ctr + 1
					If rsTS("fri") <> 0 Then ctr = ctr + 1
					If rsTS("sat") <> 0 Then ctr = ctr + 1
					If rsTS("sun") <> 0 Then ctr = ctr + 1
					rsTS.MoveNext
				Loop
				rsTS.Close
				Set rsTS = Nothing
				'''''G''''''
				LServ = CDbl(ctr / TCon)
				'''''REPORT'''''
				Session("MSG") = "Census Information for the year " & Year(Date) - 1 & " report."
				strBODY = "<tr><td bgcolor='#040C8B' height='70px' valign='center'><font size='1' face='trebuchet ms' color='white'>" & _
					"A.	NUMBER OF CONSUMERS ACTIVE IN<br>&nbsp;&nbsp;&nbsp;&nbsp;PROGRAM ON DECEMBER 31, " & Year(Date) - 2 & _
					"<br> &nbsp;&nbsp;&nbsp;&nbsp;CARRIED INTO " & Year(Date) - 1 & "</font></td><td align='center' " & _
					"width='100px'>" & TActive & "</td></tr>" & _
					"<tr><td bgcolor='#040C8B' height='70px' valign='center'><font size='1' face='trebuchet ms' color='white'>B.	" & _
					"NUMBER OF CONSUMERS DISCHARGED <br> &nbsp;&nbsp;&nbsp;&nbsp;FROM PROGRAM DURING " & Year(Date) - 1 & _
					"</font></td><td align='center'>" & _
					TIA & "</td></tr>" & _
					"<tr><td bgcolor='#040C8B' height='70px' valign='center'><font size='1' face='trebuchet ms' color='white'>C.	NUMBER OF NEW CONSUMERS ADMITTED <br>" & _
					" &nbsp;&nbsp;&nbsp;&nbsp;TO PROGRAM IN " & Year(Date) - 1 & "</font></td><td align='center'>" & _
					NEWC & "</td></tr>" & _
					"<tr><td bgcolor='#040C8B' height='70px' valign='center'><font size='1' face='trebuchet ms' color='white'>D.	NUMBER OF CONSUMERS ACTIVE IN <br>" & _
					" &nbsp;&nbsp;&nbsp;&nbsp;PROGRAM ON DECEMBER 31,  " & Year(Date) - 1& "</font></td><td align='center'>" & _
					TACT & "</td></tr>" & _
					"<tr><td bgcolor='#040C8B' height='70px' valign='center'><font size='1' face='trebuchet ms' color='white'>E.	TOTAL NUMBER CONSUMERS SERVED IN <br>&nbsp;&nbsp;&nbsp;&nbsp;" & _
					Year(Date) - 1 & "</font></td><td align='center'>" & _
					TCon & "</td></tr>" & _
					"<tr><td bgcolor='#040C8B' height='70px' valign='center'><font size='1' face='trebuchet ms' color='white'>F.	TOTAL CONSUMER DAYS IN " & _
					Year(Date) - 1 & "</font></td><td align='center'>" & _
					ctr & "</td></tr>" & _
					"<tr><td bgcolor='#040C8B' height='70px' valign='center'><font size='1' face='trebuchet ms' color='white'>G.	AVERAGE LENGTH OF SERVICE IN PROGRAM <br>" & _
					"&nbsp;&nbsp;&nbsp;&nbsp;IN " & Year(Date) - 1 & "</font></td><td align='center'>" & _
					Z_FormatNumber(LServ, 2) & "</td></tr>"
				ElseIf Request("SelRep") = 14 Then 
					Session("MSG") = "All Active Consumers Sorted by RIHCC And Town report."
					strHEAD = "<tr bgcolor='#040C8B'><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Consumer Name</b>" & _
						"</font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Address</b></font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Phone No.</b></font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>RIHCC</b></font></td></tr>"
					Set rsTBL = Server.CreateObject("ADODB.RecordSet")
				'sqlTBL = "SELECT * " & _
				'	" FROM Consumer_t, C_Status_t, Proj_man_T " & _ 
				'	" WHERE PMID = Proj_Man_T.ID AND Consumer_t.medicaid_number = C_Status_t.medicaid_number" & _
				'	" AND Active = True ORDER BY Consumer_t.lname, Consumer_t.fname, Proj_Man_T.Lname, Proj_Man_T.Fname, CiTY"
				sqlTBL = "SELECT * " & _
					" FROM Consumer_t, C_Status_t, Proj_man_T " & _ 
					" WHERE PMID = Proj_Man_T.ID AND Consumer_t.medicaid_number = C_Status_t.medicaid_number" & _
					" AND Active = True ORDER BY Proj_Man_T.Lname, Proj_Man_T.Fname, Consumer_t.lname, Consumer_t.fname, CiTY "
				rsTBL.Open sqlTBL, g_strCONN, 1, 3
				Do Until rsTBL.EOF
					Set rsPM = Server.CreateObject("ADODB.RecordSet")
					sqlPM = "SELECT * FROM Proj_Man_T WHERE ID = " & rsTBL("PMID")
					rsPM.Open sqlPM, g_strCONN, 1, 3
					If Not rsPM.EOF Then
						PMname = rsPM("lname") & ", " & rsPM("fname")
					Else
						PMname = rsPM("ID")
					End If
					rsPM.Close
					Set rsPM = Nothing
					strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("Consumer_t.lname") & ", " & _
						rsTBL("Consumer_t.fname") & "</font></td><td align='center'><font size='1' face='trebuchet ms'>" & _
						rsTBL("Address") & ", " & rsTBL("City") & ", " & rsTBL("State") & ", " & rsTBL("Zip") & "</font></td><td align='center'>" & _
						"<font size='1' face='trebuchet ms'>&nbsp;" & rsTBL("PhoneNo") & "</td><td align='center'><font size='1' " & _
						"face='trebuchet ms'>&nbsp;" & PMname & "</font></td></tr>"
					rsTBL.MoveNext
				Loop
				rsTBL.Close
				Set rsTBL = Nothing
			ElseIf Request("SelRep") = 15 Then 
				Session("MSG") = "PCSP Worker by Insurance Expiration Date report."
				strHEAD = "<tr bgcolor='#040C8B'><td align='center'><font size='1' face='trebuchet ms' color='white'><b>PCSP Worker Name</b>" & _
						"</font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Address</b></font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Insurance Expiration Date</b></font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>RIHCC</b></font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>RIHCC2</b></font></td></tr>"
				Set rsTBL = Server.CreateObject("ADODB.RecordSet")
				sqlTBL = "SELECT * FROM worker_t WHERE status = 'Active' AND Driver = True ORDER BY insuranceexpdate, worker_t.lname, worker_t.fname"
				rsTBL.Open sqlTBL, g_strCONN, 1, 3
				Do Until rsTBL.EOF	
					strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("lname") & ", " & _
						rsTBL("fname") & "</font></td><td align='center'><font size='1' face='trebuchet ms'>" & _
						rsTBL("Address") & ", " & rsTBL("City") & ", " & rsTBL("State") & ", " & rsTBL("Zip") & "</font></td><td align='center'>" & _
						"<font size='1' face='trebuchet ms'>" & rsTBL("Insuranceexpdate") & "</td><td align='center'>" & _
						"<font size='1' face='trebuchet ms'>" & GetName3(rsTBL("PM1")) & "</td><td align='center'>" & _
						"<font size='1' face='trebuchet ms'>" & GetName3(rsTBL("PM2")) & "</td></tr>"
					rsTBL.MoveNext
				Loop
				rsTBL.Close
				Set rsTBL = Nothing			
			ElseIf Request("SelRep") = 16 Then
				Session("MSG") = "Consumer Health report."
				strHEAD = "<tr bgcolor='#040C8B'><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Consumer Name</b>" & _
						"</font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Age</b></font></td><td align='center'>" & _
						"<font size='1' face='trebuchet ms' color='white'><b>Gender</b></font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Rating</b></font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Ambulation</b></font></td><td align='center'><font size='1' face='trebuchet ms' color='white'>" & _
						"<b>ADL</b></font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Others</b></font></td><td align='center'><font size='1' " & _
						"face='trebuchet ms' color='white'><b>Diagnosis</b></font></td></tr>"
				Set rsTBL = Server.CreateObject("ADODB.RecordSet")
				sqlTBL = "SELECT consumer_t.medicaid_number, lname, fname, DOB, gender, active, c_health_t.* FROM consumer_t, " & _
					"c_health_t, c_status_t WHERE consumer_t.medicaid_number = c_status_t.medicaid_number AND " & _
					"consumer_t.medicaid_number = c_health_t.medicaid_number " & _
					"AND active = true ORDER BY lname, fname"
				rsTBL.Open sqlTBL, g_strCONN, 1, 3
				tmpIDx = ""
				Do Until rsTBL.EOF	
					Age = DateDiff("yyyy", rsTBL("DOB"), Now)
					
					Amb = ""
					If rsTBL("Indept") = true Then Amb = "Independent"
					If rsTBL("cane") = true Then Amb = "Cane"
					If rsTBL("walker") = true Then Amb = "Walker"
					If rsTBL("walk") = true Then Amb = "Walk/Wheel"
					If rsTBL("wheelchair") = true Then Amb = "Wheelchair"
					
					ADL = ""
					If rsTBL("ADL_Indep") = true Then ADL = "Independent"
					If rsTBL("monitor") = true Then ADL = "Monitor"
					If rsTBL("minass") = true Then ADL = "Min. Assistance"
					If rsTBL("ass") = true Then ADL = "Assistance"
					If rsTBL("Complete") = true Then ADL = "Complete Care"
					
					OQC = ""
					If rsTBL("Use") = true Then OQC = "O<sub>2</sub> Use<br> "
					If rsTBL("mental_h") = true Then OQC = OQC & "Mental Health<br> "
					If rsTBL("drug") = true Then OQC = OQC & "Drug Use/Abuse<br> "
					If rsTBL("iso") = true Then OQC = OQC & "Isolation<br> "
					If rsTBL("dem") = true Then OQC = OQC & "Dementia/Alzheimer's<br> "
					If rsTBL("terminal") = true Then OQC = OQC & "Terminal/Hospice<br> "
					If rsTBL("tob") = true Then OQC = OQC & "Tobacco Use<br> "
					If rsTBL("obese") = true Then OQC = OQC & "Obesity<br>"
					If rsTBL("Para") = true Then OQC = OQC & "Paralysis<br> "
					If rsTBL("Quad") = true Then OQC = OQC & "Quadriplegic"
						
					strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("lname") & ", " & _
						rsTBL("fname") & "</font></td><td align='center'><font size='1' face='trebuchet ms'>" & _
						Age & "</font></td><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("Gender")  & "</font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("Rating")  & "</font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms'>" & Amb  & "</font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms'>" & ADL  & "</font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms'>&nbsp;" & OQC  & "</font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms'>"
					
					Set rsDiag = Server.CreateObject("ADODB.RecordSet")
					sqlDiag = "SELECT * FROM c_Diagnosis_t WHERE c_diagnosis_t.Medicaid_number = '" & rsTBL("consumer_t.medicaid_number") & "' "
					'response.write sqlDiag & vbcrlf
					rsDiag.Open sqlDiag, g_strCONN, 1, 3
					If rsDiag.EOF Then
						strBODY = strBODY & "&nbsp;</font></td></tr>"
					Else
						Do Until rsDiag.EOF
							strBODY = strBODY & rsDiag("Diagnosis") & "<br> "
							rsDiag.MoveNext
						Loop
					End If
					rsDiag.Close
					Set rsDiag = Nothing
					strBODY = strBODY & "</font></td></tr>"
					rsTBL.MoveNext
				Loop
				rsTBL.Close
				Set rsTBL = Nothing			
			ElseIf Request("SelRep") = 17 Then 
				Session("MSG") = "Consumer Start Date report."
				strHEAD = "<tr bgcolor='#040C8B'><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Consumer Name</b>" & _
						"</font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Start Date</b></font></td></tr>"
				Set rsTBL = Server.CreateObject("ADODB.RecordSet")
				sqlTBL = "SELECT fname, lname, start_date, active  FROM consumer_t, c_status_t WHERE " & _
					"Consumer_t.medicaid_number = c_status_t.medicaid_number AND Active = True ORDER BY start_date, lname, fname"
				rsTBL.Open sqlTBL, g_strCONN, 1, 3
				Do Until rsTBL.EOF	
					strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("lname") & ", " & _
						rsTBL("fname") & "</font></td><td align='center'><font size='1' face='trebuchet ms'>&nbsp;" & _
						rsTBL("start_date") & "</font></td></tr>"
					rsTBL.MoveNext
				Loop
				rsTBL.Close
				Set rsTBL = Nothing			
			ElseIf Request("SelRep") = 18 Then 
				Session("MSG") = "Consumer Start and End Date report"
				strHEAD = "<tr bgcolor='#040C8B'><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Consumer Name</b>" & _
						"</font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Start Date</b></font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>End Date</b></font></td></tr>"
				Set rsTBL = Server.CreateObject("ADODB.RecordSet")
				sqlTBL = "SELECT * FROM consumer_t, c_status_t WHERE consumer_t.medicaid_number = c_status_t.medicaid_number AND Active = true"
				If Request("FrmD8") <> "" Then
					If IsDate(Request("FrmD8")) Then
						sqlTBL = sqlTBL & " AND Start_date >= #" & Request("FrmD8") & "# " 
						Session("Msg") = Session("Msg") & " from " & Request("FrmD8")
					End If
				End If
				If Request("ToD8") <> "" Then
					If IsDate(Request("ToD8")) Then
						sqlTBL = sqlTBL & " AND Start_date <= #" & Request("ToD8") & "# " 
						Session("Msg") = Session("Msg") & " to " & Request("ToD8")
					End If
				End If
				sqlTBL  = sqlTBL  & " ORDER BY lname, fname"
				Session("Msg") = Session("Msg") & ". " 
				'response.write sqlTBL
				rsTBL.Open sqlTBL, g_strCONN, 1, 3
				Do Until rsTBL.EOF	
					strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & _
						rsTBL("lname") & ", " & rsTBL("fname") &"</font></td><td align='center'><font size='1' face='trebuchet ms'>&nbsp;" & rsTBL("Start_date") & _
						"</font></td><td align='center'><font size='1' face='trebuchet ms'>&nbsp;" & _
						rsTBL("Termdate") &  "</font></td></tr>"
					rsTBL.MoveNext
				Loop
				rsTBL.Close
				Set rsTBL = Nothing	
			ElseIf Request("SelRep") = 19 Then 
				Session("MSG") = "Consumer Date Of Birth report."
				strHEAD = "<tr bgcolor='#040C8B'><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Date Of Birth</b>" & _
						"</font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Name</b></font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Address</b></font></td></tr>"
				Set rsTBL = Server.CreateObject("ADODB.RecordSet")
				sqlTBL = "SELECT DOB, fname, lname, address, city, state, zip, active FROM consumer_t, c_status_t " & _
					"WHERE Consumer_t.medicaid_number = c_status_t.medicaid_number AND Active = True ORDER BY Month(DOB), Day(DOB), lname, fname"
				rsTBL.Open sqlTBL, g_strCONN, 1, 3
				Do Until rsTBL.EOF	
					strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>&nbsp;" & _
						rsTBL("DOB") & "</font></td><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("lname") & _
						", " & rsTBL("fname") & "</font></td><td align='center'><font size='1' face='trebuchet ms'>" & _
						rsTBL("Address") & ", " & rsTBL("City") & ", " & rsTBL("State") & ", " & rsTBL("Zip") & "</font></td></tr>"
					rsTBL.MoveNext
				Loop
				rsTBL.Close
				Set rsTBL = Nothing	
			
			ElseIf Request("SelRep") = 20 Then 
				Session("MSG") = "All Active PCSP Worker by RIHCC And Town report."
				strHEAD = "<tr bgcolor='#040C8B'><td align='center'><font size='1' face='trebuchet ms' color='white'><b>PCSP Worker Name</b>" & _
						"</font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Address</b></font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Phone No.</b></font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>RIHCC1</b></font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>RIHCC2</b></font></td></tr>"
				Set rsTBL = Server.CreateObject("ADODB.RecordSet")
				sqlTBL = "SELECT * FROM worker_t " & _
					"WHERE status = 'Active' ORDER BY worker_t.lname, worker_T.fname, worker_t.City"

				rsTBL.Open sqlTBL, g_strCONN, 1, 3
				Do Until rsTBL.EOF	
					PMname = ""
					PMname2 = ""

					strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("lname") & ", " & _
						rsTBL("fname") & "</font></td><td align='center'><font size='1' face='trebuchet ms'>" & _
						rsTBL("Address") & ", " & rsTBL("City") & ", " & rsTBL("State") & ", " & _
						rsTBL("Zip") & "</font></td><td align='center'>" & _
						"<font size='1' face='trebuchet ms'>&nbsp;" & rsTBL("PhoneNo") & "</td><td align='center'>" & _
						"<font size='1' face='trebuchet ms'>" & GetName3(rsTBL("PM1")) & "</td><td align='center'>" & _
						"<font size='1' face='trebuchet ms'>" & GetName3(rsTBL("PM2")) & "</td></tr>"
					rsTBL.MoveNext
				Loop
				rsTBL.Close
				Set rsTBL = Nothing	
			ElseIf Request("SelRep") = 21 Then 
				Session("MSG") = "PCSP Worker (Inactive) Termination Date report"
				strHEAD = "<tr bgcolor='#040C8B'><td align='center'><font size='1' face='trebuchet ms' color='white'><b>PCSP Worker Name</b>" & _
						"</font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Termination Date</b></font></td></tr>"	
				Set rsTBL = Server.CreateObject("ADODB.RecordSet")
				sqlTBL = "SELECT * FROM worker_t WHERE Status = 'InActive' ORDER BY Term_date, lname, fname"
				rsTBL.Open sqlTBL, g_strCONN, 1, 3
				Do Until rsTBL.EOF	
					strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("lname") & ", " & _
						rsTBL("fname") & "</font></td><td align='center'>" & _
						"<font size='1' face='trebuchet ms'>&nbsp;" & rsTBL("Term_date") & "</td></tr>"
					rsTBL.MoveNext
				Loop
				rsTBL.Close
				Set rsTBL = Nothing	
			ElseIf Request("SelRep") = 22 Then 
				Session("MSG") = "Phone Contact for PCSP Worker report."
				typ = Request("SelRep")
				strHEAD = "<tr bgcolor='#040C8B'><td align='center'><font size='1' face='trebuchet ms' color='white' ><b>Phone Contact Date</b>" & _
					"</font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Name</b></font></td><td align='center'>" & _
					"<font size='1' face='trebuchet ms' color='white'><b>Address</b></font></td><td align='center'><font size='1' " & _
					"face='trebuchet ms' color='white'><b>Phone No.</b></font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>RIHCC</b></font>" & _
					"</td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>RIHCC2</b></font>" & _
					"</td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Comment</b></font>" & _
					"</td></tr>"
				
				Set rsTBL = Server.CreateObject("ADODB.RecordSet")
		
				'sqlTBL = "SELECT * FROM Worker_t, w_log_t " & _
				'		"WHERE Worker_t.social_security_number = w_log_t.ssn AND " & _
				'		"status = 'Active' AND NOT IsNull(phonec) ORDER BY Worker_t.social_security_number, phonec DESC"
						
				sqlTBL = "SELECT * " & _ 
					"FROM Worker_t, w_log_t, proj_man_t  " & _
					"WHERE social_security_number = ssn " & _
					"AND PM1 = Proj_Man_T.ID " & _
					"AND status = 'Active' " & _
					"AND NOT IsNull(phonec) " & _
					"ORDER BY Proj_Man_T.Lname, Proj_Man_T.Fname,ssn, worker_T.lname ASC, phonec DESC"
				
				rsTBL.Open sqlTBL, g_strCONN, 1, 3
				
				tmpIDx = ""
				x = 0
				
				Do Until rsTBL.EOF
					If rsTBL("social_security_number") <> tmpIDx then 
						ReDim Preserve tmp(x)
						tmp(x) = rsTBL("w_log_t.index") & "|" & rsTBL("phoneC") & "|" & rsTBL("PM1") & "|" & rsTBL("PM2")
						'tmp(x) = rsTBL("w_log_t.index") & "|" & rsTBL("phoneC")
						x = x + 1
					End If
					tmpIDx = rsTBL("social_security_number")
					rsTBL.MoveNext
				Loop
				
				For i = x - 2 to 0 Step - 1
					For j = 0 To i
						tmpj = split(tmp(j),"|")
						tmpj1 = split(tmp(j+1),"|")
						If Cdate(tmpj(1)) < Cdate(tmpj1(1)) AND tmpj(2) = tmpj1(2) Then
						'If Cdate(tmpj(1)) < Cdate(tmpj1(1)) AND tmpj(2) = tmpj1(2) AND tmpj(3) = tmpj1(3) Then
						'If Cdate(tmpj(1)) < Cdate(tmpj1(1)) Then
							intTemp = tmp(j + 1)
			              tmp(j + 1) = tmp(j)
			              tmp(j) = intTemp
						End If
					Next 
				Next 
							
				rsTBL.Close
				Set rsTBL = Nothing
				
				Set rsTBL2 = Server.CreateObject("ADODB.RecordSet")
				zzz = 0
				Do Until zzz = x 
					
					tmp2 = split(tmp(zzz),"|")	
					
					sqlTBL2 = "SELECT * FROM Worker_t, w_log_t " & _
						"WHERE Worker_t.social_security_number = w_log_t.ssn AND " & _
						"status = 'Active' AND NOT IsNull(phonec) AND w_log_t.index = " & tmp2(0)
					
					rsTBL2.Open sqlTBL2, g_strCONN, 1, 3
					'''PM name
					'	Set rsPM = Server.CreateObject("ADODB.RecordSet")
					'	sqlPM = "SELECT * FROM Proj_Man_T WHERE ID = " & tmp2(2)
						'sqlPM = "SELECT * FROM Proj_Man_T WHERE ID = " & rsTBL2("PMID")
						'response.write sqlPM
					'	rsPM.Open sqlPM, g_strCONN, 1, 3
					'	If Not rsPM.EOF Then
					'		PMname = rsPM("lname") & ", " & rsPM("fname")
					'	Else
					'		PMname = rsPM("ID")
					'	End If
					'	rsPM.Close
					'	Set rsPM = Nothing
						strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL2("phonec") & _
								"</td><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL2("lname") & ", " & _
								rsTBL2("fname") & "</font></td><td align='center'><font size='1' face='trebuchet ms'>" & _
								rsTBL2("Address") & ", " & rsTBL2("City") & ", " & rsTBL2("State") & ", " & rsTBL2("zip") & "</font></td><td align='center'>" & _
								"<font size='1' face='trebuchet ms'>&nbsp;" & rsTBL2("PhoneNo") & "</font></td><td align='center'>" & _
								"<font size='1' face='trebuchet ms'>" & GetName3(tmp2(2)) & "</font></td><td align='center'>" & _
								"<font size='1' face='trebuchet ms'>" & GetName3(tmp2(3)) & "</font></td><td align='center' width='300px'>" & _
								"<font size='1' face='trebuchet ms'>" & rsTBL2("pcom") & "</font></td></tr>"
					rsTBL2.Close
					zzz =zzz + 1
				Loop
				Set rsTBL2 = Nothing
			
			ElseIf Request("SelRep") = 23 Then '1
				Session("MSG") = "Site Visit for PCSP Worker report."
				typ = Request("SelRep")
				
				strHEAD = "<tr bgcolor='#040C8B'><td align='center'><font size='1' face='trebuchet ms' color='white' ><b>Site Visit Date</b>" & _
					"</font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Name</b></font></td><td align='center'>" & _
					"<font size='1' face='trebuchet ms' color='white'><b>Address</b></font></td><td align='center'><font size='1' " & _
					"face='trebuchet ms' color='white'><b>Phone No.</b></font></td><td align='center'>" & _
					"<font size='1' face='trebuchet ms' color='white'><b>RIHCC</b></font></td><td align='center'>" & _
					"<font size='1' face='trebuchet ms' color='white'><b>RIHCC2</b></font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Comment</b></font>" & _
					"</td></tr>"
				
				Set rsTBL = Server.CreateObject("ADODB.RecordSet")
		
				'sqlTBL = "SELECT * FROM Worker_t, w_log_t " & _
				'		"WHERE Worker_t.social_security_number = w_log_t.ssn AND " & _
				'		"status = 'Active' AND NOT IsNull(sitev) ORDER BY Worker_t.social_security_number, sitev DESC"
				
				sqlTBL = "SELECT * " & _ 
					"FROM Worker_t, w_log_t, consumer_t, conwork_t, proj_man_t  " & _
					"WHERE Worker_t.social_security_number = w_log_t.ssn " & _
					"AND PM1 = Proj_Man_T.ID " & _
					"AND consumer_t.medicaid_number = CID " & _
					"AND CStr(worker_t.index) = WID " & _
					"AND status = 'Active' " & _
					"AND NOT IsNull(sitev) " & _
					"ORDER BY Proj_Man_T.Lname, Proj_Man_T.Fname, worker_T.social_security_number ASC, sitev DESC"
							
				rsTBL.Open sqlTBL, g_strCONN, 1, 3
				
				tmpIDx = ""
				x = 0
				
				Do Until rsTBL.EOF
					If rsTBL("worker_t.social_security_number") <> tmpIDx then 
						ReDim Preserve tmp(x)
						tmp(x) = rsTBL("w_log_t.index") & "|" & rsTBL("sitev") & "|" & rsTBL("PM1") & "|" & rsTBL("PM2") 
						x = x + 1
					End If
					tmpIDx = rsTBL("worker_t.social_security_number")
					rsTBL.MoveNext
				Loop
				
				
					For i = x - 2 to 0 Step - 1
						For j = 0 To i
							tmpj = split(tmp(j),"|")
							tmpj1 = split(tmp(j+1),"|")
							'If Cdate(tmpj(1)) < Cdate(tmpj1(1)) Then
							If Cdate(tmpj(1)) < Cdate(tmpj1(1)) AND tmpj(2) = tmpj1(2) Then
								intTemp = tmp(j + 1)
				              tmp(j + 1) = tmp(j)
				              tmp(j) = intTemp
							End If
						Next 
					Next 
			
				rsTBL.Close
				Set rsTBL = Nothing
				
				Set rsTBL2 = Server.CreateObject("ADODB.RecordSet")
				zzz = 0
				Do Until zzz = x 
					
					tmp2 = split(tmp(zzz),"|")	
					
					sqlTBL2 = "SELECT * FROM Worker_t, w_log_t " & _
						"WHERE Worker_t.social_security_number = w_log_t.ssn AND " & _
						"status = 'Active' AND NOT IsNull(sitev) AND w_log_t.index = " & tmp2(0)
					
					rsTBL2.Open sqlTBL2, g_strCONN, 1, 3
						'Set rsPM = Server.CreateObject("ADODB.RecordSet")
						'sqlPM = "SELECT * FROM Proj_Man_T WHERE ID = " & rsTBL2("PMID")
						'sqlPM = "SELECT * FROM Proj_Man_T WHERE ID = " & tmp2(2)
						'rsPM.Open sqlPM, g_strCONN, 1, 3
						'If Not rsPM.EOF Then
						'	PMname = rsPM("lname") & ", " & rsPM("fname")
						'Else
						'	PMname = rsPM("ID")
						'End If
						'rsPM.Close
						'Set rsPM = Nothing
						strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL2("sitev") & _
								"</td><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL2("lname") & ", " & _
								rsTBL2("fname") & "</font></td><td align='center'><font size='1' face='trebuchet ms'>" & _
								rsTBL2("Address") & ", " & rsTBL2("City") & ", " & rsTBL2("State") & ", " & rsTBL2("zip") & "</font></td><td align='center'>" & _
								"<font size='1' face='trebuchet ms'>" & rsTBL2("PhoneNo") & "</font></td><td align='center'>" & _
								"<font size='1' face='trebuchet ms'>" & GetName3(tmp2(2)) & "</font></td><td align='center'>" & _
								"<font size='1' face='trebuchet ms'>" & GetName3(tmp2(3)) & "</font></td><td align='center' width='300px'>" & _
								"<font size='1' face='trebuchet ms'>" & rsTBL2("scom") & "</font></td></tr>"
								'add PM
					rsTBL2.Close
					zzz =zzz + 1
				Loop
				Set rsTBL2 = Nothing
		ElseIf Request("SelRep") = 24 Then '1
			Session("MSG") = "Total Hours for PCSP Worker report "
			sqlProc = "SELECT * FROM [Tsheets_t], [worker_t] WHERE emp_id = Social_Security_Number"
			Err = 0
			If Request("FrmD8") <> "" Then
					If IsDate(Request("FrmD8")) Then
						sqlProc = sqlProc & " AND date >= #" & dateAdd("d", -7, Request("FrmD8")) & "# "
						Session("Msg") = Session("Msg") & " from " & Request("FrmD8")
					Else
						Err = 1
					End If
			End If
			If Request("ToD8") <> "" Then
					If IsDate(Request("ToD8")) Then
						sqlProc = sqlProc & " AND date  <= #" & dateAdd("d", 7, Request("ToD8")) & "#" 
						Session("Msg") = Session("Msg") & " to " & Request("ToD8")
					Else
						Err = 1
					End If
			End If
			Session("Msg") = Session("Msg") & "<br> * Extended Hours"
			If Err <> 0 Then Response.Redirect "specrep.asp?err=24" 
			Set rsProc = Server.CreateObject("ADODB.RecordSet")
			
			sqlProc = sqlProc & " ORDER BY lname ASC, fname ASC, date DESC, ID DESC"
			
			rsProc.Open sqlProc, g_strCONN, 3, 1
			tmpEID = ""
			If Not rsProc.EOF Then
				Do Until rsProc.EOF
					tmpName = GetName(rsProc("emp_id"))
					tmpName2 = GetName2(rsProc("client"))
					THours = rsProc("mon") + rsProc("tue") + rsProc("wed") + rsProc("thu") + rsProc("fri") + rsProc("sat") + rsProc("sun")
					FtotHrs = 0
					If THours <> 0 Then 
						tmphrsMon = ValidDate(Request("FrmD8"), Request("ToD8"), rsProc("date"), rsProc("mon"), "MON")
            tmphrsTue = ValidDate(Request("FrmD8"), Request("ToD8"), rsProc("date"), rsProc("tue"), "TUE")
            tmphrsWed = ValidDate(Request("FrmD8"), Request("ToD8"), rsProc("date"), rsProc("wed"), "WED")
            tmphrsThu = ValidDate(Request("FrmD8"), Request("ToD8"), rsProc("date"), rsProc("thu"), "THU")
            tmphrsFri = ValidDate(Request("FrmD8"), Request("ToD8"), rsProc("date"), rsProc("fri"), "FRI")
            tmphrsSat = ValidDate(Request("FrmD8"), Request("ToD8"), rsProc("date"), rsProc("sat"), "SAT")
            tmphrsSun = ValidDate(Request("FrmD8"), Request("ToD8"), rsProc("date"), rsProc("sun"), "SUN")
            FtotHrs = tmphrsMon + tmphrsTue + tmphrsWed + tmphrsThu + tmphrsFri + tmphrsSat + tmphrsSun
          End If
          'If tmpEID = "" And FtotHrs <> 0 Then 
        	'	strBODY = strBODY & "<tr bgcolor='#040C8B'><td align='left' colspan='2'><font size='1' face='trebuchet ms' color='white'>PCSP Worker:<b> " & tmpName & _
					'		"</b></font></td><td align='right' colspan='3'><font size='1' face='trebuchet ms' color='white'>Social Security Number:<b> " & Right(rsProc("emp_id"), 4) & _
					'		"</b></font></td></tr>"
					'	
					'	strBODY = strBODY &	"<tr bgcolor='#040C8B'><td align='center' width='100px'><font size='1' face='trebuchet ms' color='white'><b>Medicaid</b></font></td>" & _
					'		"<td align='center' width='150px'>" & vbCrLf & _
					'		"<font size='1' face='trebuchet ms' color='white'><b>Name</b></font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Timesheet Week</b></font></td>" & vbCrLf & _
					'		"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Hours</b></font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Notes</b></font></td></tr>"
        	'End If
					If rsProc("emp_id") <> tmpEID Then 'And FtotHrs <> 0 Then
					
						strBODY = strBODY & "<tr bgcolor='#040C8B'><td align='left' colspan='2'><font size='1' face='trebuchet ms' color='white'>PCSP Worker:<b> " & tmpName & _
							"</b></font></td><td align='right' colspan='3'><font size='1' face='trebuchet ms' color='white'>Social Security Number:<b> " & Right(rsProc("emp_id"), 4) & _
							"</b></font></td></tr>"
						
						strBODY = strBODY &	"<tr bgcolor='#040C8B'><td align='center' width='100px'><font size='1' face='trebuchet ms' color='white'><b>Medicaid</b></font></td>" & _
							"<td align='center' width='150px'>" & vbCrLf & _
							"<font size='1' face='trebuchet ms' color='white'><b>Name</b></font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Timesheet Week</b></font></td>" & vbCrLf & _
							"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Hours</b></font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Notes</b></font></td></tr>"
					
					End If
					If FtotHrs <> 0 Then
						tmpTSWk1 = rsProc("date")
						If Request("FrmD8") <> "" Then
							If Cdate(tmpTSWk1) < Cdate(Request("FrmD8")) Then tmpTSWk1 = Request("FrmD8")
						End If
						tmpTSWk2 = Cdate(rsProc("date")) + 6
						If Request("ToD8") <> "" Then
							If Cdate(tmpTSWk2) > Cdate(Request("ToD8")) Then tmpTSWk2 = Request("ToD8")
						End If
						If rsProc("EXT") = True Then
							mark = 1
							strBODY = strBODY & "<tr>"
							If Session("lngType") = 1 Or Session("lngType") = 2 Then 
								strBODY = strBODY & "<td align='center'><font size='1' face='trebuchet ms'>" & rsProc("client") & "</font>" & vbCrLf & _
								"</td>"
								Else
								strBODY = strBODY & "<Td>&nbsp;</td>"
							End If 	
								strBODY = strBODY & "<td align='center'><font size='1' face='trebuchet ms'><b>*</b>" & tmpName2 & "</font></td><td align='center'>" & vbCrLf & _
								"<font size='1' face='trebuchet ms'>" & tmpTSWk1 & " - " & tmpTSWk2 & "</font></td>" & vbCrLf & _
								"<td align='center'><font size='1' face='trebuchet ms'>" & Z_FormatNumber(FtotHrs,2) & "</font></td><td align='center'>" & _
								"<textarea rows='1' readonly style='font-size: 10;' >" & rsProc("misc_notes") & "</textarea></td></tr>"
						Else
							tmpMSG = ""
							If mark = 0 Then tmpMSG = rsProc("misc_notes")
							strBODY = strBODY & "<tr>"
							If Session("lngType") = 1 Or Session("lngType") = 2 Then 
								strBODY = strBODY & "<td align='center'><font size='1' face='trebuchet ms'>" & rsProc("client") & "</font></td>"
							Else
							strBODY = strBODY & "<Td>&nbsp;</td>"
						End If 
						strBODY = strBODY & "<td align='center'><font size='1' face='trebuchet ms'>" & tmpName2 & "</font></td><td align='center'><font size='1' face='trebuchet ms'>" & vbCrLf & _
								tmpTSWk1 & " - " & tmpTSWk2 & "</font></td><td align='center'><font size='1' face='trebuchet ms'>" & vbCrLf & _
								Z_FormatNumber(FtotHrs,2) & "</font></td><td align='center'>" & vbCrLf & _
								"<textarea readonly  rows='1' style='font-size: 10;' >" & tmpMSG & "</textarea></td></tr>"
							mark = 0
						End If
					End If
					tmpEID = rsProc("emp_id")
					rsProc.MoveNext
				Loop
			Else
				Session("MSG") = Session("MSG") & "<br> No records found."
			End If
			rsProc.Close
			Set rsProc = Nothing
			
		ElseIf Request("SelRep") = 25 Then
			Session("MSG") = "Active PCSP Workers with no log report. "
			
			strHEAD = "<tr bgcolor='#040C8B'><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Name</b></font></td><td align='center'>" & _
					"<font size='1' face='trebuchet ms' color='white'><b>Address</b></font></td><td align='center'><font size='1' " & _
					"face='trebuchet ms' color='white'><b>Phone No.</b></font></td><td align='center'>" & _
					"<font size='1' face='trebuchet ms' color='white'><b>RIHCC</b></font></td></tr>"
					
			Set rsTBL = Server.CreateObject("ADODB.RecordSet")
			
			sqlTBL = "SELECT * FROM worker_t, conwork_t, consumer_t, Proj_man_T " & _
					"WHERE PMID = Proj_Man_T.ID " & _
					"AND consumer_t.medicaid_number = CID " & _
					"AND CStr(worker_t.index) = WID " & _
					"AND status = 'Active' " & _
					"ORDER BY Proj_man_t.Lname, Proj_man_t.Fname, Worker_t.Lname, worker_t.Fname"
				
			rsTBL.Open sqlTBL, g_strCONN, 1, 3
			
			Do Until rsTBL.EOF
				Set rsLog = Server.CreateObject("ADODB.RecordSet")
				sqlLog = "SELECT * FROM W_Log_t WHERE ssn = '" & rsTBL("worker_t.Social_Security_Number") & "' "
				rsLog.Open sqlLog, g_strCONN, 1, 3
				If rsLog.EOF Then
					Set rsPM = Server.CreateObject("ADODB.RecordSet")
					sqlPM = "SELECT * FROM Proj_Man_T WHERE ID = " & rsTBL("PMID")
					rsPM.Open sqlPM, g_strCONN, 1, 3
					If Not rsPM.EOF Then
						PMname = rsPM("lname") & ", " & rsPM("fname")
					Else
						PMname = rsPM("ID")
					End If
					rsPM.Close
					Set rsPM = Nothing
					strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("Worker_t.lname") & ", " & _
								rsTBL("Worker_t.fname") & "</font></td><td align='center'><font size='1' face='trebuchet ms'>" & _
								rsTBL("Worker_t.Address") & ", " & rsTBL("Worker_t.City") & ", " & rsTBL("Worker_t.State") & ", " & rsTBL("Worker_t.zip") & "</font></td><td align='center'>" & _
								"<font size='1' face='trebuchet ms'>&nbsp;" & rsTBL("Worker_t.PhoneNo") & "</font></td><td align='center'>" & _
								"<font size='1' face='trebuchet ms'>&nbsp;" & PMname & "</font></td></tr>"
				End If
				rsLog.CLose
				Set rsLog = Nothing
				rsTBL.MoveNext
			Loop
			rsTBL.Close
			Set rsTBL = Nothing 
			
			
		ElseIf Request("SelRep") = 26 Then
			Session("MSG") = "Active Consumers with no log report. "
			
			strHEAD = "<tr bgcolor='#040C8B'><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Name</b></font></td><td align='center'>" & _
					"<font size='1' face='trebuchet ms' color='white'><b>Address</b></font></td><td align='center'><font size='1' " & _
					"face='trebuchet ms' color='white'><b>Phone No.</b></font></td><td align='center'>" & _
					"<font size='1' face='trebuchet ms' color='white'><b>RIHCC</b></font></td></tr>"
					
			Set rsTBL = Server.CreateObject("ADODB.RecordSet")
			
			sqlTBL = "SELECT * FROM Consumer_t, C_Status_T, Proj_man_T WHERE " & _
				"Consumer_t.medicaid_number = C_Status_T.medicaid_number AND PMID = Proj_man_T.ID AND Active = true " & _
				"ORDER BY Proj_man_t.Lname, Proj_man_t.Fname, Consumer_t.Lname, Consumer_t.Fname"
				
			rsTBL.Open sqlTBL, g_strCONN, 1, 3
			
			Do Until rsTBL.EOF
				Set rsLog = Server.CreateObject("ADODB.RecordSet")
				sqlLog = "SELECT * FROM C_Site_Visit_Dates_t WHERE medicaid_number = '" & rsTBL("Consumer_t.medicaid_number") & "' "
				rsLog.Open sqlLog, g_strCONN, 1, 3
				If rsLog.EOF Then
					Set rsPM = Server.CreateObject("ADODB.RecordSet")
					sqlPM = "SELECT * FROM Proj_Man_T WHERE ID = " & rsTBL("PMID")
					rsPM.Open sqlPM, g_strCONN, 1, 3
					If Not rsPM.EOF Then
						PMname = rsPM("lname") & ", " & rsPM("fname")
					Else
						PMname = rsPM("ID")
					End If
					rsPM.Close
					Set rsPM = Nothing
					strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("Consumer_t.lname") & ", " & _
								rsTBL("Consumer_t.fname") & "</font></td><td align='center'><font size='1' face='trebuchet ms'>" & _
								rsTBL("Address") & ", " & rsTBL("City") & ", " & rsTBL("State") & ", " & rsTBL("zip") & "</font></td><td align='center'>" & _
								"<font size='1' face='trebuchet ms'>&nbsp;" & rsTBL("PhoneNo") & "</font></td><td align='center'>" & _
								"<font size='1' face='trebuchet ms'>&nbsp;" & PMname & "</font></td></tr>"
				End If
				rsLog.CLose
				Set rsLog = Nothing
				rsTBL.MoveNext
			Loop
			rsTBL.Close
			Set rsTBL = Nothing 
		ElseIf Request("SelRep") = 27 Then
			'Session("MSG") = "Active Consumers with no log report. "
			selPay = ""
			selMed = ""
			if Request("seltype") = 1 then selPay = "SELECTED"
			if Request("seltype") = 2 then selMed = "SELECTED"
			PDate = Date
			PDate2 = Date
			If Request("closedate") <> "" Then 
				If IsDate(Request("closedate")) Then
					Pdate = Request("closedate")
				Else
					Session("MSG") = "Enter valid date."
					Response.Redirect "SpecRep.asp"
				End If
			End If 
			If Request("Todate") <> "" Then 
				If IsDate(Request("Todate")) Then
					Pdate2 = Request("Todate")
				Else
					Session("MSG") = "Enter valid date."
					Response.Redirect "SpecRep.asp"
				End If
			End If 
			Set rsProc = Server.CreateObject("ADODB.RecordSet")
			sqlProc = "SELECT * FROM [Tsheets_t]"
			If Request("seltype") = 1 Then
				sqlProc = sqlProc & ", worker_t  WHERE emp_id = social_security_number "
				
			ElseIf Request("seltype") = 2 Then
				
				sqlProc = sqlProc & ", consumer_t  WHERE client = medicaid_number "
			
			End If
			sqlProc = sqlProc & "AND date <= #" & Pdate2 & "# AND date >= #" & Pdate & "#"
			If Request("seltype") = 1 Then
				'sqlProc = sqlProc & " IsNull(ProcPay) ORDER BY lname, fname ASC, date DESC"
				sqlProc = sqlProc & " ORDER BY lname, fname ASC, date DESC"
			ElseIf Request("seltype") = 2 Then
				
				'sqlProc = sqlProc & " IsNull(ProcMed) AND EXT = False ORDER BY lname, fname ASC, date DESC"
				sqlProc = sqlProc & " AND EXT = False ORDER BY lname, fname ASC, date DESC"
			End If
			'response.write sqlproc
			Session("sqlVar2") = Z_DoEncrypt(Pdate & "|" & Request("seltype"))
			rsProc.Open sqlProc, g_strCONN, 1, 3
			If Not rsProc.EOF Then
				''''process
				Do Until rsProc.EOF
					
					If Request("seltype") = 1 Then
					
							tmpName = rsProc("lname") & ", " & rsProc("fname")
				
						THours = rsProc("mon") + rsProc("tue") + rsProc("wed") + rsProc("thu") + rsProc("fri") + rsProc("sat") + rsProc("sun")
						strHEAD = "<tr bgcolor='#040C8B'><td align='center'><font size='1' face='trebuchet ms' color='white'><b>SSN</b></font></td><td align='center'>" & _
						"<font size='1' face='trebuchet ms' color='white'><b>Name</b></font></td><td align='center'><font size='1' " & _
						"face='trebuchet ms' color='white'><b>Timesheet Week</b></font></td><td align='center'>" & _
						"<font size='1' face='trebuchet ms' color='white'><b>Hours</b></font></td><td align='center'>" & _
						"<font size='1' face='trebuchet ms' color='white'><b>Mileage</b></font></td><td align='center'>" & _
						"<font size='1' face='trebuchet ms' color='white'><b>Notes</b></font></td></tr>"
						maxfont = "black"
						If rsProc("MAX") = True Then maxfont = "Red"
						MC = ""
						If rsProc("milecap") = True Then MC = "**"
						If Thours <> 0 Then
							'mark = 0
							If rsProc("EXT") = True Then
								mark = 1
								strBODY = strBODY & "<tr bgcolor='#F8F8FF'><td align='center'><font size='1' face='trebuchet ms' color='" & maxfont & "'>&nbsp;" & Right(rsProc("emp_id"), 4) & "&nbsp;</font></td><td align='center'>*<font size='1' face='trebuchet ms' color='" & maxfont & "'>" & _
									tmpName & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & maxfont & "'>&nbsp;" & rsProc("date") & " - " & Cdate(rsProc("date")) + 6 & _
									"&nbsp;</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & maxfont & "'>" & Z_FormatNumber(THours,2) & _
									"</font></td><td>&nbsp;</td><td align='center'><font size='1' face='trebuchet ms' color='" & maxfont & "'>&nbsp;" & rsProc("misc_notes") & "</font></td></tr>"
							Else
								tmpMSG = ""
								If mark = 0 Then tmpMSG = rsProc("misc_notes")
								strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms' color='" & maxfont & "'>&nbsp;" & Right(rsProc("emp_id"), 4) & "&nbsp;</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & maxfont & "'>" & _
									tmpName & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & maxfont & "'>&nbsp;" & rsProc("date") & " - " & Cdate(rsProc("date")) + 6 & _
									"&nbsp;</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & maxfont & "'>" & Z_FormatNumber(THours,2) & "</font></td>" & _
									"<td align='center'><font size='1' face='trebuchet ms' color='" & maxfont & "'>" & MC & Z_FormatNumber(rsproc("mile"),2) & _
									"</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & maxfont & "'>&nbsp;" & tmpMSG & "</font></td></tr>"
								mark = 0
							End If
						End If
					Else
						
						tmpName = rsProc("lname") & ", " & rsProc("fname")
					
						THours = rsProc("mon") + rsProc("tue") + rsProc("wed") + rsProc("thu") + rsProc("fri") + rsProc("sat") + rsProc("sun")
						strHEAD = "<tr bgcolor='#040C8B'><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Medicaid</b></font></td><td align='center'>" & _
						"<font size='1' face='trebuchet ms' color='white'><b>Name</b></font></td><td align='center'><font size='1' " & _
						"face='trebuchet ms' color='white'><b>Timesheet Week</b></font></td><td align='center'>" & _
						"<font size='1' face='trebuchet ms' color='white'><b>Hours</b></font></td></tr>"
						If THours <> 0 Then
							strBODY = strBODY & "<tr>"
							If Session("lngType") = 1 Or Session("lngType") = 2 Then 
							strBODY = strBODY & "<td align='center'><font size='1' face='trebuchet ms'>&nbsp;" & rsProc("client") & "&nbsp;</font></td>"
								Else
							strBODY = strBODY & "<Td>&nbsp;</td>"
						End If 
						strBODY = strBODY & "<td align='center'><font size='1' face='trebuchet ms'>" & _
								tmpName & "</font></td><td align='center'><font size='1' face='trebuchet ms'>&nbsp;" & rsProc("date") & " - " & Cdate(rsProc("date")) + 6 & _
								"&nbsp;</font></td><td align='center'><font size='1' face='trebuchet ms'>" & Z_FormatNumber(THours,2) & "</font></td></tr>"
					  End if
						
					End If
					
					rsProc.MoveNext
				Loop
				Session("MSG") = "Records from " & Pdate & " to " & Pdate2 & " for "
				If Request("seltype") = 1 Then
						Session("MSG") = Session("MSG") & "payroll report. <br>* extended hours<br>** over mileage cap<br>red font - over max hours"
				Else
						Session("MSG") = Session("MSG") & "medicaid report."
				End If 
			
			Else
				Session("MSG") = "No records found."
				
			End If
			rsProc.Close
			Set rsProc = Nothing
			ElseIf Request("selRep") = 28 Then
					Session("MSG") = "Active PCSP Worker mailing list. "
							Set rsProc = Server.CreateObject("ADODB.RecordSet")
				sqlProc = "SELECT * FROM Worker_T WHERE Status = 'Active' ORDER BY Lname, Fname"
				rsProc.Open sqlProc, g_strCONN, 3, 1
				If Not rsProc.EOF Then
					strHEAD = "<tr bgcolor='#040C8B'><td align='center'>" & _
							"<font size='1' face='trebuchet ms' color='white'><b>Name</b></font></td><td align='center'><font size='1' " & _
							"face='trebuchet ms' color='white'><b>Mailing Address</b></font></td></tr>"
					Do Until rsProc.EOF
						strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & rsProc("lname") & ", " & _
										rsProc("fname") & "</font></td><td align='center'><font size='1' face='trebuchet ms'>" & _
										rsProc("mAddress") & ", " & rsProc("mCity") & ", " & rsProc("mState") & ", " & rsProc("mzip") & "</font></td></tr>" & vbCrLf
						rsPRoc.MoveNext
					Loop
				Else
					Session("MSG") = "No records found."
				End If
				rsProc.Close
				Set rsProc = Nothing 	
			ElseIf Request("SelRep") = 29 Then 
				Session("MSG") = "Active Consumer Start and Ammendment Expiration Date report"
				strHEAD = "<tr bgcolor='#040C8B'><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Consumer Name</b>" & _
						"</font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Start Date</b></font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Ammendment Expiration Date</b></font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>RIHCC</b></font></td></tr>"
				Set rsTBL = Server.CreateObject("ADODB.RecordSet")
				sqlTBL = "SELECT * FROM consumer_t, c_status_t WHERE consumer_t.medicaid_number = c_status_t.medicaid_number AND Active = true"
				If Request("FrmD8") <> "" Then
					If IsDate(Request("FrmD8")) Then
						sqlTBL = sqlTBL & " AND End_Date >= #" & Request("FrmD8") & "# " 
						Session("Msg") = Session("Msg") & " from " & Request("FrmD8")
					End If
				End If
				If Request("ToD8") <> "" Then
					If IsDate(Request("ToD8")) Then
						sqlTBL = sqlTBL & " AND End_Date <= #" & Request("ToD8") & "# " 
						Session("Msg") = Session("Msg") & " to " & Request("ToD8")
					End If
				End If
				sqlTBL  = sqlTBL  & " ORDER BY lname, fname"
				Session("Msg") = Session("Msg") & ". " 
				'response.write sqlTBL
				rsTBL.Open sqlTBL, g_strCONN, 1, 3
				Do Until rsTBL.EOF	
					strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & _
						rsTBL("lname") & ", " & rsTBL("fname") &"</font></td><td align='center'><font size='1' face='trebuchet ms'>&nbsp;" & rsTBL("Start_date") & _
						"</font></td><td align='center'><font size='1' face='trebuchet ms'>&nbsp;" & _
						rsTBL("End_Date") &  "</font></td><td align='center'><font size='1' face='trebuchet ms'>&nbsp;" & _
						GetPM(rsTBL("PMID")) &  "</font></td></tr>"
					rsTBL.MoveNext
				Loop
				rsTBL.Close
				Set rsTBL = Nothing	
			ElseIf Request("SelRep") = 30 Then 
				Session("MSG") = "All Active PCSP Worker Extended Hours report"
				Set rsTBL = Server.CreateObject("ADODB.RecordSet")
				sqlTBL = "SELECT * FROM tsheets_T, worker_T, consumer_T WHERE EXT = true AND emp_id = worker_T.social_security_number AND " & _
					"consumer_T.Medicaid_number = client ORDER BY worker_T.lname, worker_T.fname ASC, date DESC"
				rsTBL.Open sqlTBL, g_strCONN, 3, 1
				strHEAD = "<tr bgcolor='#040C8B'><td align='center'><font size='1' face='trebuchet ms' color='white'><b>SSN</b></font></td><td align='center'>" & _
						"<font size='1' face='trebuchet ms' color='white'><b>Name</b></font></td><td align='center'><font size='1' " & _
						"face='trebuchet ms' color='white'><b>Timesheet Week</b></font></td><td align='center'>" & _
						"<font size='1' face='trebuchet ms' color='white'><b>Hours</b></font></td><td align='center'>" & _
						"<font size='1' face='trebuchet ms' color='white'><b>Consumer</b></font></td><td align='center'>" & _
						"<font size='1' face='trebuchet ms' color='white'><b>RIHCC</b></font></td><td align='center'>" & _
						"<font size='1' face='trebuchet ms' color='white'><b>RIHCC2</b></font></td><td align='center'>" & _
						"<font size='1' face='trebuchet ms' color='white'><b>Notes</b></font></td></tr>"
				Do Until rsTBL.EOF
					THours = rsTBL("mon") + rsTBL("tue") + rsTBL("wed") + rsTBL("thu") + rsTBL("fri") + rsTBL("sat") + rsTBL("sun")
					tmpName = rsTBL("worker_T.lname") & ", " & rsTBL("worker_T.fname")
					tmpCName = rsTBL("consumer_T.lname") & ", " & rsTBL("consumer_T.fname")
					If Thours <> 0 Then
						strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>&nbsp;" & Right(rsTBL("emp_id"), 4) & "&nbsp;</font></td><td align='center'><font size='1' face='trebuchet ms'>" & _
									tmpName & "</font></td><td align='center'><font size='1' face='trebuchet ms'>&nbsp;" & rsTBL("date") & " - " & Cdate(rsTBL("date")) + 6 & _
									"&nbsp;</font></td><td align='center'><font size='1' face='trebuchet ms'>" & Z_FormatNumber(THours,2) & _
									"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & tmpCName & _
									"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & GetName3(rsTBL("PM1")) & _
									"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & GetName3(rsTBL("PM2")) & _
									"</font></td><td align='center'><textarea rows='1' readonly>" & rsTBL("misc_notes") & "</textarea></td></tr>"
					End If
					rsTBL.MoveNext
				Loop
				rsTBL.Close
				Set rsTBL = Nothing
			ElseIf Request("SelRep") = 31 Then 
				If Request("seltype2") = 1 Then
					Session("MSG") = "Payroll report"
					strHEAD = "<tr bgcolor='#040C8B'><td align='center'><font size='1' face='trebuchet ms' color='white'><b>SSN</b></font></td><td align='center'>" & _
						"<font size='1' face='trebuchet ms' color='white'><b>Name</b></font></td><td align='center'><font size='1' " & _
						"face='trebuchet ms' color='white'><b>Timesheet Week</b></font></td><td align='center'>" & _
						"<font size='1' face='trebuchet ms' color='white'><b>Hours</b></font></td><td align='center'>" & _
						"<font size='1' face='trebuchet ms' color='white'><b>Notes</b></font></td></tr>"
				ElseIf Request("seltype2") = 2 Then
					Session("MSG") = "Medicaid report"
					strHEAD = "<tr bgcolor='#040C8B'><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Medicaid</b></font></td><td align='center'>" & _
						"<font size='1' face='trebuchet ms' color='white'><b>Name</b></font></td><td align='center'><font size='1' " & _
						"face='trebuchet ms' color='white'><b>Timesheet Week</b></font></td><td align='center'>" & _
						"<font size='1' face='trebuchet ms' color='white'><b>Hours</b></font></td><td align='center'>" & _
						"<font size='1' face='trebuchet ms' color='white'><b>Code</b></font></td><td align='center'>" & _
						"<font size='1' face='trebuchet ms' color='white'><b>Notes</b></font></td></tr>"
				ElseIf Request("seltype2") = 3 Then
					Session("MSG") = "Private Pay/Contract/Admin report"
					strHEAD = "<tr bgcolor='#040C8B'><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Medicaid</b></font></td><td align='center'>" & _
						"<font size='1' face='trebuchet ms' color='white'><b>Name</b></font></td><td align='center'><font size='1' " & _
						"face='trebuchet ms' color='white'><b>Timesheet Week</b></font></td><td align='center'>" & _
						"<font size='1' face='trebuchet ms' color='white'><b>Hours</b></font></td><td align='center'>" & _
						"<font size='1' face='trebuchet ms' color='white'><b>Rate</b></font></td><td align='center'>" & _
						"<font size='1' face='trebuchet ms' color='white'><b>Code</b></font></td><td align='center'>" & _
						"<font size='1' face='trebuchet ms' color='white'><b>Notes</b></font></td></tr>"
				End IF
				Set rsTBL = Server.CreateObject("ADODB.RecordSet")
				sqlTBL = "SELECT * FROM [Tsheets_t]"
				If Request("seltype2") = 1 Then
					sqlTBL = sqlTBL & ", worker_t  WHERE emp_id = social_security_number "
				ElseIf Request("seltype2") = 2 Then
					sqlTBL = sqlTBL & ", consumer_t  WHERE client = medicaid_number and code = 'M' "
				ElseIf Request("seltype2") = 3 Then
					sqlTBL = sqlTBL & ", consumer_t  WHERE client = medicaid_number and code <> 'M' "
				End If
				Err = 0
				If Request("FrmD8") <> "" Then
					If IsDate(Request("FrmD8")) Then
						'If (Month(Request("FrmD8")) - 1) <> 0 Then 
							'sqlTBL = sqlTBL & " AND Month(date) >= " & Month(Request("FrmD8")) - 1 & " " 
							sqlTBL = sqlTBL & " AND date >= #" & CDate(Request("FrmD8")) - 6 & "#" 
						'Else
						'	tmpYear = Year(Request("FrmD8")) - 1
						'	sqlTBL = sqlTBL & " AND date >= #12/1/" & tmpYear & "#"
						'End If
						Session("Msg") = Session("Msg") & " from " & Request("FrmD8")
					Else
						Err = 1
					End If
				End If
				If Request("ToD8") <> "" Then
					If IsDate(Request("ToD8")) Then
						'If Month(Request("ToD8")) <> 1 Then
						'	sqlTBL = sqlTBL & " AND Month(date) - 1 <= " & Month(Request("ToD8")) & " " 
						'Else
							sqlTBL = sqlTBL & " AND date  <= #" & CDate(Request("ToD8")) + 6 & "#" 
						'End If
						Session("Msg") = Session("Msg") & " to " & Request("ToD8")
					Else
						Err = 1
					End If
				End If
				If Request("seltype2") = 1 Then
					sqlTBL = sqlTBL & " ORDER BY lname ASC, fname ASC, date DESC"
				ElseIf Request("seltype2") = 2 or Request("seltype2") = 3 Then
					sqlTBL = sqlTBL & " AND EXT = False ORDER BY lname ASC, fname ASC, date DESC"
				End If
				Session("Msg") = Session("Msg") & ". " 
				If Request("seltype2") = 1 Then Session("Msg") = Session("Msg") & "<br>* - Extended hours"
				If Err <> 0 Then Response.Redirect "specrep.asp?err=31" 
				'response.write sqlTBL
				rsTBL.Open sqlTBL, g_strCONN, 1, 3
				Do Until rsTBL.EOF	
					tmpThrs = rsTBL("mon") + rsTBL("tue") + rsTBL("wed") + rsTBL("thu") + rsTBL("fri") + rsTBL("sat") + rsTBL("sun") 
					If tmpThrs <> 0 Then
						tmphrsMon = ValidDate(Request("FrmD8"), Request("ToD8"), rsTBL("date"), rsTBL("mon"), "MON")
            tmphrsTue = ValidDate(Request("FrmD8"), Request("ToD8"), rsTBL("date"), rsTBL("tue"), "TUE")
            tmphrsWed = ValidDate(Request("FrmD8"), Request("ToD8"), rsTBL("date"), rsTBL("wed"), "WED")
            tmphrsThu = ValidDate(Request("FrmD8"), Request("ToD8"), rsTBL("date"), rsTBL("thu"), "THU")
            tmphrsFri = ValidDate(Request("FrmD8"), Request("ToD8"), rsTBL("date"), rsTBL("fri"), "FRI")
            tmphrsSat = ValidDate(Request("FrmD8"), Request("ToD8"), rsTBL("date"), rsTBL("sat"), "SAT")
            tmphrsSun = ValidDate(Request("FrmD8"), Request("ToD8"), rsTBL("date"), rsTBL("sun"), "SUN")
            FtotHrs = tmphrsMon + tmphrsTue + tmphrsWed + tmphrsThu + tmphrsFri + tmphrsSat + tmphrsSun
						If FtotHrs <> 0 Then
							tmpTSWk1 = rsTBL("date")
							If Request("FrmD8") <> "" Then
								If Cdate(tmpTSWk1) < Cdate(Request("FrmD8")) Then tmpTSWk1 = Request("FrmD8")
							End If
							tmpTSWk2 = Cdate(rsTBL("date")) + 6
							If Request("ToD8") <> "" Then
								If Cdate(tmpTSWk2) > Cdate(Request("ToD8")) Then tmpTSWk2 = Request("ToD8")
							End If
							If Request("seltype2") = 2 or Request("seltype2") = 3 Then
								strBODY = strBODY & "<tr>"
									If Session("lngType") = 1 Or Session("lngType") = 2 Then 
										strBODY = strBODY & "<td align='center'><font size='1' face='trebuchet ms' >&nbsp;" & rsTBL("client") & "&nbsp;</font></td>"
												Else
									strBODY = strBODY & "<Td>&nbsp;</td>"
								End If 	
						strBODY = strBODY & "<td align='center'><font size='1' face='trebuchet ms' color='" & maxfont & "'>" & _
											GetName2(rsTBL("client")) & "</font></td><td align='center'><font size='1' face='trebuchet ms' >&nbsp;" & tmpTSWk1 & " - " & tmpTSWk2 & _
											"&nbsp;</font></td><td align='center'><font size='1' face='trebuchet ms' >" & Z_FormatNumber(FtotHrs,2) & "</font></td>"
								If Request("seltype2") = 3 Then	
											strBODY = strBODY & "<td align='center'><font size='1' face='trebuchet ms' >" & z_formatnumber(rsTBL("rate"),2) & _
											"</font></td>"
								End If
								strBODY = strBODY &	"<td align='center'><font size='1' face='trebuchet ms' >" & rsTBL("code") & _
											"</font></td><td align='center'><font size='1' face='trebuchet ms' >&nbsp;" & rsTBL("misc_notes") & "</font></td></tr>"
							ElseIf Request("seltype2") = 1 Then
								tmpEXT = ""
								If rsTBL("EXT") = True Then tmpEXT = "*"
								strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms' >&nbsp;" & Right(rsTBL("emp_id"), 4) & "&nbsp;</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & maxfont & "'><b>" & _
											tmpEXT & "</b>" & GetName(rsTBL("emp_id")) & "</font></td><td align='center'><font size='1' face='trebuchet ms' >&nbsp;" & tmpTSWk1 & " - " & tmpTSWk2 & _
											"&nbsp;</font></td><td align='center'><font size='1' face='trebuchet ms' >" & Z_FormatNumber(FtotHrs,2) & _
											"</font></td><td align='center'><font size='1' face='trebuchet ms' >&nbsp;" & rsTBL("misc_notes") & "</font></td></tr>"
							End If
						End If
					End If
					rsTBL.MoveNext
				Loop
				rsTBL.Close
				Set rsTBL = Nothing	
			ElseIf Request("SelRep") = 32 Then 
				Session("MSG") = "Extended hours report"
					strHEAD = "<tr bgcolor='#040C8B'><td align='center'><font size='1' face='trebuchet ms' color='white'><b>SSN</b></font></td><td align='center'>" & _
						"<font size='1' face='trebuchet ms' color='white'><b>PCSP Worker</b></font></td><td align='center'>" & _
						"<font size='1' face='trebuchet ms' color='white'><b>Consumer</b></font></td><td align='center'>" & _
						"<font size='1' face='trebuchet ms' color='white'><b>RIHCC</b></font></td><td align='center'><font size='1' " & _
						"face='trebuchet ms' color='white'><b>Timesheet Week</b></font></td><td align='center'>" & _
						"<font size='1' face='trebuchet ms' color='white'><b>Hours</b></font></td><td align='center'>" & _
						"<font size='1' face='trebuchet ms' color='white'><b>Notes</b></font></td></tr>"
				Set rsTBL = Server.CreateObject("ADODB.RecordSet")
				sqlTBL = "SELECT * FROM tsheets_T, worker_T, consumer_T WHERE emp_id = worker_T.Social_Security_Number AND ext = true" & _
					" AND client = Medicaid_Number" 
				Err = 0
				If Request("FrmD8") <> "" Then
					If IsDate(Request("FrmD8")) Then
						'If (Month(Request("FrmD8")) - 1) <> 0 Then 
						'	sqlTBL = sqlTBL & " AND Month(date) >= " & Month(Request("FrmD8")) - 1 & " " 
							sqlTBL = sqlTBL & " AND date >= #" & CDate(Request("FrmD8")) - 6 & "#"
						'Else
						'	tmpYear = Year(Request("FrmD8")) - 1
						'	sqlTBL = sqlTBL & " AND date >= #12/1/" & tmpYear & "#"
						'End If
						Session("Msg") = Session("Msg") & " from " & Request("FrmD8")
					Else
						Err = 1
					End If
				End If
				If Request("ToD8") <> "" Then
					If IsDate(Request("ToD8")) Then
						'If Month(Request("ToD8")) <> 1 Then
						'	sqlTBL = sqlTBL & " AND Month(date) - 1 <= " & Month(Request("ToD8")) & " " 
						'Else
							sqlTBL = sqlTBL & " AND date  <= #" & CDate(Request("ToD8")) + 6 & "#" 
						'End If
						Session("Msg") = Session("Msg") & " to " & Request("ToD8")
					Else
						Err = 1
					End If
				End If
				sqlTBL  = sqlTBL  & " ORDER BY worker_T.lname, worker_T.fname, date, timestamp"
				Session("Msg") = Session("Msg") & ". " 
				If Err <> 0 Then Response.Redirect "specrep.asp?err=32" 
				rsTBL.Open sqlTBL, g_strCONN, 1, 3
				Do Until rsTBL.EOF	
					tmpThrs = rsTBL("mon") + rsTBL("tue") + rsTBL("wed") + rsTBL("thu") + rsTBL("fri") + rsTBL("sat") + rsTBL("sun") 
					If tmpThrs <> 0 Then
						tmphrsMon = ValidDate(Request("FrmD8"), Request("ToD8"), rsTBL("date"), rsTBL("mon"), "MON")
			         		tmphrsTue = ValidDate(Request("FrmD8"), Request("ToD8"), rsTBL("date"), rsTBL("tue"), "TUE")
			            	tmphrsWed = ValidDate(Request("FrmD8"), Request("ToD8"), rsTBL("date"), rsTBL("wed"), "WED")
			            	tmphrsThu = ValidDate(Request("FrmD8"), Request("ToD8"), rsTBL("date"), rsTBL("thu"), "THU")
			            	tmphrsFri = ValidDate(Request("FrmD8"), Request("ToD8"), rsTBL("date"), rsTBL("fri"), "FRI")
			            	tmphrsSat = ValidDate(Request("FrmD8"), Request("ToD8"), rsTBL("date"), rsTBL("sat"), "SAT")
			            	tmphrsSun = ValidDate(Request("FrmD8"), Request("ToD8"), rsTBL("date"), rsTBL("sun"), "SUN")
			            	FtotHrs = tmphrsMon + tmphrsTue + tmphrsWed + tmphrsThu + tmphrsFri + tmphrsSat + tmphrsSun
						If FtotHrs <> 0 Then
						tmpTSWk1 = rsTBL("date")
						If Request("FrmD8") <> "" Then
							If Cdate(tmpTSWk1) < Cdate(Request("FrmD8")) Then tmpTSWk1 = Request("FrmD8")
						End If
						tmpTSWk2 = Cdate(rsTBL("date")) + 6
						If Request("ToD8") <> "" Then
							If Cdate(tmpTSWk2) > Cdate(Request("ToD8")) Then tmpTSWk2 = Request("ToD8")
						End If
							strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms' >&nbsp;" & Right(rsTBL("emp_id"), 4) & "&nbsp;</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & maxfont & "'>" & _
										GetName(rsTBL("emp_id")) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & maxfont & "'>" & _
										GetName2(rsTBL("client")) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & maxfont & "'>" & _
										GetName3(rsTBL("PMID")) & "</font></td><td align='center'><font size='1' face='trebuchet ms' >&nbsp;" & tmpTSWK1 & " - " & tmpTSWK2 & _
										"&nbsp;</font></td><td align='center'><font size='1' face='trebuchet ms' >" & Z_FormatNumber(FtotHrs,2) & _
										"</font></td><td align='center'><font size='1' face='trebuchet ms' >&nbsp;" & rsTBL("misc_notes") & "</font></td></tr>"
						End If
					End If
					rsTBL.MoveNext
				Loop
				rsTBL.Close
				Set rsTBL = Nothing	
			ElseIf Request("SelRep") = 33 Then
				Session("MSG") = "PCSP Worker by Driver's License Expiration Date report."
				strHEAD = "<tr bgcolor='#040C8B'><td align='center'><font size='1' face='trebuchet ms' color='white'><b>PCSP Worker Name</b>" & _
						"</font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Address</b></font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Date of Birth</b></font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>License Number</b></font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Expiration Date</b></font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>RIHCC</b></font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>RIHCC2</b></font></td></tr>"
				Set rsTBL = Server.CreateObject("ADODB.RecordSet")
				sqlTBL = "SELECT * FROM worker_t WHERE status = 'Active' AND Driver = True ORDER BY LicenseExpDate, worker_t.lname, worker_t.fname"
				rsTBL.Open sqlTBL, g_strCONN, 1, 3
				Do Until rsTBL.EOF	
					strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("lname") & ", " & _
						rsTBL("fname") & "</font></td><td align='center'><font size='1' face='trebuchet ms'>" & _
						rsTBL("Address") & ", " & rsTBL("City") & ", " & rsTBL("State") & ", " & rsTBL("Zip") & "</font></td><td align='center'>" & _
						"<font size='1' face='trebuchet ms'>&nbsp;" & rsTBL("DOB") & "</td><td align='center'>" & _
						"<font size='1' face='trebuchet ms'>&nbsp;" & rsTBL("LicenseNo") & "</td><td align='center'>" & _
						"<font size='1' face='trebuchet ms'>&nbsp;" & rsTBL("LicenseExpDate") & "</td><td align='center'>" & _
						"<font size='1' face='trebuchet ms'>&nbsp;" & GetName3(rsTBL("PM1")) & "</td><td align='center'>" & _
						"<font size='1' face='trebuchet ms'>&nbsp;" & GetName3(rsTBL("PM2")) & "</td></tr>"
					rsTBL.MoveNext
				Loop
				rsTBL.Close 
			ElseIf Request("SelRep") = 34 Then
				Session("MSG") = "PCSP Worker with Active Consumers report."
				typ = Request("SelRep")
				strHEAD = "<tr bgcolor='#040C8B'><td align='center'><font size='1' face='trebuchet ms' color='white'><b>PCSP Worker Name</b>" & _
						"</font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Mailing Address</b></font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Date Of Hire</b></font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Consumer Name</b></font></td></tr>"
				Set rsTBL = Server.CreateObject("ADODB.RecordSet")
				sqlTBL = "SELECT * " & _
					"FROM worker_t, consumer_T, ConWork_T, C_Status_t " & _
					"WHERE worker_T.Status = 'Active' " & _
					"AND worker_T.index = Cint(WID) " & _
					"AND CID = Consumer_t.medicaid_Number " & _ 
					"AND Consumer_t.Medicaid_number = C_Status_t.Medicaid_number " & _
					"AND Active = True " & _
					"ORDER BY Month(Date_Hired), Worker_T.Lname, Worker_T.Fname, Consumer_T.Lname, Consumer_T.Fname"
				rsTBL.Open sqlTBL, g_strCONN, 1, 3
				Do Until rsTBL.EOF	
					strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("Worker_T.lname") & ", " & _
						rsTBL("Worker_T.fname") & "</font></td><td align='center'><font size='1' face='trebuchet ms'>" & _
						rsTBL("Worker_T.mAddress") & ", " & rsTBL("Worker_T.mCity") & ", " & rsTBL("Worker_T.mState") & ", " & rsTBL("Worker_T.mZip") & "</font>" & _
						"</td><td align='center'><font size='1' face='trebuchet ms'>&nbsp;" & rsTBL("Date_Hired") & "</td>" & _
						"<td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("Consumer_T.lname") & ", " & _
						rsTBL("Consumer_T.fname") & "</font></td></tr>"
					rsTBL.MoveNext
				Loop
				rsTBL.Close 
			ElseIf Request("SelRep") = 35 Then
				Session("MSG") = "Representative report."
				strHEAD = "<tr bgcolor='#040C8B'><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Representative Name</b>" & _
						"</font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Address</b></font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Consumer Name</b></font></td></tr>"
				Set rsTBL = Server.CreateObject("ADODB.RecordSet")
				sqlTBL = "SELECT * " & _
					"FROM Representative_T, consumer_T, Conrep_T, C_Status_t " & _
					"WHERE Cint(RID) = Representative_T.index " & _
					"AND CID = Consumer_t.medicaid_Number " & _ 
					"AND C_Status_t.Medicaid_number = Consumer_t.Medicaid_number AND Active = True " & _
					"ORDER BY Representative_T.Lname, Representative_T.Fname"
				rsTBL.Open sqlTBL, g_strCONN, 1, 3
				Do Until rsTBL.EOF	
					strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("Representative_T.lname") & ", " & _
						rsTBL("Representative_T.fname") & "</font></td><td align='center'><font size='1' face='trebuchet ms'>" & _
						rsTBL("Representative_T.Address") & ", " & rsTBL("Representative_T.City") & ", " & rsTBL("Representative_T.State") & ", " & rsTBL("Representative_T.Zip") & "</font>" & _
						"</td><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("Consumer_T.lname") & ", " & _
						rsTBL("Consumer_T.fname") & "</font></td></tr>"
					rsTBL.MoveNext
				Loop
				rsTBL.Close 
			ElseIf Request("SelRep") = 36 Then
				Session("MSG") = "Case Manager report."
				strHEAD = "<tr bgcolor='#040C8B'><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Case Manager Name</b>" & _
						"</font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Address</b></font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Agency</b></font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Consumer Name</b></font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>RIHCC</b></font></td></tr>"
				Set rsTBL = Server.CreateObject("ADODB.RecordSet")
				sqlTBL = "SELECT * " & _
					"FROM Case_manager_T, consumer_T, CMCon_T " & _
					"WHERE Cint(CMID) = Case_manager_T.index " & _
					"AND CID = Consumer_t.medicaid_Number " & _ 
					"ORDER BY Case_manager_T.Lname, Case_manager_T.Fname, consumer_T.lname, consumer_t.fname"
				rsTBL.Open sqlTBL, g_strCONN, 1, 3
				Do Until rsTBL.EOF	
					strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("Case_manager_T.lname") & ", " & _
						rsTBL("Case_manager_T.fname") & "</font></td><td align='center'><font size='1' face='trebuchet ms'>" & _
						rsTBL("Case_manager_T.Address") & ", " & rsTBL("Case_manager_T.City") & ", " & rsTBL("Case_manager_T.State") & ", " & rsTBL("Case_manager_T.Zip") & "</font>" & _
						"</td><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("agency") & "</font></td><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("Consumer_T.lname") & ", " & _
						rsTBL("Consumer_T.fname") & "</font></td><td align='center'><font size='1' face='trebuchet ms'>" & GetName3(rsTBL("PMID"))& "</font></td></tr>"
					rsTBL.MoveNext
				Loop
				rsTBL.Close 
			ElseIf Request("SelRep") = 37 Then
				Session("MSG") = "Inactive Consumers report."
				strHEAD = "<tr bgcolor='#040C8B'><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Consumer Name</b>" & _
						"</font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Inactive Date</b></font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Reason</b></font></td></tr>"
				Set rsTBL = Server.CreateObject("ADODB.RecordSet")
				sqlTBL = "SELECT * " & _
					"FROM consumer_T, C_Status_T " & _
					"WHERE consumer_T.medicaid_number = C_Status_T.medicaid_number " & _
					"AND Active = False " & _ 
					"ORDER BY Lname, Fname"
				rsTBL.Open sqlTBL, g_strCONN, 1, 3
				Do Until rsTBL.EOF	
					tmpReason = "&nbsp;"
					If rsTBL("Enter_Nursing_Home") = True Then tmpReason = "Entered Nursing Home or Other Setting, "
					If rsTBL("Unable_Self_Direct") = True Then tmpReason = tmpReason & "Unable to Self-Direct, "	
					If rsTBL("Unable_Suitable_Worker") = True Then tmpReason = tmpReason & "Unable to find Suitable Worker, "
					If rsTBL("Death") = True Then tmpReason = tmpReason & "Death, "
					If rsTBL("A_Other") <> "" Then tmpReason = tmpReason & rsTBL("A_Other")
					strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("lname") & ", " & _
						rsTBL("fname") & "</font></td><td align='center'><font size='1' face='trebuchet ms'>" & _
						rsTBL("Inactive_Date") & "&nbsp;</font>" & _
						"</td><td align='center'><font size='1' face='trebuchet ms'>" & tmpReason & "</font></td></tr>"
					rsTBL.MoveNext
				Loop
				rsTBL.Close 
			ElseIf Request("SelRep") = 38 Then
				Session("MSG") = "Consumers Current Care Plan report."
				strHEAD = "<tr bgcolor='#040C8B'><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Consumer Name</b>" & _
						"</font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Current Care Plan</b></font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>RIHCC</b></font></td></tr>"
				Set rsTBL = Server.CreateObject("ADODB.RecordSet")
				sqlTBL = "SELECT * " & _
					"FROM consumer_T, C_Status_T " & _
					"WHERE consumer_T.medicaid_number = C_Status_T.medicaid_number " & _
					"AND Active = True " & _ 
					"ORDER BY CarePlan,Lname, Fname"
				rsTBL.Open sqlTBL, g_strCONN, 1, 3
				Do Until rsTBL.EOF	
					strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("lname") & ", " & _
						rsTBL("fname") & "</font></td><td align='center'><font size='1' face='trebuchet ms'>" & _
						rsTBL("CarePlan") & "&nbsp;</font></td><td align='center'><font size='1' face='trebuchet ms'>" & _
						GetPM(rsTBL("PMID")) & "&nbsp;</font></td></tr>"
					rsTBL.MoveNext
				Loop
				rsTBL.Close 
			ElseIf Request("SelRep") = 39 Then
			PDate = Date
			PDate2 = Date
			If Request("closedate") <> "" Then 
				If IsDate(Request("closedate")) Then
					Pdate = Request("closedate")
				Else
					Session("MSG") = "Enter valid date."
					Response.Redirect "SpecRep.asp"
				End If
			Else
					Session("MSG") = "Date cannot be blank."
					Response.redirect "SpecRep.asp"
			End If 
			If Request("Todate") <> "" Then 
				If IsDate(Request("Todate")) Then
					Pdate2 = Request("Todate")
				Else
					Session("MSG") = "Enter valid date."
					Response.Redirect "SpecRep.asp"
				End If
			End If 
				Session("MSG") = "PCSP worker with unsubmitted timesheets (" & Pdate & " - " & Pdate2 & ") report."
					strHEAD = "<tr bgcolor='#040C8B'><td align='center'><font size='1' face='trebuchet ms' color='white'><b>PCSP Worker</b>" & _
						"</font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Timesheet Week</b></font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Consumer</b></font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>RIHCC</b></font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>RIHCC2</b></font></td></tr>"
				Set rsTBL = Server.CreateObject("ADODB.RecordSet")
				sqlTBL = "SELECT * " & _
					"FROM worker_T, consumer_T, ConWork_T, Proj_man_T, C_status_T " & _
					"WHERE worker_T.Status = 'Active' " & _
					"AND PM1 = Proj_man_T.ID " & _
					"AND worker_T.index = Cint(WID) " & _
					"AND CID = Consumer_t.medicaid_Number " & _
					"AND CID = C_status_T.medicaid_Number " & _
					"AND onHold = false " & _
					"ORDER BY Proj_man_t.Lname, Proj_man_t.Fname, worker_T.Lname, worker_T.Fname "
				rsTBL.Open sqlTBL, g_strCONN, 1, 3
				Do Until rsTBL.EOF	
					Set rsChkMe = Server.CreateObject("ADODB.RecordSet")
					sqlChkMe = "SELECT * FROM Tsheets_T WHERE emp_id = '" & rsTBL("worker_T.Social_security_number") & "' AND [Date] = #" & _
						Pdate & "# AND client = '" & rsTBL("consumer_T.Medicaid_number") & "'" 'AND [Date] >= #" & Pdate2 & "# "
					rsChkMe.Open sqlChkMe, g_strCONN, 3, 1
					If rsChkMe.EOF Then
						strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("worker_T.lname") & ", " & _
							rsTBL("worker_T.fname") & "</font></td><td align='center'><font size='1' face='trebuchet ms'>" & _
							Pdate & " - " & Cdate(Pdate) + 6 & "&nbsp;</font></td><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("consumer_T.lname") & ", " & _
							rsTBL("consumer_T.fname") & "</font></td><td align='center'><font size='1' face='trebuchet ms'>" & _
							GetName3(rsTBL("pm1")) & "</font></td><td align='center'><font size='1' face='trebuchet ms'>" & _
							GetName3(rsTBL("pm2")) & "</font></td></tr>"
					End If
					rsChkMe.Close
					Set rsChkMe = Nothing
					Set rsChkMe2 = Server.CreateObject("ADODB.RecordSet")
					sqlChkMe2 = "SELECT * FROM Tsheets_T WHERE emp_id = '" & rsTBL("worker_T.Social_security_number") & "' AND [Date] = #" & _
						Cdate(Pdate2) - 6 & "# AND client = '" & rsTBL("consumer_T.Medicaid_number") & "'" 'AND [Date] >= #" & Pdate2 & "# "
					rsChkMe2.Open sqlChkMe2, g_strCONN, 3, 1
					If rsChkMe2.EOF Then
						strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("worker_T.lname") & ", " & _
							rsTBL("worker_T.fname") & "</font></td><td align='center'><font size='1' face='trebuchet ms'>" & _
							Cdate(Pdate2) - 6 & " - " & Pdate2 & "&nbsp;</font></td><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("consumer_T.lname") & ", " & _
							rsTBL("consumer_T.fname") & "</font></td><td align='center'><font size='1' face='trebuchet ms'>" & _
							GetName3(rsTBL("pm1")) & "</font></td><td align='center'><font size='1' face='trebuchet ms'>" & _
							GetName3(rsTBL("pm2")) & "</font></td></tr>"
					End If
					rsChkMe2.Close
					Set rsChkMe2 = Nothing
					rsTBL.MoveNext
				Loop
				rsTBL.Close 
				Set rstbl = NOthing
			ElseIf Request("SelRep") = 40 Then
				Session("MSG") = "Consumers Start and Inactive Date report."
				strHEAD = "<tr bgcolor='#040C8B'><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Consumer Name</b>" & _
						"</font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>RIHCC</b>" & _
						"</font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Start Date</b></font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Inactive Date</b></font></td></tr>"
				Set rsTBL = Server.CreateObject("ADODB.RecordSet")
				sqlTBL = "SELECT * " & _
					"FROM consumer_T, C_Status_T " & _
					"WHERE consumer_T.medicaid_number = C_Status_T.medicaid_number " & _
					"ORDER BY Inactive_Date, Lname, Fname"
				rsTBL.Open sqlTBL, g_strCONN, 1, 3
				Do Until rsTBL.EOF	
					strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("lname") & ", " & _
						rsTBL("fname") & "</font></td><td align='center'><font size='1' face='trebuchet ms'>" & _
						Getname3(rsTBL("PMID")) & "&nbsp;</font></td><td align='center'><font size='1' face='trebuchet ms'>" & _
						rsTBL("Start_Date") & "&nbsp;</font></td><td align='center'><font size='1' face='trebuchet ms'>" & _
						rsTBL("Inactive_Date") & "&nbsp;</font></td></tr>"
					rsTBL.MoveNext
				Loop
				rsTBL.Close 
			ElseIf Request("SelRep") = 41 Then
				Session("MSG") = "Consumers logs"
				strHEAD = "<tr bgcolor='#040C8B'><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Consumer Name</b></font></td>"
						If Request("SelLog") = 1 Then
							strHEAD = strHEAD & "<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Site Visit Date</b></font></td>"
							strHEAD = strHEAD & "<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Comment</b></font></td></tr>"
						ElseIf Request("SelLog") = 2 Then
							strHEAD = strHEAD & "<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Phone Call Date</b></font></td>"
							strHEAD = strHEAD & "<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Comment</b></font></td></tr>"
						ElseIf Request("SelLog") = 3 Then
							strHEAD = strHEAD & "<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Misc. Contact Date</b></font></td>"
							strHEAD = strHEAD & "<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Comment</b></font></td></tr>"
						Else
							strHEAD = strHEAD & "<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Site Visit Date</b></font></td>"
							strHEAD = strHEAD & "<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Phone Call Date</b></font></td>"
							strHEAD = strHEAD & "<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Misc. Contact Date</b></font></td>"
							strHEAD = strHEAD & "<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Comment</b></font></td></tr>"
						End If
						
				Set rsTBL = Server.CreateObject("ADODB.RecordSet")
				sqlTBL = "SELECT * FROM Consumer_T, C_Site_Visit_Dates_T, C_Status_T WHERE Consumer_T.Medicaid_Number = C_Site_Visit_Dates_T.Medicaid_Number" & _
					" AND consumer_T.Medicaid_Number = C_Status_T.Medicaid_Number AND Active = true"
				If Request("SelLog") = 1 Then
					sqlTBL = sqlTBL & " AND NOT IsNull(Site_V_Date)"
					Session("Msg") = Session("MSG") & " (Site Visit) " 
				ElseIf Request("SelLog") = 2 Then
					sqlTBL = sqlTBL & " AND NOT IsNull(phoneCon_last)"
					Session("Msg") = Session("MSG") & " (Phone Call) "
				ElseIf Request("SelLog") = 3 Then
					sqlTBL = sqlTBL & " AND NOT IsNull(MiscCon)"
					Session("Msg") = Session("MSG") & " (Misc. Contact) "
				End If
				If Request("SelCon") <> "0" Then 
					sqlTBL = sqlTBL & " AND Consumer_T.Medicaid_Number = '" & Request("SelCon") & "'"
					Session("Msg") = Session("Msg") & " of " & GetName2(Request("SelCon"))
				End If
				Err = 0
				If Request("FrmD8") <> "" Then
					If Not IsDate(Request("FrmD8")) Then
						Err = 1
					Else
						If Request("SelLog") = 1 Then
							sqlTBL = sqlTBL & " AND Site_V_Date >= #" & Request("FrmD8") & "#"
							Session("Msg") = Session("Msg") & " from " & Request("FrmD8")
						ElseIf Request("SelLog") = 2 Then
							sqlTBL = sqlTBL & " AND phoneCon_last >= #" & Request("FrmD8") & "#"
							Session("Msg") = Session("Msg") & " from " & Request("FrmD8")
						ElseIf Request("SelLog") = 3 Then
							sqlTBL = sqlTBL & " AND MiscCon >= #" & Request("FrmD8") & "#"
							Session("Msg") = Session("Msg") & " from " & Request("FrmD8")
						Else
							sqlTBL = sqlTBL & " AND (Site_V_Date >= #" & Request("FrmD8") & "#"
							sqlTBL = sqlTBL & " OR phoneCon_last >= #" & Request("FrmD8") & "#"
							sqlTBL = sqlTBL & " OR miscCon >= #" & Request("FrmD8") & "#)"
							Session("Msg") = Session("Msg") & " from " & Request("FrmD8")
						End If
					End If
				End If
				If Request("ToD8") <> "" Then
					If Not IsDate(Request("ToD8")) Then
						Err = 1
					Else
						If Request("SelLog") = 1 Then
							sqlTBL = sqlTBL & " AND Site_V_Date <= #" & Request("ToD8") & "#"
							Session("Msg") = Session("Msg") & " to " & Request("ToD8")
						ElseIf Request("SelLog") = 2 Then
							sqlTBL = sqlTBL & " AND phoneCon_last <= #" & Request("ToD8") & "#"
							Session("Msg") = Session("Msg") & " to " & Request("ToD8")
						ElseIf Request("SelLog") = 3 Then
							sqlTBL = sqlTBL & " AND MiscCon <= #" & Request("ToD8") & "#"
							Session("Msg") = Session("Msg") & " to " & Request("ToD8")
						Else
							sqlTBL = sqlTBL & " AND (Site_V_Date <= #" & Request("ToD8") & "#"
							sqlTBL = sqlTBL & " OR phoneCon_last <= #" & Request("ToD8") & "#"
							sqlTBL = sqlTBL & " OR miscCon <= #" & Request("ToD8") & "#)"
							Session("Msg") = Session("Msg") & " to " & Request("ToD8")
						End If
					End If
				End If
				If Err <> 0 Then  Response.Redirect "specrep.asp?err=41" 
				If Request("SelLog") = 1 Then
					sqlTBL = sqlTBL & " ORDER BY Site_V_Date DESC, Lname ASC, Fname ASC"
				ElseIf Request("SelLog") = 2 Then
					sqlTBL = sqlTBL & " ORDER BY phoneCon_last DESC, Lname ASC, Fname ASC"
				ElseIf Request("SelLog") = 3 Then
					sqlTBL = sqlTBL & " ORDER BY miscCon DESC, Lname ASC, Fname ASC"
				Else
					sqlTBL = sqlTBL & " ORDER BY Lname ASC, Fname ASC"
				End If
				rsTBL.Open sqlTBL, g_strCONN, 3, 1
				x = 0
				Do Until rsTBL.EOF
					If Request("SelLog") <> 0 Then
						strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("lname") & ", " & rsTBL("fname") & "</font></td>"
						If Request("SelLog") = 1 Then
							newComment = Replace(rsTBL("Comments"), "|",  " ")
							strBODY = strBODY & "<td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("Site_V_Date") & "&nbsp;</font></td>" & _
								 "<td align='center'><font size='1' face='trebuchet ms'>" & newComment & "&nbsp;</font></td></tr>"
						ElseIf Request("SelLog") = 2 Then
							strBODY = strBODY & "<td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("phoneCon_Last") & "&nbsp;</font></td>" & _
								"<td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("PCom") & "&nbsp;</font></td></tr>"
						ElseIf Request("SelLog") = 3 Then
							strBODY = strBODY & "<td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("miscCon") & "&nbsp;</font></td>" & _
								"<td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("Misc") & "&nbsp;</font></td></tr>"
						End If
					Else
						ReDim Preserve tmp(x)
						tmp(x) = rsTBL("C_Site_Visit_Dates_t.index") & "|" & rsTBL("Site_V_date") & "|" & rsTBL("phoneCon_last") & "|" & rsTBL("miscCon")
						x = x + 1
					End If
					rsTBL.MoveNExt
				Loop
				rsTBl.Close
				Set rsTBL = Nothing
				If Request("SelLog") = 0 Then
					For i = x - 2 to 0 Step - 1
						For j = 0 To i
							tmpj = split(tmp(j),"|")
							tmpj1 = split(tmp(j+1),"|")
							tmpDateLog = tmpj(1)
							tmpDateLog1 = tmpj1(1)
							if tmpDateLog = "" Then tmpDateLog = tmpj(2)
							if tmpDateLog = "" Then tmpDateLog = tmpj(3)
							if tmpDateLog1 = "" Then tmpDateLog1 = tmpj1(2)
							if tmpDateLog1 = "" Then tmpDateLog1 = tmpj1(3)
							If Cdate(tmpDateLog) < Cdate(tmpDateLog1) Then
								intTemp = tmp(j + 1)
				              tmp(j + 1) = tmp(j)
				              tmp(j) = intTemp
							End If
						Next 
					Next 
					Set rsTBL2 = Server.CreateObject("ADODB.RecordSet")
					zzz = 0
					Do Until zzz = x 
						tmp2 = split(tmp(zzz),"|")	
						sqlTBL2 = "SELECT * FROM Consumer_T, C_Site_Visit_Dates_T WHERE Consumer_t.Medicaid_number =" & _
							" C_Site_Visit_Dates_t.Medicaid_number AND C_Site_Visit_Dates_t.index = " & tmp2(0)
							
						rsTBL2.Open sqlTBL2, g_strCONN, 1, 3
						If Not rsTBL2.EOF Then
							strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL2("lname") & ", " & rsTBL2("fname") & "</font></td>" & _
								"<td align='center'><font size='1' face='trebuchet ms'>" & rsTBL2("Site_V_Date") & "&nbsp;</font></td>" & _
								"<td align='center'><font size='1' face='trebuchet ms'>" & rsTBL2("phoneCon_last") & "&nbsp;</font></td>" & _
								"<td align='center'><font size='1' face='trebuchet ms'>" & rsTBL2("miscCon") & "&nbsp;</font></td>"
							If rsTBL2("Site_V_Date") <> "" Then 
								newComment = Replace(rsTBL2("Comments"), "|",  " ")
								strBODY = strBODY & "<td align='center'><font size='1' face='trebuchet ms'>" & newComment & "&nbsp;</font></td></tr>"
							ElseIf rsTBL2("phoneCon_last") <> "" Then
								 strBODY = strBODY & "<td align='center'><font size='1' face='trebuchet ms'>" & rsTBL2("PCom") & "&nbsp;</font></td></tr>"
							ElseIf rsTBL2("miscCon") <> "" Then
								 strBODY = strBODY & "<td align='center'><font size='1' face='trebuchet ms'>" & rsTBL2("Misc") & "&nbsp;</font></td></tr>"
							End If
						End If
						rsTBL2.Close
						'Set rsTBL2 = Nothing
						zzz = zzz + 1
					Loop
				End If
			ElseIf Request("SelRep") = 42 Then
				Session("MSG") = "PCSP Worker Separation Code report."
				typ = Request("SelRep")
				strHEAD = "<tr bgcolor='#040C8B'><td align='center'><font size='1' face='trebuchet ms' color='white'><b>PCSP Worker Name</b>" & _
						"</font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Date of Hire</b></font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Date Of Termination</b></font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Separation Code</b></font></td></tr>"
				Set rsTBL = Server.CreateObject("ADODB.RecordSet")
				sqlTBL = "SELECT * " & _
					"FROM worker_t " & _
					"WHERE Status = 'InActive' ORDER BY Lname, Fname"
				rsTBL.Open sqlTBL, g_strCONN, 1, 3
				Do Until rsTBL.EOF	
					sepcode = "&nbsp;"
					If rsTBL("sep_code") <> "" Then sepcode =  rsTBL("sep_code")
					strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("lname") & ", " & _
						rsTBL("fname") & "</font></td><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("Date_Hired") & "</td>" & _
						"<td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("term_date") & "</font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms'>" & sepcode & "</font></td></tr>"
					rsTBL.MoveNext
				Loop
				rsTBL.Close 
			ElseIf Request("SelRep") = 43 Then
				Session("MSG") = "PCSP Workers' Driver Conviction and Criminal History report."	
				strHEAD = "<tr bgcolor='#040C8B'><td align='center'><font size='1' face='trebuchet ms' color='white'><b>PCSP Worker Name</b>" & _
					"</font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Address</b></font></td>" & _
					"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Phone No.</b></font></td>" & _
					"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Driver Convictions</b></font></td>" & _
					"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Criminal History</b></font></td>" & _
					"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Waiver Date</b></font></td></tr>"
				Set rsTBL = Server.CreateObject("ADODB.RecordSet")
				sqlTBL = "SELECT * FROM worker_T, w_files_T WHERE status = 'Active' AND SSN = Social_Security_Number ORDER BY Lname, Fname"
				rsTBL.Open sqlTBL, g_strCONN, 1, 3
				Do Until rsTBL.EOF
					If Not(rsTBL("DC1") = "" And rsTBL("DC2") = "" And rsTBL("DC3") = "" And rsTBL("DC4") = "" _
						And rsTBL("CH1") = ""  And rsTBL("CH2") = "" And rsTBL("CH3") = "" And rsTBL("CH4") = "") Then
						strBODY = strBODY & "<tr bgcolor='#d4d0c8'><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("lname") & ", " & _
							rsTBL("fname") & "</font></td><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("Address") & ", " & rsTBL("City") &  ", " & rsTBL("State") & ", " & rsTBL("Zip") & "</font></td>" & _
							"<td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("PhoneNo") & "&nbsp;</font></td>" & _
							"<td align='center' colspan='3'>&nbsp;</td></tr>"
					End If
					If rsTBL("DC1") <> "" Then 	strBODY = strBODY & "<tr><td align='center' colspan='3'></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("DC1") & "</font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms'>&nbsp;</font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("DC1date") & "&nbsp;</font></td></tr>"
					If rsTBL("DC2") <> "" Then strBODY = strBODY &  "<tr><td align='center' colspan='3'></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("DC2") & "</font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms'>&nbsp;</font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("DC2date") & "&nbsp;</font></td></tr>"
					If rsTBL("DC3") <> "" Then strBODY = strBODY &  "<tr><td align='center' colspan='3'></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("DC3") & "</font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms'>&nbsp;</font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("DC3date") & "&nbsp;</font></td></tr>"
					If rsTBL("DC4") <> "" Then strBODY = strBODY &  "<tr><td align='center' colspan='3'></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("DC4") & "</font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms'>&nbsp;</font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("DC4date") & "&nbsp;</font></td></tr>"
					If rsTBL("CH1") <> "" Then strBODY = strBODY & "<tr><td align='center' colspan='3'></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms'>&nbsp;</font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("CH1") & "</font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("CH1date") & "&nbsp;</font></td></tr>"
					If rsTBL("CH2") <> "" Then strBODY = strBODY &  "<tr><td align='center' colspan='3'></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms'>&nbsp;</font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("CH2") & "</font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("CH2date") & "&nbsp;</font></td></tr>"	
					If rsTBL("CH3") <> "" Then strBODY = strBODY &  "<tr><td align='center' colspan='3'></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms'>&nbsp;</font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("CH3") & "</font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("CH3date") & "&nbsp;</font></td></tr>"	
					If rsTBL("CH4") <> "" Then strBODY = strBODY &  "<tr><td align='center' colspan='3'></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms'>&nbsp;</font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("CH4") & "</font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("CH4date") & "&nbsp;</font></td></tr>"		
					rsTBL.MoveNext
				Loop
				rsTBL.Close
				Set rsTBL = Nothing
			ElseIf Request("SelRep") = 44 Then
				Session("MSG") = "PCSP Workers over 80 hours"	
				If Request("closedate") = "" Then
					Session("MSG") = "Please select a pay period."
					Response.Redirect "SpecRep.asp"
				End If
				Pdate = Request("closedate")
				Session("MSG") = Session("MSG") & " from " & Pdate
				Pdate2 = Request("Todate")
				Session("MSG") = Session("MSG") & " to " & Pdate2
				Session("MSG") = Session("MSG") & " report." 
				strHEAD = "<tr bgcolor='#040C8B'><td align='center'><font size='1' face='trebuchet ms' color='white'><b>PCSP Worker Name</b>" & _
					"</font></td>" & _
					"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Hours</b></font></td>" & _
					"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>RIHCC</b></font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>RIHCC2</b></font></td></tr>" 
				Set rsTBL = Server.CreateObject("ADODB.RecordSet")
				sqlTBL  = "SELECT * FROM Tsheets_t, worker_T, proj_man_T, consumer_T  WHERE consumer_T.Medicaid_Number = client AND " & _
					"PMID = Proj_Man_T.ID AND emp_id = worker_T.social_security_number AND date <= #" & Pdate2 & "# AND date >= #" & Pdate & "#" & _
					" ORDER BY proj_man_T.Lname, proj_man_T.Fname, worker_T.Lname, worker_T.Fname"
				rsTBL.Open sqlTBL, g_strCONN, 1, 3
				x = 0
				Do Until rsTBL.EOF
					strEmpID = rsTBL("client")
					strWorID = rsTBL("emp_id")
					dblHours = rsTBL("mon") + rsTBL("tue") + rsTBL("wed") + rsTBL("thu") + rsTBL("fri") + rsTBL("sat") + rsTBL("sun")
					strDate = rsTBL("date")
					strPMID = rsTBL("PM1")
					strPMID2 = rsTBL("PM2")
					strWeekLabel = Z_Find2WkPeriod(tmpDate)
					lngIdx = SearchArrays2(strWeekLabel,  strEmpID, strWorID, tmpDates, tmpEmpID, tmpWorID)
					If lngIdx < 0 Then ' this is the first time i've encountered the date and id pair, so i make a new entry
						ReDim Preserve tmpDates(x)
						ReDim Preserve tmpWorID(x)
						ReDim Preserve tmpEmpID(x)
						ReDim Preserve tmpHrs(x)
						ReDim Preserve tmpPMID(x)
						ReDim Preserve tmpPMID2(x)
						
						tmpDates(x) = strWeekLabel
						tmpEmpID(x) = strEmpID
						tmpWorID(x) = strWorID
						tmpHrs(x) = dblHours
						tmpPMID(x) = strPMID
						tmpPMID2(x) = strPMID2
						x = x + 1
					Else
						tmpHrs(lngIdx) = tmpHrs(lngIdx) + dblHours
						
					End If
					rsTBL.MoveNext	
				Loop
				rsTBL.Close
				Set rsTBL = Nothing
				y = 0
				Do Until y = x 
					If tmpHrs(y) > 80 Then
						strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & GetName(tmpWorID(y)) & _
							"</font></td>" & _
							"<td align='center'><font size='1' face='trebuchet ms'>" & tmpHrs(y) & "</font></td>" & _
							"<td align='center'><font size='1' face='trebuchet ms'>" & GetName3(tmpPMID(y)) & "</font></td><td align='center'><font size='1' face='trebuchet ms'>" & GetName3(tmpPMID2(y)) & "</font></td></tr>"
					End If
					y = y + 1
				Loop 
			ElseIf Request("SelRep") = 45 Then
				Session("MSG") = "Consumer on hold report"
				strHEAD = "<tr bgcolor='#040C8B'><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Consumer</b>" & _
					"</font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>RIHCC</b></font></td>" & _
					"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Date</b></font></td>" & _
					"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Reason</b></font></td></tr>" 
				Set rsTBL = Server.CreateObject("ADODB.RecordSet")	
				sqlTBL = "SELECT * FROM Consumer_T, C_Status_T WHERE Consumer_T.medicaid_number = C_Status_T.medicaid_number AND onhold=true ORDER BY lname, fname"
				rsTBL.Open sqlTBL, g_strCONN, 1, 3
				Do Until rsTBL.EOF
					strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("lname") & ", " & rsTBL("fname") & _
							"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & GetName3(rsTBL("PMID")) & _
							"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("H_From_Date") & " - " & rsTBL("H_To_Date")  & "</font></td>"
					tmpReason = ""
					If rsTBL("In_Hospital") = True Then tmpReason = tmpReason & "In Hospital/Rehab/Nursing Home" & vbCrlf
					If rsTBL("New_Worker") = True Then tmpReason = tmpReason & "Needs new worker" & vbCrlf
					If rsTBL("H_Other") <> "" Then tmpReason = tmpReason &  rsTBL("H_Other") & vbCrlf
					If tmpReason = "" Then tmpReason = "&nbsp;"
					strBODY = strBODY & "<td align='center'><font size='1' face='trebuchet ms'>" & tmpReason & "</font></td></tr>" & _
							
					rsTBL.MoveNext
				Loop
				rsTBL.CLose
				Set rsTBL =Nothing
			ElseIf Request("SelRep") = 46 Then
				Session("MSG") = "Consumer with no PCSP Worker report"
				strHEAD = "<tr bgcolor='#040C8B'><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Consumer</b>" & _
					"</font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Address</b></font></td>" & _
					"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Phone No.</b></font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>RIHCC</b></font></td></tr>" 
				Set rsTBL = Server.CreateObject("ADODB.RecordSet")	
				sqlTBL =  "SELECT * FROM consumer_T, C_Status_T, proj_man_T  WHERE PMID = proj_man_T.id AND Consumer_T.medicaid_number = C_Status_T.medicaid_number AND Active = True ORDER BY proj_man_T.Lname, proj_man_T.fname, Consumer_T.Lname, Consumer_T.Fname"
				rsTBL.Open sqlTBL, g_strCONN, 1, 3
				Do Until rsTBL.EOF
					Set rsLink = Server.CreateObject("ADODB.RecordSet")
					sqlLink = "SELECT * FROM ConWork_t WHERE CID = '" & rsTBL("consumer_T.medicaid_number") & "' "
					rsLink.Open sqlLink, g_strCONN, 1, 3
					If rsLink.EOF Then
						strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & _
							rsTBL("Consumer_T.lname") & ", " & rsTBL("Consumer_T.fname") & "</font></td><td align='center'><font size='1' face='trebuchet ms'>" & _
							rsTBL("Address") & ", " & rsTBL("City") & ", " & rsTBL("State") & ", " & rsTBL("Zip") & "</font></td>" & _
							"<td align='center'><font size='1' face='trebuchet ms'>" &  rsTBL("PhoneNo") & "</td>" & _
							"<td align='center'><font size='1' face='trebuchet ms'>" &  rsTBL("proj_man_T.Lname") & ", " & rsTBL("proj_man_T.fname") & "</td></tr>"
					End If
					rsLink.Close
					Set rsLink = Nothing
					rsTBL.MoveNext
				Loop
				rsTBL.CLose
				Set rsTBL =Nothing
			ElseIf Request("SelRep") = 47 Then
				Session("MSG") = "PCSP Worker logs"
					strHEAD = "<tr bgcolor='#040C8B'><td align='center'><font size='1' face='trebuchet ms' color='white'><b>PCSP Worker Name</b></font></td>"
					If Request("SelLog2") = 1 Then
						strHEAD = strHEAD & "<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Site Visit Date</b></font></td>"
						strHEAD = strHEAD & "<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Comment</b></font></td></tr>"
					ElseIf Request("SelLog2") = 2 Then
						strHEAD = strHEAD & "<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Phone Call Date</b></font></td>"
						strHEAD = strHEAD & "<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Comment</b></font></td></tr>"
					ElseIf Request("SelLog2") = 3 Then
						strHEAD = strHEAD & "<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Misc. Contact Date</b></font></td>"
						strHEAD = strHEAD & "<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Comment</b></font></td></tr>"
					Else
						strHEAD = strHEAD & "<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Site Visit Date</b></font></td>"
						strHEAD = strHEAD & "<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Phone Call Date</b></font></td>"
						strHEAD = strHEAD & "<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Misc. Contact Date</b></font></td>"
						strHEAD = strHEAD & "<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Comment</b></font></td></tr>"
					End If
					Set rsTBL = Server.CreateObject("ADODB.RecordSet")
					sqlTBL = "SELECT * FROM Worker_T, W_Log_T WHERE Worker_T.social_security_number = w_log_T.ssn" & _
						" AND status = 'Active'"
					If Request("SelLog2") = 1 Then
						sqlTBL = sqlTBL & " AND NOT IsNull(siteV)"
						Session("Msg") = Session("MSG") & " (Site Visit) " 
					ElseIf Request("SelLog2") = 2 Then
						sqlTBL = sqlTBL & " AND NOT IsNull(phoneC)"
						Session("Msg") = Session("MSG") & " (Phone Call) "
					ElseIf Request("SelLog2") = 3 Then
						sqlTBL = sqlTBL & " AND NOT IsNull(misc)"
						Session("Msg") = Session("MSG") & " (Misc. Contact) "
					End If
					If Request("SelWor") <> "0" Then 
						sqlTBL = sqlTBL & " AND worker_T.social_security_number = '" & Request("SelWor") & "'"
						Session("Msg") = Session("Msg") & " of " & GetName(Request("SelWor"))
					End If
						If Request("FrmD8") <> "" Then
							If Not IsDate(Request("FrmD8")) Then
								Err = 1
							Else
								If Request("SelLog2") = 1 Then
									sqlTBL = sqlTBL & " AND siteV >= #" & Request("FrmD8") & "#"
									Session("Msg") = Session("Msg") & " from " & Request("FrmD8")
								ElseIf Request("SelLog2") = 2 Then
									sqlTBL = sqlTBL & " AND phoneC >= #" & Request("FrmD8") & "#"
									Session("Msg") = Session("Msg") & " from " & Request("FrmD8")
								ElseIf Request("SelLog2") = 3 Then
									sqlTBL = sqlTBL & " AND misc >= #" & Request("FrmD8") & "#"
									Session("Msg") = Session("Msg") & " from " & Request("FrmD8")
								Else
									sqlTBL = sqlTBL & " AND (siteV >= #" & Request("FrmD8") & "#"
									sqlTBL = sqlTBL & " OR phoneC >= #" & Request("FrmD8") & "#"
									sqlTBL = sqlTBL & " OR misc >= #" & Request("FrmD8") & "#)"
									Session("Msg") = Session("Msg") & " from " & Request("FrmD8")
								End If
							End If
						End If
						If Request("ToD8") <> "" Then
						If Not IsDate(Request("ToD8")) Then
							Err = 1
						Else
							If Request("SelLog2") = 1 Then
								sqlTBL = sqlTBL & " AND siteV <= #" & Request("ToD8") & "#"
								Session("Msg") = Session("Msg") & " to " & Request("ToD8")
							ElseIf Request("SelLog2") = 2 Then
								sqlTBL = sqlTBL & " AND phoneC <= #" & Request("ToD8") & "#"
								Session("Msg") = Session("Msg") & " to " & Request("ToD8")
							ElseIf Request("SelLog2") = 3 Then
								sqlTBL = sqlTBL & " AND misc <= #" & Request("ToD8") & "#"
								Session("Msg") = Session("Msg") & " to " & Request("ToD8")
							Else
								sqlTBL = sqlTBL & " AND (siteV <= #" & Request("ToD8") & "#"
								sqlTBL = sqlTBL & " OR phoneC <= #" & Request("ToD8") & "#"
								sqlTBL = sqlTBL & " OR misc <= #" & Request("ToD8") & "#)"
								Session("Msg") = Session("Msg") & " to " & Request("ToD8")
							End If
						End If
					End If
					If Err <> 0 Then  Response.Redirect "specrep.asp?err=47" 
					If Request("SelLog") = 1 Then
						sqlTBL = sqlTBL & " ORDER BY siteV DESC, Lname ASC, Fname ASC"
					ElseIf Request("SelLog") = 2 Then
						sqlTBL = sqlTBL & " ORDER BY phoneC DESC, Lname ASC, Fname ASC"
					ElseIf Request("SelLog") = 3 Then
						sqlTBL = sqlTBL & " ORDER BY misc DESC, Lname ASC, Fname ASC"
					Else
						sqlTBL = sqlTBL & " ORDER BY Lname ASC, Fname ASC"
					End If
					response.write "<!-- HOY: " & sqlTBL & " --->"
					rsTBL.Open sqlTBL, g_strCONN, 3, 1
				
				x = 0
				Do Until rsTBL.EOF
					If Request("SelLog2") <> 0 Then
						strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("lname") & ", " & rsTBL("fname") & "</font></td>"
						If Request("SelLog2") = 1 Then
							newComment = Replace(rsTBL("scom"), "|",  " ")
							strBODY = strBODY & "<td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("sitev") & "&nbsp;</font></td>" & _
								 "<td align='center'><font size='1' face='trebuchet ms'>" & newComment & "&nbsp;</font></td></tr>"
						ElseIf Request("SelLog2") = 2 Then
							strBODY = strBODY & "<td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("phonec") & "&nbsp;</font></td>" & _
								"<td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("PCom") & "&nbsp;</font></td></tr>"
						ElseIf Request("SelLog2") = 3 Then
							strBODY = strBODY & "<td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("misc") & "&nbsp;</font></td>" & _
								"<td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("mcom") & "&nbsp;</font></td></tr>"
						End If
					Else
						ReDim Preserve tmp(x)
						tmp(x) = rsTBL("w_log_T.index") & "|" & rsTBL("sitev") & "|" & rsTBL("phoneC") & "|" & rsTBL("misc")
						x = x + 1
					End If
					rsTBL.MoveNExt
				Loop
				rsTBl.Close
				Set rsTBL = Nothing
				If Request("SelLog2") = 0 Then
					For i = x - 2 to 0 Step - 1
						For j = 0 To i
							tmpj = split(tmp(j),"|")
							tmpj1 = split(tmp(j+1),"|")
							tmpDateLog = tmpj(1)
							tmpDateLog1 = tmpj1(1)
							if tmpDateLog = "" Then tmpDateLog = tmpj(2)
							if tmpDateLog = "" Then tmpDateLog = tmpj(3)
							if tmpDateLog1 = "" Then tmpDateLog1 = tmpj1(2)
							if tmpDateLog1 = "" Then tmpDateLog1 = tmpj1(3)
							If Cdate(tmpDateLog) < Cdate(tmpDateLog1) Then
								intTemp = tmp(j + 1)
				              tmp(j + 1) = tmp(j)
				              tmp(j) = intTemp
							End If
						Next 
					Next 
					Set rsTBL2 = Server.CreateObject("ADODB.RecordSet")
					zzz = 0
					Do Until zzz = x 
						tmp2 = split(tmp(zzz),"|")	
						sqlTBL2 = "SELECT * FROM worker_T, w_log_T WHERE worker_T.social_security_number =" & _
							" w_log_T.ssn AND w_log_T.index = " & tmp2(0)
							
						rsTBL2.Open sqlTBL2, g_strCONN, 1, 3
						If Not rsTBL2.EOF Then
							strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL2("lname") & ", " & rsTBL2("fname") & "</font></td>" & _
								"<td align='center'><font size='1' face='trebuchet ms'>" & rsTBL2("sitev") & "&nbsp;</font></td>" & _
								"<td align='center'><font size='1' face='trebuchet ms'>" & rsTBL2("phoneC") & "&nbsp;</font></td>" & _
								"<td align='center'><font size='1' face='trebuchet ms'>" & rsTBL2("misc") & "&nbsp;</font></td>"
							If rsTBL2("sitev") <> "" Then 
								newComment = Replace(rsTBL2("scom"), "|",  " ")
								strBODY = strBODY & "<td align='center'><font size='1' face='trebuchet ms'>" & newComment & "&nbsp;</font></td></tr>"
							ElseIf rsTBL2("phoneC") <> "" Then
								 strBODY = strBODY & "<td align='center'><font size='1' face='trebuchet ms'>" & rsTBL2("pcom") & "&nbsp;</font></td></tr>"
							ElseIf rsTBL2("misc") <> "" Then
								 strBODY = strBODY & "<td align='center'><font size='1' face='trebuchet ms'>" & rsTBL2("mcom") & "&nbsp;</font></td></tr>"
							End If
						End If
						rsTBL2.Close
						'Set rsTBL2 = Nothing
						zzz = zzz + 1
					Loop
				'End If
				End If	
			ElseIf Request("SelRep") = 48 Then
				Session("MSG") = "Consumers hours billed."
				strHEAD = "<tr bgcolor='#040C8B'><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Start Date</b>" & _
					"</font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Consumer</b></font></td>" & _
					"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>RIHCC</b></font></td>" & _
					"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Last Submitted Date</b></font></td></tr>" 
				Set rsTBL = Server.CreateObject("ADODB.RecordSet")	
				sqlTBL =  "SELECT * FROM consumer_T, C_Status_T,  proj_man_T WHERE Proj_man_T.ID = PMID AND Consumer_T.medicaid_number = C_Status_T.medicaid_number " & _
					"AND Active = True ORDER BY proj_man_T.Lname, proj_man_T.fname, consumer_T.lname, consumer_T.fname"
				rsTBL.Open sqlTBL, g_strCONN, 1, 3
				Do Until rsTBL.EOF	
					Set rsTS = Server.CreateObject("ADODB.RecordSet")
					sqlTS = "SELECT * FROM Tsheets_T WHERE client = '" & rsTBL("Consumer_T.medicaid_number") & "' AND NOT IsNull(ProcPay) ORDER BY Date DESC"
					rsTS.Open sqlTS, g_strCONN, 1, 3
					If Not rsTS.EOF Then
						Cname = rsTBL("consumer_T.lname") & ", " & rsTBL("consumer_T.fname")
						PMname = rsTBL("proj_man_T.lname") & ", " & rsTBL("proj_man_T.fname")
						strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("start_date") & _
							"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & Cname & "</font></td>" & _
							"<td align='center'><font size='1' face='trebuchet ms'>" & PMname & "</font></td>" & _
							"<td align='center'><font size='1' face='trebuchet ms'>" &  rsTS("ProcPay") & "</td></tr>"	
					End If
					rsTS.Close
					Set rsTS = Nothing
					rsTBL.MoveNext
				Loop
				rsTBL.CLose
				Set rsTBL = Nothing
			ElseIf Request("SelRep") = 49 Then
				Session("MSG") = "Finance Consumer List report"
				strHEAD = "<tr bgcolor='#040C8B'><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Medicaid</b>" & _
						"</font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Consumer Name</b>" & _
						"</font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Address</b>" & _
						"</font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>DOB</b>" & _
						"</font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Gender</b>" & _
						"</font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Start Date</b></font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>End Date</b></font></td></tr>"
				Set rsTBL = Server.CreateObject("ADODB.RecordSet")
				sqlTBL = "SELECT * FROM consumer_t, c_status_t WHERE consumer_t.medicaid_number = c_status_t.medicaid_number AND Active = true"
				If Request("FrmD8") <> "" Then
					If IsDate(Request("FrmD8")) Then
						sqlTBL = sqlTBL & " AND Start_date >= #" & Request("FrmD8") & "# " 
						Session("Msg") = Session("Msg") & " from " & Request("FrmD8")
					End If
				End If
				If Request("ToD8") <> "" Then
					If IsDate(Request("ToD8")) Then
						sqlTBL = sqlTBL & " AND Start_date <= #" & Request("ToD8") & "# " 
						Session("Msg") = Session("Msg") & " to " & Request("ToD8")
					End If
				End If
				sqlTBL  = sqlTBL  & " ORDER BY start_date DESC, lname, fname"
				Session("Msg") = Session("Msg") & ". " 
				'response.write sqlTBL
				rsTBL.Open sqlTBL, g_strCONN, 1, 3
				Do Until rsTBL.EOF	
					tmpADR = rsTBL("Address") & ", " & rsTBL("City") & ", " & rsTBL("State") & ", " & rsTBL("Zip")
					strBODY = strBODY & "<tr>" 
					If Session("lngType") = 1 Or Session("lngType") = 2 Then
						strBODY = strBODY & "<td align='center'><font size='1' face='trebuchet ms'>" & _
						rsTBL("consumer_t.medicaid_number") &"</font></td>"
					Else
							strBODY = strBODY & "<Td>&nbsp;</td>"
						End If	
					strBODY = strBODY & "<td align='center'><font size='1' face='trebuchet ms'>" & _
						rsTBL("lname") & ", " & rsTBL("fname") &"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & _
						tmpADR &"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & _
						rsTBL("DOB") &"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & _
						rsTBL("Gender") &"</font></td><td align='center'><font size='1' face='trebuchet ms'>&nbsp;" & rsTBL("Start_date") & _
						"</font></td><td align='center'><font size='1' face='trebuchet ms'>&nbsp;" & _
						rsTBL("Termdate") &  "</font></td></tr>"
					rsTBL.MoveNext
				Loop
				rsTBL.Close
				Set rsTBL = Nothing		
			ElseIf Request("SelRep") = 50 Then
				Session("MSG") = "PCSP Workers over 40 hours"	
				If Request("frmd8") = "" Then
					Session("MSG") = "Date required."
					Response.Redirect "SpecRep.asp"
				End If
				If Not IsDate(Request("frmd8")) Then
					Session("MSG") = "Invalid Date."
					Response.Redirect "SpecRep.asp"
				End If
				Session("MSG") = Session("MSG") & " from " & GetSun(Request("frmd8"))
				Session("MSG") = Session("MSG") & " to " &  GetSat(Request("frmd8"))
				Session("MSG") = Session("MSG") & " report." 
				strHEAD = "<tr bgcolor='#040C8B'><td align='center'><font size='1' face='trebuchet ms' color='white'><b>PCSP Worker Name</b>" & _
					"</font></td>" & _
					"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Hours</b></font></td>" & _
					"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>RIHCC</b></font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>RIHCC2</b></font></td></tr>" 
				Set rsTBL = Server.CreateObject("ADODB.RecordSet")
				sqlTBL  = "SELECT * FROM Tsheets_t, worker_T, proj_man_T, consumer_T  WHERE consumer_T.Medicaid_Number = client AND " & _
					"PMID = Proj_Man_T.ID AND emp_id = worker_T.social_security_number AND date >= #" & GetSun(Request("frmd8")) & "# AND date <= #" & GetSat(Request("frmd8")) & "#" & _
					" ORDER BY proj_man_T.Lname, proj_man_T.Fname, worker_T.Lname, worker_T.Fname"
					'response.write sqltBl
				rsTBL.Open sqlTBL, g_strCONN, 1, 3	
				x = 0
				Do Until rsTBL.EOF
					strEmpID = rsTBL("client")
					strWorID = rsTBL("emp_id")
					dblHours = rsTBL("mon") + rsTBL("tue") + rsTBL("wed") + rsTBL("thu") + rsTBL("fri") + rsTBL("sat") + rsTBL("sun")
					strDate = rsTBL("date")
					strPMID = rsTBL("PM1")
					strPMID2 = rsTBL("PM2")
					strWeekLabel = GetSun(Request("frmd8"))
					lngIdx = SearchArrays2(strWeekLabel,  strEmpID, strWorID, tmpDates, tmpEmpID, tmpWorID)
					If lngIdx < 0 Then ' this is the first time i've encountered the date and id pair, so i make a new entry
						ReDim Preserve tmpDates(x)
						ReDim Preserve tmpWorID(x)
						ReDim Preserve tmpEmpID(x)
						ReDim Preserve tmpHrs(x)
						ReDim Preserve tmpPMID(x)
						ReDim Preserve tmpPMID2(x)
						
						tmpDates(x) = strWeekLabel
						tmpEmpID(x) = strEmpID
						tmpWorID(x) = strWorID
						tmpHrs(x) = dblHours
						tmpPMID(x) = strPMID
						tmpPMID2(x) = strPMID2
						x = x + 1
					Else
						tmpHrs(lngIdx) = tmpHrs(lngIdx) + dblHours
						
					End If
					rsTBL.MoveNext	
				Loop
				rsTBL.Close
				Set rsTBL = Nothing
				y = 0
				Do Until y = x 
					If tmpHrs(y) > 40 Then
						strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & GetName(tmpWorID(y)) & _
							"</font></td>" & _
							"<td align='center'><font size='1' face='trebuchet ms'>" & tmpHrs(y) & "</font></td>" & _
							"<td align='center'><font size='1' face='trebuchet ms'>" & GetName3(tmpPMID(y)) & "</font></td><td align='center'><font size='1' face='trebuchet ms'>" & GetName3(tmpPMID2(y)) & "</font></td></tr>"
					End If
					y = y + 1
				Loop 
			ElseIf Request("SelRep") = 51 Then
				Session("MSG") = "Private Pay Consumers from " & Request("closedate") & " to " & Request("Todate")
				strHEAD = "<tr bgcolor='#040C8B'><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Name</b>" & _
					"</font></td>" & _
					"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>ID</b></font></td>" & _
					"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Timesheet Week</b></font></td>" & _
					"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Hours</b></font></td>" & _
					"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>PCSP Worker</b></font></td></tr>"
					
				Set rsTBL = Server.CreateObject("ADODB.RecordSet")	
				
				sqlTBL = "SELECT * FROM tsheets_T, Consumer_T, c_status_T WHERE date >= #" & Request("closedate") & "# AND date <= #" & _
					Request("Todate") & "# AND (client LIKE '%Private Pay%' OR (code = 'P' OR code = 'C')) AND consumer_t.medicaid_number = c_status_t.medicaid_number " & _
					"AND Active = true AND Client = consumer_t.medicaid_number ORDER BY client"

				rsTBL.Open sqlTBL, g_strCONN, 1,3 
				x = 0 
				If Not rsTBL.EOF Then
					Do Until rsTBL.EOF
						myDATE2 = rsTBL("date")
						strEmpID = rsTBL("client")
						strWorID = rsTBL("emp_id")
						If Not IsNull(rsTBL("lname")) Then
							strName = Replace(rsTBL("lname"),",","") & ", " & rsTBL("fname")
						Else
							strName = rsTBL("lname") & ", " & rsTBL("fname")
						End If
						dblHours = rsTBL("mon") + rsTBL("tue") + rsTBL("wed") + rsTBL("thu") + rsTBL("fri") + rsTBL("sat") + rsTBL("sun")
						' find the 2-week period
						strWeekLabel = Z_Find2WkPeriod(myDate2)
						' search for it in the arrays
						lngIdx = SearchArrays(strWeekLabel, strEmpID, strWorID, tmpDates, tmpEmpID, tmpWorID)
						If lngIdx < 0 Then
							ReDim Preserve tmpDates(x)
							ReDim Preserve tmpWorID(x)
							ReDim Preserve tmpEmpID(x)
							ReDim Preserve tmpHrs(x)
							
							tmpDates(x) = strWeekLabel
							tmpEmpID(x) = strEmpID
							tmpWorID(x) = strWorID
							tmpHrs(x) = dblHours
							x = x + 1
						Else
							
							tmpHrs(lngIdx) = tmpHrs(lngIdx) + dblHours
							
						End If
						rsTBL.MoveNext
					Loop
				End If
				rsTBL.Close
				Set rsTBL = Nothing	
				y = 0
				myHours= 0
				Do Until y = x 
					strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & GetName2(tmpEmpID(y)) & _
						"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & tmpEmpID(y) & _
						"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & tmpDates(y) & _
						"</font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms'>" & tmpHrs(y) & "</font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms'>" & GetName(tmpWorID(y)) & "</font></td></tr>"
						myHours = myHours + tmpHrs(y)
					y = y + 1
				Loop 
				strBODY = strBODY & "<tr><td colspan='4' align='right'><font size='1' face='trebuchet ms'>TOTAL:</font></td><td align='center'><font size='1' face='trebuchet ms'>" & _
					Z_FormatNumber(myHours, 2) & "</font></td></tr>" 
				'strHEAD = "<tr bgcolor='#040C8B'><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Last Name</b>" & _
				'	"</font></td>" & _
				'	"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>First Name</b></font></td>" & _
				'	"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Address</b></font></td>" & _
				'	"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Phone</b></font></td>" & _
				'	"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Status</b></font></td></tr>"
				'Set rsTBL = Server.CreateObject("ADODB.RecordSet")	
				'sqlTBL = "SELECT * FROM Consumer_T, c_status_T WHERE Consumer_T.Medicaid_Number = c_status_T.Medicaid_Number " & _
				'	"AND Consumer_T.Medicaid_Number LIKE '%Private Pay%' ORDER BY Active ASC, Lname, Fname"
				'rsTBL.Open sqlTBL, g_strCONN, 1, 3	
				'Do Until rsTBL.EOF
				'	mystat = "Inactive"
				'	If rsTBL("Active") = True Then mystat = "Active"
				'	tmpAdr = rsTBL("Address") & ", " & rsTBL("City") & ", " & rsTBL("State") & ", " & rsTBL("zip")
				'	strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("lname") & _
				'		"</font></td>" & _
				'		"<td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("fname") & "</font></td>" & _
				'		"<td align='center'><font size='1' face='trebuchet ms'>" & tmpAdr & "</font></td>" &_
				'		"<td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("PhoneNo") & "</font></td>" & _
				'		"<td align='center'><font size='1' face='trebuchet ms'>" & mystat & "</font></td</tr>"
				'	rsTBL.MoveNext
				'Loop
				'rsTBL.Close
				'Set rsTBL = Nothing
			ElseIf Request("SelRep") = 52 Then
				Session("MSG") = "PCSP Worker Violations"	
				strHEAD = "<tr bgcolor='#040C8B'><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Last Name</b>" & _
					"</font></td>" & _
					"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>First Name</b></font></td>" & _
					"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>RIHCC</b></font></td>" & _
					"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>RIHCC2</b></font></td>" & _
					"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Warning Date</b></font></td>" & _
					"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Violation</b></font></td></tr>"
				Set rsTBL = Server.CreateObject("ADODB.RecordSet")	
				sqlTBL = "SELECT * FROM worker_T, w_vio_T, proj_man_T WHERE ssn = worker_T.Social_Security_Number " & _
					"AND PM1 = Proj_Man_T.ID"
				If Request("SelWor") <> "0" Then 
					sqlTBL = sqlTBL & " AND worker_T.social_security_number = '" & Request("SelWor") & "'"
					Session("Msg") = Session("Msg") & " of " & GetName(Request("SelWor"))
				End If
				If Request("FrmD8") <> "" Then
							If Not IsDate(Request("FrmD8")) Then
								Err = 1
							Else
								sqlTBL = sqlTBL & " AND viodate >= #" & Request("FrmD8") & "#"
								Session("Msg") = Session("Msg") & " from " & Request("FrmD8")
							End If
						End If
						If Request("ToD8") <> "" Then
						If Not IsDate(Request("ToD8")) Then
							Err = 1
						Else
							sqlTBL = sqlTBL & " AND viodate <= #" & Request("ToD8") & "#"
							Session("Msg") = Session("Msg") & " to " & Request("ToD8")
						End If
					End If
					sqlTBL = sqlTBL & " ORDER BY Proj_Man_T.Lname, Proj_Man_T.Fname, worker_t.lname, viodate"
					If Err <> 0 Then  Response.Redirect "specrep.asp?err=47" 
				rsTBL.Open sqlTBL, g_strCONN, 1, 3	
				Do Until rsTBL.EOF
					APM = GetAPM(rsTBL("worker_T.index"))
					strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("worker_T.lname") & _
						"</font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("worker_T.fname") & "</font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms'>" & GetName3(rsTBL("PM1")) & "</font></td>" &_
						"<td align='center'><font size='1' face='trebuchet ms'>" & GetName3(rsTBL("PM2")) & "</font></td>" &_
						"<td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("viodate") & "</font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("notes") & "</font></td</tr>"
					rsTBL.MoveNext
				Loop
				rsTBL.Close
				Set rsTBL = Nothing
			ElseIf Request("SelRep") = 53 Then
				Session("MSG") = "Consumers' First Billable Date"	
				strHEAD = "<tr bgcolor='#040C8B'><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Last Name</b>" & _
					"</font></td>" & _
					"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>First Name</b></font></td>" & _
					"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>First Billable Date</b></font></td>" & _
					"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Referral Date</b></font></td>" & _
					"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Start Date</b></font></td></tr>"
				Set rsTBL = Server.CreateObject("ADODB.RecordSet")
					
				sqlTBL = "SELECT * FROM consumer_T, C_Status_T WHERE consumer_T.Medicaid_Number = C_Status_T.Medicaid_Number AND ACtive = True ORDER BY lname, fname ASC"
				rsTBL.Open sqlTBL, g_strCONN, 1, 3	
				Do Until rsTBL.EOF
					Set rsTBL2 = Server.CreateObject("ADODB.RecordSet")
					sqlTBL2 = "SELECT * FROM tsheets_T WHERE client = '" & rsTBL("consumer_T.Medicaid_Number") & "' AND NOT ISNULL(procmed) ORDER BY Date ASC"
					rsTBL2.Open sqlTBL2, g_strCONN, 1, 3	
					If Not rsTBL2.EOF Then
						strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("lname") & _
							"</font></td>" & _
							"<td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("fname") & "</font></td>" & _
							"<td align='center'><font size='1' face='trebuchet ms'>" & rsTBL2("Date") & "</font></td>" & _
							"<td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("Referral_Date") & "</font></td>" & _
							"<td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("Start_Date") & "</font></td</tr>"
					End If
					rstbl2.CLose
					Set rstbl2 = Nothing
					rsTBL.MoveNext
				Loop
				rsTBL.Close
				Set rsTBL = Nothing
			ElseIf Request("SelRep") = 54 Then
				Session("MSG") = "Consumers' Mileage Cap"	
				strHEAD = "<tr bgcolor='#040C8B'><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Last Name</b>" & _
				"</font></td>" & _
				"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>First Name</b></font></td>" & _
				"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>RIHCC</b></font></td>" & _
				"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>PCSP Worker</b></font></td>" & _
				"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Max Hours</b></font></td>" & _
				"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Mileage Cap</b></font></td></tr>"
				Set rsTBL = Server.CreateObject("ADODB.RecordSet")
				sqlTBL = "SELECT * FROM consumer_T, c_Status_T, conwork_T, worker_T WHERE consumer_T.Medicaid_Number = C_Status_T.Medicaid_Number " & _
					"AND ACtive = True AND Cint(WID) = worker_T.index AND CID = Consumer_T.Medicaid_Number ORDER BY Consumer_T.lname, Consumer_T.fname, PMID"
					
				rsTBL.Open sqlTBl, g_strCONN, 1, 3
				If Not rsTBL.EOF Then
					Do Until rsTBl.EOF
						strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("Consumer_T.lname") & _
								"</font></td>" & _
								"<td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("Consumer_T.fname") & "</font></td>" & _
								"<td align='center'><font size='1' face='trebuchet ms'>" & GetName3(rsTBL("PMID")) & "</font></td>" & _
								"<td align='center'><font size='1' face='trebuchet ms'>" & GetName(rsTBL("worker_T.Social_Security_Number")) & "</font></td>" & _
								"<td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("MaxHrs") & "</font></td>" & _
								"<td align='center'><font size='1' face='trebuchet ms'>" & Z_CZero(rsTBL("milecap")) & "</font></td></tr>" 
						rsTBL.MoveNext
					Loop
				End If 
				rsTBL.Close
				Set rsTL = Nothing	
			ElseIf Request("SelRep") = 55 Then
				SEssion("MSG") = "Billed Mileage report"
				Set rsTBL = Server.CreateObject("ADODB.RecordSet")
				sqlTBL = "SELECT * FROM Tsheets_T, consumer_T, worker_T WHERE Medicaid_Number = client AND emp_ID = Worker_T.Social_Security_Number AND ProcMile IS NULL"
				err = 0
				If Request("closedate") <> "" Then
					If IsDate(Request("closedate")) Then
						sundate = Request("closedate")
						sqlTBL = sqlTBL & " AND date >= #" & CDate(Request("closedate")) & "#" 
						Session("Msg") = Session("Msg") & " from " & Request("closedate")
					Else
						Err = 1
					End If
				End If
				If Request("Todate") <> "" Then
					If IsDate(Request("Todate")) Then
						satdate = Request("Todate")
						sqlTBL = sqlTBL & " AND date  <= #" & CDate(Request("Todate")) & "#" 
						Session("Msg") = Session("Msg") & " to " & Request("Todate")
					Else
						Err = 1
					End If
				End If
				If Err <> 0 Then Response.Redirect "specrep.asp?err=55" 
				'sqlTBL = sqlTBL & " ORDER BY lname, fname"
				If Request("seluri") = 0 Then
					strHEAD = "<tr bgcolor='#040C8B'><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Worker Name</b></font></td><td align='center'><font size='1' " & _
							"face='trebuchet ms' color='white'><b>File Number</b></font></td><td align='center'><font size='1' " & _
							"face='trebuchet ms' color='white'><b>RIHCC 1</b></font></td><td align='center'><font size='1' " & _
							"face='trebuchet ms' color='white'><b>RIHCC 2</b></font></td><td align='center'>" & _
							"<font size='1' face='trebuchet ms' color='white'><b>Timesheet Week</b></font></td><td align='center'>" & _
							"<font size='1' face='trebuchet ms' color='white'><b>Total Miles</b></font></td></tr>"
					Session("Msg") = Session("Msg") & " (TOTAL)"
					sqlTBL = sqlTBL & " ORDER BY worker_T.lname, worker_T.fname"
				ElseIf Request("seluri") = 1 Then
					strHEAD = "<tr bgcolor='#040C8B'><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Consumer Name</b></font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Mileage Cap</b></font></td><td align='center'><font size='1' " & _
						"face='trebuchet ms' color='white'><b>RIHCC</b></font></td><td align='center'>" & _
						"<font size='1' face='trebuchet ms' color='white'><b>Worker Name</b></font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>File Number</b></font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Timesheet Week</b></font></td><td align='center'>" & _
						"<font size='1' face='trebuchet ms' color='white'><b>Total Miles</b></font></td></tr>"
					Session("Msg") = Session("Msg") & " (DETAILED)"
					sqlTBL = sqlTBL & " ORDER BY consumer_T.lname, consumer_T.fname"
				End If
				myUri = Request("seluri")
				If Request("seluri") = 0 Then
					rsTBL.open sqlTBL, g_strCONN, 1,3
					If Not rsTBL.EOF Then
						x = 0
						Do Until rsTBL.EOF
							totmile = Z_CZero(rsTBL("mile")) + Z_CZero(rsTBL("amile"))
							If totmile <> 0 Then
								myempid = rsTBL("emp_ID")
								mycliid = rsTBL("client")
								mydate = rsTBL("date")
								mymiles = totmile
								mymilecap = rsTBL("consumer_T.milecap")
								lngIdx = SearchArrays3(myempid)
								If lngIdx < 0 Then
									ReDim Preserve tmpemp(x)
									ReDim Preserve tmpcon(x)
									ReDim Preserve tmpDates(x)
									ReDim Preserve tmpmile(x)
									ReDim Preserve tmpmilecap(x)
										
									tmpemp(x) = myempid
									tmpcon(x) = mycliid
									tmpDates(x) = mydate
									tmpmile(x) = mymiles
									tmpmilecap(x) = mymilecap
									x = x + 1
								Else
									tmpmile(lngIdx) = tmpmile(lngIdx) + mymiles
								End If
								
							End If
							rsTBL.MoveNext
						Loop
					End If
					rsTBL.Close
					Set rsTBL = Nothing
					y = 0
					Do Until y = x 
						tmpTSWk1 = Request("closedate")'tmpDates(y)
						tmpTSWk2 = Request("Todate") 'dateadd("d", 6, tmpDates(y))'Cdate(rsTBL("date")) + 6
						
						strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms' >" & GetName(tmpemp(y)) & "</font></td>" & _
							"<td align='center'><font size='1' face='trebuchet ms' >" & GetFileNum(tmpemp(y)) & _
							"</font></td><td align='center'><font size='1' face='trebuchet ms' >" & GetName3(GetPM1(tmpemp(y))) & _
							"</font></td><td align='center'><font size='1' face='trebuchet ms' >" & GetName3(GetPM2(tmpemp(y))) & _
							"</font></td><td align='center'><font size='1' face='trebuchet ms' >" & tmpTSWk1 & " - " & tmpTSWk2 & _
							"</font></td><td align='center'><font size='1' face='trebuchet ms' >" & tmpmile(y) & "</font></td></tr>"	
						y = y + 1
					Loop 
				ElseIf Request("seluri") = 1 Then
					rsTBL.open sqlTBL, g_strCONN, 1,3
					If Not rsTBL.EOF Then
						Do Until rsTBL.EOF
							totmile = Z_CZero(rsTBL("mile")) + Z_CZero(rsTBL("amile"))
							If totmile <> 0 Then
								conname = rsTBL("consumer_T.lname") & ", " & rsTBL("consumer_T.fname")
								worname = GetName(rsTBL("emp_ID"))
								APM = GetAPM2(rsTBL("client"))
								tmpTSWk1 = rsTBL("date")
								'If Request("FrmD8") <> "" Then
								'	If Cdate(tmpTSWk1) < Cdate(Request("FrmD8")) Then tmpTSWk1 = Request("FrmD8")
								'End If
								tmpTSWk2 = dateadd("d", 6, rsTBL("date"))'Cdate(rsTBL("date")) + 6
								'If Request("ToD8") <> "" Then
								'	If Cdate(tmpTSWk2) > Cdate(Request("ToD8")) Then tmpTSWk2 = Request("ToD8")
								'End If
								tmpFileNum = GetFileNum(rsTBL("emp_id"))
								strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms' >" & conname & "</font></td>" & _
									"<td align='center'><font size='1' face='trebuchet ms' >" & _
									Z_CZero(rsTBL("consumer_T.milecap")) & "</font></td><td align='center'><font size='1' face='trebuchet ms' >&nbsp;" & APM & _
									"&nbsp;</font></td><td align='center'><font size='1' face='trebuchet ms' >" & worname & _
									"</font></td><td align='center'><font size='1' face='trebuchet ms' >" & tmpFileNum & "</font></td>" & _
									"<td align='center'><font size='1' face='trebuchet ms' >" & tmpTSWk1 & " - " & tmpTSWk2 & _
									"</font></td><td align='center'><font size='1' face='trebuchet ms' >" & totmile & "</font></td></tr>"	
							End If
							rsTBL.MoveNext
						Loop
					End If
					rsTBL.Close
					Set rsTBL = Nothing
				End If
				
				'not w/n pay period
				If Request("closedate") <> "" Then
					Set rsMile = Server.CreateObject("ADODB.RecordSet")
					sqlMile = "SELECT * FROM Tsheets_T, consumer_T, worker_T WHERE Medicaid_Number = client AND emp_ID = Worker_T.Social_Security_Number " & _
						 "AND date < #" & CDate(Request("closedate")) & "# AND ProcMile IS NULL"
					If Request("seluri") = 0 Then
						sqlMile = sqlMile & " ORDER BY worker_T.lname, worker_T.fname"
						strHEAD2 = "<tr bgcolor='#040C8B'><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Worker Name</b></font></td><td align='center'><font size='1' " & _
							"face='trebuchet ms' color='white'><b>File Number</b></font></td><td align='center'><font size='1' " & _
							"face='trebuchet ms' color='white'><b>RIHCC 1</b></font></td><td align='center'><font size='1' " & _
							"face='trebuchet ms' color='white'><b>RIHCC 2</b></font></td><td align='center'>" & _
							"<font size='1' face='trebuchet ms' color='white'><b>Timesheet Week</b></font></td><td align='center'>" & _
							"<font size='1' face='trebuchet ms' color='white'><b>Total Miles</b></font></td></tr>"
					ElseIf Request("seluri") = 1 Then
						sqlMile = sqlMile & " ORDER BY consumer_T.lname, consumer_T.fname"
						strHEAD2 = "<tr bgcolor='#040C8B'><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Consumer Name</b></font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Mileage Cap</b></font></td><td align='center'><font size='1' " & _
						"face='trebuchet ms' color='white'><b>RIHCC</b></font></td><td align='center'>" & _
						"<font size='1' face='trebuchet ms' color='white'><b>Worker Name</b></font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>File Number</b></font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Timesheet Week</b></font></td><td align='center'>" & _
						"<font size='1' face='trebuchet ms' color='white'><b>Total Miles</b></font></td></tr>"
					End If
					
					If Request("seluri") = 0 Then
						rsMile.open sqlMile, g_strCONN, 1,3
						If Not rsMile.EOF Then
							x = 0
							Do Until rsMile.EOF
								totmile = Z_CZero(rsMile("mile")) + Z_CZero(rsMile("amile"))
								If totmile <> 0 Then
									myempid = rsMile("emp_ID")
									mycliid = rsMile("client")
									mydate = rsMile("date")
									mymiles = totmile
									mymilecap = rsMile("consumer_T.milecap")
									lngIdx = SearchArrays3(myempid)
									If lngIdx < 0 Then
										ReDim Preserve tmpemp(x)
										ReDim Preserve tmpcon(x)
										ReDim Preserve tmpDates(x)
										ReDim Preserve tmpmile(x)
										ReDim Preserve tmpmilecap(x)
											
										tmpemp(x) = myempid
										tmpcon(x) = mycliid
										tmpDates(x) = mydate
										tmpmile(x) = mymiles
										tmpmilecap(x) = mymilecap
										x = x + 1
									Else
										tmpmile(lngIdx) = tmpmile(lngIdx) + mymiles
									End If
									
								End If
								rsMile.MoveNext
							Loop
						End If
						rsMile.Close
						Set rsMile = Nothing
						y = 0
						Do Until y = x 
							tmpTSWk1 = tmpDates(y)
							tmpTSWk2 = dateadd("d", 6, tmpDates(y))'Cdate(rsTBL("date")) + 6
							
							strBODY2 = strBODY2 & "<tr><td align='center'><font size='1' face='trebuchet ms' >" & GetName(tmpemp(y)) & "</font></td>" & _
								"<td align='center'><font size='1' face='trebuchet ms' >" & GetFileNum(tmpemp(y)) & _
								"</font></td><td align='center'><font size='1' face='trebuchet ms' >" & GetName3(GetPM1(tmpemp(y))) & _
								"</font></td><td align='center'><font size='1' face='trebuchet ms' >" & GetName3(GetPM2(tmpemp(y))) & _
								"</font></td><td align='center'><font size='1' face='trebuchet ms' >" & tmpTSWk1 & " - " & tmpTSWk2 & _
								"</font></td><td align='center'><font size='1' face='trebuchet ms' >" & tmpmile(y) & "</font></td></tr>"	
							y = y + 1
						Loop 
					ElseIf Request("seluri") = 1 Then
						rsMile.open sqlMile, g_strCONN, 1,3
						If Not rsMile.EOF Then
							Do Until rsMile.EOF
								totmile = Z_CZero(rsMile("mile")) + Z_CZero(rsMile("amile"))
								If totmile <> 0 Then
									conname = rsMile("consumer_T.lname") & ", " & rsMile("consumer_T.fname")
									worname = GetName(rsMile("emp_ID"))
									APM = GetAPM2(rsMile("client"))
									tmpTSWk1 = rsMile("date")
									'If Request("FrmD8") <> "" Then
									'	If Cdate(tmpTSWk1) < Cdate(Request("FrmD8")) Then tmpTSWk1 = Request("FrmD8")
									'End If
									tmpTSWk2 = dateadd("d", 6, rsMile("date"))'Cdate(rsTBL("date")) + 6
									'If Request("ToD8") <> "" Then
									'	If Cdate(tmpTSWk2) > Cdate(Request("ToD8")) Then tmpTSWk2 = Request("ToD8")
									'End If
									tmpFileNum = GetFileNum(rsMile("emp_id"))
									strBODY2 = strBODY2 & "<tr><td align='center'><font size='1' face='trebuchet ms' >" & conname & "</font></td>" & _
										"<td align='center'><font size='1' face='trebuchet ms' >" & _
										Z_CZero(rsMile("consumer_T.milecap")) & "</font></td><td align='center'><font size='1' face='trebuchet ms' >&nbsp;" & APM & _
										"&nbsp;</font></td><td align='center'><font size='1' face='trebuchet ms' >" & worname & _
										"</font></td><td align='center'><font size='1' face='trebuchet ms' >" & tmpFileNum & "</font></td>" & _
										"<td align='center'><font size='1' face='trebuchet ms' >" & tmpTSWk1 & " - " & tmpTSWk2 & _
										"</font></td><td align='center'><font size='1' face='trebuchet ms' >" & totmile & "</font></td></tr>"	
								End If
								rsMile.MoveNext
							Loop
						End If
						rsMile.Close
						Set rsMile = Nothing
					End If
				End If
			ElseIf Request("SelRep") = 56 Then
				Session("MSG") = "Private Pay Consumers"
				strHEAD = "<tr bgcolor='#040C8B'><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Name</b>" & _
					"</font></td>" & _
					"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>ID</b></font></td>" & _
					"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Contract Hours</b></font></td></tr>"
					
				Set rsTBL = Server.CreateObject("ADODB.RecordSet")	
				sqlTBL = "SELECT * FROM Consumer_T, c_status_T WHERE (consumer_t.medicaid_number LIKE '%Private Pay%' OR (code = 'P' OR code = 'C' OR code = 'A')) AND consumer_t.medicaid_number = c_status_t.medicaid_number " & _
					"AND Active = true ORDER BY lname, fname"
				rsTBL.Open sqlTBL, g_strCONN, 1,3 
				If Not rsTBL.EOF Then
					tmpcontract = 0
					Do Until rsTBL.EOF
						strID = rsTBL("consumer_t.medicaid_number")
						If Not IsNull(rsTBL("lname")) Then
							strName = Replace(rsTBL("lname"),",","") & ", " & rsTBL("fname")
						Else
							strName = rsTBL("lname") & ", " & rsTBL("fname")
						End If
						strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & GetName2(strID) & _
						"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & strID & _
						"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & Z_Czero(rsTBL("contract")) & _
						"</font></td></tr>"
						tmpcontract = tmpcontract + Z_Czero(rsTBL("contract"))
						rsTBL.MoveNext
					Loop
					strBODY = strBODY & "<tr><td>&nbsp;</td><td align='right'><font size='1' face='trebuchet ms'>TOTAL:</font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms'>" & tmpcontract & "</font></td></tr>"
				End If
				rsTBL.Close
				Set rsTBL = Nothing	
			ElseIf Request("SelRep") = 57 Then
				Session("MSG") = "PCSP Workers Contact Info"
				strHEAD = "<tr bgcolor='#040C8B'><td align='center'><font size='1' face='trebuchet ms' color='white'><b>First Name</b>" & _
					"</font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Last Name</b>" & _
					"</font></td>" & _
					"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Phone</b></font></td>" & _
					"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Mobile</b></font></td>" & _
					"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Email</b></font></td></tr>"
					
				Set rsTBL = Server.CreateObject("ADODB.RecordSet")	
				sqlTBL = "SELECT * FROM Worker_T WHERE status = 'Active' ORDER BY lname, fname"
				rsTBL.Open sqlTBL, g_strCONN, 1,3 
				If Not rsTBL.EOF Then
					Do Until rsTBL.EOF
						
						strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("lname") & _
						"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("fname") & _
						"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("PhoneNo") & _
						"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("CellNo") & _
						"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("eMail") & _
						"</font></td></tr>"
						rsTBL.MoveNext
					Loop
				End If
				rsTBL.Close
				Set rsTBL = Nothing	
			ElseIf Request("SelRep") = 58 Then
				Session("MSG") = "PCSP Workers Driver Info"
				strHEAD = "<tr bgcolor='#040C8B'><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Last Name</b>" & _
					"</font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>First Name</b>" & _
					"</font></td>" & _
					"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Mailing Address</b></font></td>" & _
					"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Phone</b></font></td>" & _
					"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Mobile</b></font></td>" & _
					"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Email</b></font></td>" & _
					"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>RIHCC 1</b></font></td>" & _
					"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>RIHCC 2</b></font></td></tr>"
					
				Set rsTBL = Server.CreateObject("ADODB.RecordSet")	
				sqlTBL = "SELECT * FROM Worker_T WHERE status = 'Active' AND Driver = True ORDER BY lname, fname"
				rsTBL.Open sqlTBL, g_strCONN, 1,3 
				If Not rsTBL.EOF Then
					Do Until rsTBL.EOF
						maddress = rsTBL("maddress") & ", " &  rsTBL("mcity") & ", " &  rsTBL("mState") & ", " &  rsTBL("mzip")
						strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("lname") & _
						"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("fname") & _
						"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & maddress & _
						"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("PhoneNo") & _
						"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("CellNo") & _
						"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("eMail") & _
						"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & GetName3(rsTBL("pm1")) & _
						"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & GetName3(rsTBL("pm2")) & _
						"</font></td></tr>"
						rsTBL.MoveNext
					Loop
				End If
				rsTBL.Close
				Set rsTBL = Nothing	
			ElseIf Request("SelRep") = 59 Then
				Session("MSG") = "Total Hours for PCSP Worker report(detailed) "
				sqlProc = "SELECT * FROM [Tsheets_t], [worker_t] WHERE ext = false AND emp_id = Social_Security_Number"
				Err = 0
				If Request("FrmD8") <> "" Then
					If IsDate(Request("FrmD8")) Then
						sqlProc = sqlProc & " AND date >= #" & dateAdd("d", -7, Request("FrmD8")) & "# "
						Session("Msg") = Session("Msg") & " from " & Request("FrmD8")
					Else
						Err = 1
					End If
				End If
				If Request("ToD8") <> "" Then
					If IsDate(Request("ToD8")) Then
						sqlProc = sqlProc & " AND date  <= #" & dateAdd("d", 7, Request("ToD8")) & "#" 
						Session("Msg") = Session("Msg") & " to " & Request("ToD8")
					Else
						Err = 1
					End If
				End If
				Session("Msg") = Session("Msg") & "<br> * Extended Hours"
				If Err <> 0 Then Response.Redirect "specrep.asp?err=59" 
				strHEAD = "<tr bgcolor='#040C8B'><td align='center'><font size='1' face='trebuchet ms' color='white'><b>File Number</b></font></td>" & _
					"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>PCSP Worker</b></font></td><td align='center'>" & _
					"<font size='1' face='trebuchet ms' color='white'><b>Medicaid</b></font></td><td align='center'><font size='1' " & _
					"face='trebuchet ms' color='white'><b>Code</b></font></td><td align='center'><font size='1' " & _
					"face='trebuchet ms' color='white'><b>Consumer Name</b></font></td><td align='center'><font size='1' " & _
					"face='trebuchet ms' color='white'><b>Date</b></font></td><td align='center'><font size='1' " & _
					"face='trebuchet ms' color='white'><b>Regular Hours</b></font></td><td align='center'><font size='1' " & _
					"face='trebuchet ms' color='white'><b>Extended Hours</b></font></td><td align='center'>" & _
					"<font size='1' face='trebuchet ms' color='white'><b>Total Hours</b></font></td></tr>"
				Set rsProc = Server.CreateObject("ADODB.RecordSet")
				
				sqlProc = sqlProc & " ORDER BY lname ASC, fname ASC, date DESC, ID DESC"
				'response.write sqlProc
				rsProc.Open sqlProc, g_strCONN, 3, 1
				If Not rsProc.EOF Then
					Do Until rsProc.EOF
						tmpName = GetName(rsProc("emp_id"))
						tmpName2 = GetName2(rsProc("client"))
						myCode = rsProc("client")
						tmphrsMon = 0
            tmphrsTue = 0
            tmphrsWed = 0
            tmphrsThu = 0
            tmphrsFri = 0
            tmphrsSat = 0
            tmphrsSun = 0
						'regular hours
						THours = rsProc("mon") + rsProc("tue") + rsProc("wed") + rsProc("thu") + rsProc("fri") + rsProc("sat") + rsProc("sun")
						If THours <> 0 Then  
							tmphrsMon = ValidDate(Request("FrmD8"), Request("ToD8"), rsProc("date"), rsProc("mon"), "MON")
	            tmphrsTue = ValidDate(Request("FrmD8"), Request("ToD8"), rsProc("date"), rsProc("tue"), "TUE")
	            tmphrsWed = ValidDate(Request("FrmD8"), Request("ToD8"), rsProc("date"), rsProc("wed"), "WED")
	            tmphrsThu = ValidDate(Request("FrmD8"), Request("ToD8"), rsProc("date"), rsProc("thu"), "THU")
	            tmphrsFri = ValidDate(Request("FrmD8"), Request("ToD8"), rsProc("date"), rsProc("fri"), "FRI")
	            tmphrsSat = ValidDate(Request("FrmD8"), Request("ToD8"), rsProc("date"), rsProc("sat"), "SAT")
	            tmphrsSun = ValidDate(Request("FrmD8"), Request("ToD8"), rsProc("date"), rsProc("sun"), "SUN")
	          End If
	          'ext hours
	          tmphrsMonx = 0
            tmphrsTuex = 0
            tmphrsWedx = 0
            tmphrsThux = 0
            tmphrsFrix = 0
            tmphrsSatx = 0
            tmphrsSunx = 0
	          
	          Set rsext = Server.CreateObject("ADODB.RecordSet")
	          sqlext = "SELECT * FROM Tsheets_t WHERE date = #" & rsProc("date") & "# AND emp_ID = '" & rsProc("emp_id") & "' " & _
	          	"AND Client = '" & rsProc("client") & "' AND ext = true AND timestamp = #" & rsProc("timestamp") & "#"
	          rsext.Open sqlext, g_strCONN, 3, 1
	          
	          If Not rsext.EOF Then
	        		THoursx = rsext("mon") + rsext("tue") + rsext("wed") + rsext("thu") + rsext("fri") + rsext("sat") + rsext("sun")
							If THoursx <> 0 Then 
								tmphrsMonx = ValidDate(Request("FrmD8"), Request("ToD8"), rsext("date"), rsext("mon"), "MON")
		            tmphrsTuex = ValidDate(Request("FrmD8"), Request("ToD8"), rsext("date"), rsext("tue"), "TUE")
		            tmphrsWedx = ValidDate(Request("FrmD8"), Request("ToD8"), rsext("date"), rsext("wed"), "WED")
		            tmphrsThux = ValidDate(Request("FrmD8"), Request("ToD8"), rsext("date"), rsext("thu"), "THU")
		            tmphrsFrix = ValidDate(Request("FrmD8"), Request("ToD8"), rsext("date"), rsext("fri"), "FRI")
		            tmphrsSatx = ValidDate(Request("FrmD8"), Request("ToD8"), rsext("date"), rsext("sat"), "SAT")
		            tmphrsSunx = ValidDate(Request("FrmD8"), Request("ToD8"), rsext("date"), rsext("sun"), "SUN")
		          End If
		        Else
		        	Set fso = CreateObject("Scripting.FileSystemObject")
							Set ALog = fso.OpenTextFile(AdminLog, 8, True)
							Alog.WriteLine Now & ":: Extended Hours NOT FOUND (ID: " & rsProc("id") & ")" & vbCrLf
							Set Alog = Nothing
							Set fso = Nothing 
	        	End If
	        	rsext.Close
	        	Set rsext = Nothing
	        	If Not (tmphrsSun = 0 And tmphrsSunx = 0) Then
	          	mySun = GetDate(rsProc("date"), "SUN")
		          strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & rsProc("FileNum") & _
								"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & tmpName & _
								"</font></td>"
								If Session("lngType") = 1 Or Session("lngType") = 2 Then 
									strBODY = strBODY & "<td align='center'><font size='1' face='trebuchet ms'>" & rsProc("client") & _
									"</font></td>"
								Else
									strBODY = strBODY & "<Td>&nbsp;</td>"
								End If 
								strBODY = strBODY & "<td align='center'><font size='1' face='trebuchet ms'>" & GetCode(rsProc("client")) & _
								"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & tmpName2 & _
								"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & mySun & _
								"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & tmphrsSun & _
								"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & tmphrsSunx & _
								"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & tmphrsSun + tmphrsSunx & _
								"</font></td></tr>"
						End If
	          If Not (tmphrsMon = 0 And tmphrsMonx = 0) Then
	          	myMon = GetDate(rsProc("date"), "MON")
		          strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & rsProc("FileNum") & _
								"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & tmpName & _
								"</font></td>"
								If Session("lngType") = 1 Or Session("lngType") = 2 Then 
									strBODY = strBODY & "<td align='center'><font size='1' face='trebuchet ms'>" & rsProc("client") & _
									"</font></td>"
								Else
									strBODY = strBODY & "<Td>&nbsp;</td>"
								End If 
								strBODY = strBODY & "<td align='center'><font size='1' face='trebuchet ms'>" & GetCode(rsProc("client")) & _
								"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & tmpName2 & _
								"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & myMon & _
								"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & tmphrsMon & _
								"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & tmphrsMonx & _
								"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & tmphrsMon + tmphrsMonx & _
								"</font></td></tr>"
						End If
						If Not (tmphrsTue = 0 And tmphrsTuex = 0) Then
	          	myTue = GetDate(rsProc("date"), "TUE")
		          strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & rsProc("FileNum") & _
								"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & tmpName & _
								"</font></td>"
								If Session("lngType") = 1 Or Session("lngType") = 2 Then 
									strBODY = strBODY & "<td align='center'><font size='1' face='trebuchet ms'>" & rsProc("client") & _
									"</font></td>"
								Else
									strBODY = strBODY & "<Td>&nbsp;</td>"
								End If 
								strBODY = strBODY & "<td align='center'><font size='1' face='trebuchet ms'>" & GetCode(rsProc("client")) & _
								"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & tmpName2 & _
								"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & myTue & _
								"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & tmphrsTue & _
								"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & tmphrsTuex & _
								"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & tmphrsTue + tmphrsTuex & _
								"</font></td></tr>"
						End If
						If Not (tmphrsWed = 0 And tmphrsWedx = 0) Then
	          	myWed = GetDate(rsProc("date"), "WED")
		          strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & rsProc("FileNum") & _
								"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & tmpName & _
								"</font></td>"
								If Session("lngType") = 1 Or Session("lngType") = 2 Then 
									strBODY = strBODY & "<td align='center'><font size='1' face='trebuchet ms'>" & rsProc("client") & _
									"</font></td>"
								Else
									strBODY = strBODY & "<Td>&nbsp;</td>"
								End If 
								strBODY = strBODY & "<td align='center'><font size='1' face='trebuchet ms'>" & GetCode(rsProc("client")) & _
								"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & tmpName2 & _
								"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & myWed & _
								"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & tmphrsWed & _
								"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & tmphrsWedx & _
								"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & tmphrsWed + tmphrsWedx & _
								"</font></td></tr>"
						End If
						If Not (tmphrsThu = 0 And tmphrsThux = 0) Then
	          	myThu = GetDate(rsProc("date"), "THU")
		          strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & rsProc("FileNum") & _
								"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & tmpName & _
								"</font></td>"
								If Session("lngType") = 1 Or Session("lngType") = 2 Then 
									strBODY = strBODY & "<td align='center'><font size='1' face='trebuchet ms'>" & rsProc("client") & _
									"</font></td>"
								Else
									strBODY = strBODY & "<Td>&nbsp;</td>"
								End If 
								strBODY = strBODY & "<td align='center'><font size='1' face='trebuchet ms'>" & GetCode(rsProc("client")) & _
								"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & tmpName2 & _
								"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & myThu & _
								"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & tmphrsThu & _
								"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & tmphrsThux & _
								"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & tmphrsThu + tmphrsThux & _
								"</font></td></tr>"
						End If
						If Not (tmphrsFri = 0 And tmphrsFrix = 0) Then
	          	myFri = GetDate(rsProc("date"), "FRI")
		          strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & rsProc("FileNum") & _
								"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & tmpName & _
								"</font></td>"
								If Session("lngType") = 1 Or Session("lngType") = 2 Then 
									strBODY = strBODY & "<td align='center'><font size='1' face='trebuchet ms'>" & rsProc("client") & _
									"</font></td>"
								Else
									strBODY = strBODY & "<Td>&nbsp;</td>"
								End If 
								strBODY = strBODY & "<td align='center'><font size='1' face='trebuchet ms'>" & GetCode(rsProc("client")) & _
								"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & tmpName2 & _
								"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & myFri & _
								"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & tmphrsFri & _
								"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & tmphrsFrix & _
								"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & tmphrsFri + tmphrsFrix & _
								"</font></td></tr>"
						End If
						If Not (tmphrsSat = 0 And tmphrsSatx = 0) Then
	          	mySat = GetDate(rsProc("date"), "SAT")
		          strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & rsProc("FileNum") & _
								"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & tmpName & _
								"</font></td>"
								If Session("lngType") = 1 Or Session("lngType") = 2 Then 
									strBODY = strBODY & "<td align='center'><font size='1' face='trebuchet ms'>" & rsProc("client") & _
									"</font></td>"
								Else
									strBODY = strBODY & "<Td>&nbsp;</td>"
								End If 
								strBODY = strBODY & "<td align='center'><font size='1' face='trebuchet ms'>" & GetCode(rsProc("client")) & _
								"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & tmpName2 & _
								"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & mySat & _
								"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & tmphrsSat & _
								"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & tmphrsSatx & _
								"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & tmphrsSat + tmphrsSatx & _
								"</font></td></tr>"
						End If
						rsProc.MoveNext
					Loop
				Else
					Session("MSG") = Session("MSG") & "<br> No records found."	
				End If
				rsProc.Close
				Set rsProc = Nothing
			ElseIf Request("SelRep") = 60 Then
				Session("MSG") = "Total Hours of "
				sqlProc = "SELECT * FROM [Tsheets_t] "
				If Request("seltype3") = 1 Then
					Session("MSG") = Session("MSG") & "PCSP Workers "
					sqlProc = sqlProc & ", [Worker_T] WHERE emp_id = Social_Security_Number"
				Else
					Session("MSG") = Session("MSG") & "Consumers "
					sqlProc = sqlProc & ", [Consumer_T] WHERE client = Medicaid_Number"
				End If
				Err = 0
				If Request("FrmD8") <> "" Then
					If IsDate(Request("FrmD8")) Then
						sqlProc = sqlProc & " AND date >= #" & dateAdd("d", -7, Request("FrmD8")) & "# "
						Session("Msg") = Session("Msg") & " from " & Request("FrmD8")
					Else
						Err = 1
					End If
				End If
				If Request("ToD8") <> "" Then
					If IsDate(Request("ToD8")) Then
						sqlProc = sqlProc & " AND date  <= #" & dateAdd("d", 7, Request("ToD8")) & "#" 
						Session("Msg") = Session("Msg") & " to " & Request("ToD8")
					Else
						Err = 1
					End If
				End If
				If Err <> 0 Then Response.Redirect "specrep.asp?err=60" 
				If Request("seltype3") = 1 Then
					strHEAD = "<tr bgcolor='#040C8B'><td align='center'><font size='1' face='trebuchet ms' color='white'><b>File Number</b></font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>PCSP Worker</b></font></td><td align='center'><font size='1' " & _
						"face='trebuchet ms' color='white'><b>Regular Hours</b></font></td><td align='center'><font size='1' " & _
						"face='trebuchet ms' color='white'><b>Extended Hours</b></font></td><td align='center'>" & _
						"<font size='1' face='trebuchet ms' color='white'><b>Total Hours</b></font></td></tr>"
					sqlProc = sqlProc & " ORDER BY lname, fname"
				Else
					strHEAD = "<tr bgcolor='#040C8B'><td align='center'><font size='1' " & _
						"face='trebuchet ms' color='white'><b>Consumer Name</b></font></td><td align='center'><font size='1' " & _
						"face='trebuchet ms' color='white'><b>Code</b></font></td><td align='center'><font size='1' " & _
						"face='trebuchet ms' color='white'><b>Regular Hours</b></font></td><td align='center'><font size='1' " & _
						"face='trebuchet ms' color='white'><b>Extended Hours</b></font></td><td align='center'>" & _
						"<font size='1' face='trebuchet ms' color='white'><b>Total Hours</b></font></td></tr>"
					sqlProc = sqlProc & " ORDER BY code, lname, fname"
				End If	
				Set rsProc = Server.CreateObject("ADODB.RecordSet")	
				rsProc.Open sqlProc, g_strCONN, 3, 1
				x = 0
				Do Until rsProc.EOF
					If Request("seltype3") = 1 Then
						strID = rsProc("emp_id")
					Else
						strID = rsProc("client")
					End If
					
					'GET HOURS dblHours
					tmphrsMon = 0
          tmphrsTue = 0
          tmphrsWed = 0
          tmphrsThu = 0
          tmphrsFri = 0
          tmphrsSat = 0
          tmphrsSun = 0
					THours = rsProc("mon") + rsProc("tue") + rsProc("wed") + rsProc("thu") + rsProc("fri") + rsProc("sat") + rsProc("sun")
					If THours <> 0 Then  
						tmphrsMon = ValidDate(Request("FrmD8"), Request("ToD8"), rsProc("date"), rsProc("mon"), "MON")
            tmphrsTue = ValidDate(Request("FrmD8"), Request("ToD8"), rsProc("date"), rsProc("tue"), "TUE")
            tmphrsWed = ValidDate(Request("FrmD8"), Request("ToD8"), rsProc("date"), rsProc("wed"), "WED")
            tmphrsThu = ValidDate(Request("FrmD8"), Request("ToD8"), rsProc("date"), rsProc("thu"), "THU")
            tmphrsFri = ValidDate(Request("FrmD8"), Request("ToD8"), rsProc("date"), rsProc("fri"), "FRI")
            tmphrsSat = ValidDate(Request("FrmD8"), Request("ToD8"), rsProc("date"), rsProc("sat"), "SAT")
            tmphrsSun = ValidDate(Request("FrmD8"), Request("ToD8"), rsProc("date"), rsProc("sun"), "SUN")
          End If
					dblHours = tmphrsMon + tmphrsTue + tmphrsWed + tmphrsThu + tmphrsFri + tmphrsSat + tmphrsSun
					
					lngIdx = SearchArrays60(strID, tmpID)
					If lngIdx < 0 Then
						ReDim Preserve tmpID(x)
						ReDim Preserve tmpHrs(x)
						ReDim Preserve tmpHrsExt(x)
						
						tmpID(x) = strID
						If Not rsProc("Ext") Then
							tmpHrs(x) = dblHours
						Else
							tmpHrsExt(x) = dblHours
						End If
						x = x + 1
					Else
						If Not rsProc("Ext") Then
							tmpHrs(lngIdx) = tmpHrs(lngIdx) + dblHours
						Else
							tmpHrsExt(lngIdx) = tmpHrsExt(lngIdx) + dblHours
						End If
					End If
						
					rsProc.MoveNext
				Loop
				rsProc.Close
				Set rsProc = Nothing
				y = 0
				a = 0
				c = 0
				m = 0
				p = 0
				ahrs = 0
				chrs = 0
				mhrs = 0
				phrs = 0
				Do Until y = x 
					myHrs = Z_CZero(tmpHrs(y) + tmpHrsExt(y))
					If myHrs > 0 Then
						If Request("seltype3") = 1 Then
							tmpName = GetName(tmpID(y))
							strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & GetFileNum(tmpID(y)) & _
								"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & tmpName & _
								"</font></td>"
						Else
							tmpName = GetName2(tmpID(y))
							strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & tmpName & _
								"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & GetCode(tmpID(y)) & _
								"</font></td>"
							
							If GetCode(tmpID(y)) = "A" Then
								a = a + 1
								ahrs = ahrs + myHrs
							ElseIf GetCode(tmpID(y)) = "C" Then
								c = c + 1 
								chrs = chrs + myHrs
							ElseIf GetCode(tmpID(y)) = "M" Then
								m = m + 1
								mhrs = mhrs + myHrs
							ElseIf GetCode(tmpID(y)) = "P" Then
								p = p + 1
								phrs = phrs + myHrs
							End If 	 		
						End If
						strBODY = strBODY &	"<td align='center'><font size='1' face='trebuchet ms'>" & Z_CZero(tmpHrs(y)) & _
							"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & Z_CZero(tmpHrsExt(y)) & _
							"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & Z_CZero(tmpHrs(y) + tmpHrsExt(y)) & _
							"</font></td></tr>"
						End If
					y = y + 1
				Loop 
				ctotal = 0
				hrstotal = 0
				If Request("seltype3") <> 1 Then
					ctotal = a + c + m + p
					hrstotal = ahrs + chrs + mhrs + phrs
					strBODY = strBODY &	"<tr><td align='center' colspan='5'><font size='1' face='trebuchet ms'>TOTALS" & _
							"</font></td></tr><tr><td align='center'><font size='1' face='trebuchet ms'>CODE</font></td>" & _
							"<td align='center' colspan='2'><font size='1' face='trebuchet ms'>COUNT</font></td><td align='center' colspan='2'><font size='1' face='trebuchet ms'>HOURS</font></td></tr>" & _
							"<tr><td align='center'><font size='1' face='trebuchet ms'>A</font></td><td align='center' colspan='2'><font size='1' face='trebuchet ms'>" & a & "</font></td>" & _
							"<td align='center' colspan='2'><font size='1' face='trebuchet ms'>" & ahrs & "</font></td></tr>" & _
							"<tr><td align='center'><font size='1' face='trebuchet ms'>C</font></td><td align='center' colspan='2'><font size='1' face='trebuchet ms'>" & c & "</font></td>" & _
							"<td align='center' colspan='2'><font size='1' face='trebuchet ms'>" & chrs & "</font></td></tr>" & _
							"<tr><td align='center'><font size='1' face='trebuchet ms'>M</font></td><td align='center' colspan='2'><font size='1' face='trebuchet ms'>" & m & "</font></td>" & _
							"<td align='center' colspan='2'><font size='1' face='trebuchet ms'>" & mhrs & "</font></td></tr>" & _
							"<tr><td align='center'><font size='1' face='trebuchet ms'>P</font></td><td align='center' colspan='2'><font size='1' face='trebuchet ms'>" & p & "</font></td>" & _
							"<td align='center' colspan='2'><font size='1' face='trebuchet ms'>" & phrs & "</font></td></tr>" & _
							"<tr><td align='center'><font size='1' face='trebuchet ms'>Total</font></td><td align='center' colspan='2'><font size='1' face='trebuchet ms'>" & ctotal & "</font></td>" & _
							"<td align='center' colspan='2'><font size='1' face='trebuchet ms'>" & hrstotal & "</font></td></tr>"
							
				End If
			ElseIf Request("SelRep") = 61 Then
				Session("MSG") = "Consumers with PCSP report"
				sqlProc = "SELECT * FROM Consumer_T , c_Status_T WHERE Consumer_T.Medicaid_Number = c_Status_T.Medicaid_Number AND Active = True ORDER BY lname, fname"
				strHEAD = "<tr bgcolor='#040C8B'><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Consumer</b></font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>RIHCC</b></font></td><td align='center'><font size='1' " & _
						"face='trebuchet ms' color='white'><b>Max Hours</b></font></td><td align='center'><font size='1' " & _
						"face='trebuchet ms' color='white'><b>PCSP WORKER</b></font></td></tr>"
				Set rsProc = Server.CreateObject("ADODB.RecordSet")	
				rsProc.Open sqlProc, g_strCONN, 3, 1
				Do Until rsProc.EOF
					tmpname = rsProc("lname") & ", " & rsProc("fname")
					tmpRCC = GetAPM2(rsProc("Consumer_T.medicaid_number"))
					
					strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & tmpName & _
								"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & tmpRCC & _
								"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & rsProc("MaxHrs") & _
								"</font></td><td align='center'><font size='1' face='trebuchet ms'>------</font></td></tr>"
								
					'GET WORKERS
					Set rsWork = Server.CreateObject("ADODB.RecordSet")
					sqlwork = "SELECT * FROM Conwork_T WHERE CID = '" & rsProc("Consumer_T.medicaid_number") & "'"
					rsWork.Open sqlwork, g_strCONN, 3, 1
					Do Until rsWork.EOF
						strBODY = strBODY & "<tr><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td align='center'><font size='1' face='trebuchet ms'>" & GetWorkName(rsWork("WID")) & "</font></td></tr>"
						rsWork.MoveNext
					Loop
					rsWork.Close
					Set rsWork = Nothing	
					rsProc.MoveNext
				Loop
				rsProc.Close
				Set rsProc = Nothing
			ElseIf Request("SelRep") = 62 Then

				
				Session("MSG") = "Private Pay Eligible Worker"
				sqlProc = "SELECT * FROM worker_T WHERE status = 'Active' AND privatepay = true ORDER BY lname, fname"
				strHEAD = "<tr bgcolor='#040C8B'><td align='center'><font size='1' face='trebuchet ms' color='white'><b>PCSP Worker</b></font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Address</b></font></td><td align='center'><font size='1' " & _
						"face='trebuchet ms' color='white'><b>Phone Number</b></font></td><td align='center'><font size='1' " & _
						"face='trebuchet ms' color='white'><b>RHICC1</b></font></td></tr>"
				Set rsProc = Server.CreateObject("ADODB.RecordSet")	
				rsProc.Open sqlProc, g_strCONN, 3, 1
				Do Until rsProc.EOF
					tmpname = rsProc("lname") & ", " & rsProc("fname")
					tmpRCC = GetName3(z_czero(rsProc("pm1")))
					tmpAddr = rsProc("maddress") & ", " & rsProc("mcity") & ", " & rsProc("mstate") & ", " & rsProc("mzip")
					
					strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & tmpName & _
								"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & tmpAddr & _
								"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & rsProc("PhoneNo") & _
								"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & tmpRCC & "</font></td></tr>"
								
					rsProc.MoveNext
				Loop
				rsProc.Close
				Set rsProc = Nothing
			ElseIf Request("SelRep") = 63 Then

				
				Session("MSG") = "PCSP Worker Training Logs"
				sqlProc = "SELECT * FROM worker_T, w_log_T WHERE ssn = social_security_number AND status = 'Active' AND not isnull(train) ORDER BY lname, fname"
				strHEAD = "<tr bgcolor='#040C8B'><td align='center'><font size='1' face='trebuchet ms' color='white'><b>PCSP Worker</b></font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Training</b></font></td><td align='center'><font size='1' " & _
						"face='trebuchet ms' color='white'><b>Hours</b></font></td><td align='center'><font size='1' " & _
						"face='trebuchet ms' color='white'><b>Comments</b></font></td></tr>"
				Set rsProc = Server.CreateObject("ADODB.RecordSet")	
				rsProc.Open sqlProc, g_strCONN, 3, 1
				Do Until rsProc.EOF
					tmpname = rsProc("lname") & ", " & rsProc("fname")
					
					strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & tmpName & _
								"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & rsProc("train") & _
								"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & rsProc("thrs") & _
								"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & rsProc("tcom") & "</font></td></tr>"
								
					rsProc.MoveNext
				Loop
				rsProc.Close
				Set rsProc = Nothing
			ElseIf Request("SelRep") = 64 Then
				Session("MSG") = "PCSP Worker Overage Hours from " & Request("closedate") & " to " & Request("Todate")
				strHEAD = "<tr bgcolor='#040C8B'><td align='center'><font size='1' face='trebuchet ms' color='white'><b>PCSP Worker</b>" & _
					"</font></td>" & _
					"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Consumer</b></font></td>" & _
					"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Timesheet Week</b></font></td>" & _
					"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Max Hours</b></font></td>" & _
					"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Hours</b></font></td>" & _
					"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Overage Hours</b></font></td>" & _
					"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>RHICC</b></font></td></tr>" 
				sqlProc = "SELECT * FROM tsheets_t, worker_T, consumer_T WHERE emp_id = worker_T.Social_Security_Number AND client = consumer_t.medicaid_number AND date >=#" & Request("closedate") & "# AND date <= #" & _
					Request("Todate") & "# ORDER BY pm1, consumer_t.lname, consumer_t.fname, emp_id"
					
				'response.write sqlTBL
				Set rsTBL = Server.CreateObject("ADODB.RecordSet")
				rsTBL.Open sqlProc, g_strCONN, 1, 3
				x = 0
				Do Until rsTBL.EOF
					strDate = rsTBL("date")
					strEID = rsTBL("client")
					strWorID = rsTBL("emp_id")
					dblHours = rsTBL("mon") + rsTBL("tue") + rsTBL("wed") + rsTBL("thu") + rsTBL("fri") + rsTBL("sat") + rsTBL("sun")
					lngIdx = SearchArrays4(strDate, strEID, strWorID, tmpDates, tmpEmpID, tmpWorID)
					If lngIdx < 0 Then
						ReDim Preserve tmpDates(x)
						ReDim Preserve tmpWorID(x)
						ReDim Preserve tmpEmpID(x)
						ReDim Preserve tmpHrs(x)
						
						tmpDates(x) = strDate
						tmpEmpID(x) = strEID
						tmpWorID(x) = strWorID
						tmpHrs(x) = dblHours
						x = x + 1
					Else
						tmpHrs(lngIdx) = tmpHrs(lngIdx) + dblHours
					End If
					rsTBL.MoveNext 
				Loop
				rsTBL.Close
				Set rsTBL = Nothing
				y = 0
				tempID2 = ""
				Do Until y = x
					'tempID = tmpEmpID(y)
					'If tempID <> tempID2 And tempID2 <> "" Then
					'	myMaxHrs = GetAllwdHrs(tempID2)
					'	myOverHrs = tempHrs - myMaxHrs
					'	If myOverHrs > 0 Then
					'		strBODY = strBODY & strBODYtemp
					'		strBODY = strBODY & "<tr><td align='center'>&nbsp;</td><td align='center'>&nbsp;</td><td align='center'>&nbsp;</td>" & _
					'			"<td align='center'><font size='1' face='trebuchet ms'>" & tempHrs & "</td>" & _
					'			"<td align='center'><font size='1' face='trebuchet ms'>" & myOverHrs & "</td><td align='center'>&nbsp;</td></tr>"
					'	End If
					'	strBODYtemp = ""
					'	tempHrs = 0
					'End If
						myMaxHrs = GetAllwdHrs(tmpEmpID(y))
					'	tempHrs = tempHrs + tmpHrs(y)
					'	strBODYtemp = strBODYtemp & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & GetName(tmpWorID(y)) & _
					'			"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & GetName2(tmpEmpID(y)) & _
					'			"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & myMaxHrs & _
					'			"</font></td>" & _
					''			"<td align='center'><font size='1' face='trebuchet ms'>" & tmpHrs(y) & "</font></td>" & _
					'			"<td align='center'><font size='1' face='trebuchet ms'>&nbsp;</font></td>" & _
					'			"<td align='center'><font size='1' face='trebuchet ms'>" & GetPM(GetPM1(tmpWorID(y))) & "</font></td></tr>" 
					OverHrs = tmpHrs(y) - myMaxHrs
					if 	OverHrs > 0  Then		
						mytframe = tmpDates(y) & " - " & DateAdd("d", 6, tmpDates(y))
						strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & GetName(tmpWorID(y)) & _
								"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & GetName2(tmpEmpID(y)) & _
								"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & mytframe  & _
								"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & myMaxHrs & _
								"</font></td>" & _
								"<td align='center'><font size='1' face='trebuchet ms'>" & tmpHrs(y) & "</font></td>" & _
								"<td align='center'><font size='1' face='trebuchet ms'>&nbsp;</font></td>" & _
								"<td align='center'><font size='1' face='trebuchet ms'>" & GetPM(GetPM1(tmpWorID(y))) & "</font></td></tr>"
						strBODY = strBODY & "<tr><td align='center'>&nbsp;</td><td align='center'>&nbsp;</td><td align='center'>&nbsp;</td><td align='center'>&nbsp;</td>" & _
								"<td align='center'><font size='1' face='trebuchet ms'>" & tempHrs & "</td>" & _
								"<td align='center'><font size='1' face='trebuchet ms'>" & OverHrs & "</td><td align='center'>&nbsp;</td></tr>"
					End If
					'tempID2 = tempID
					y = y + 1
				Loop
			ElseIf Request("SelRep") = 65 Then	'active consumer w/ agency
				Session("MSG") = "All Active Consumer with Case Management Co. report."
				strHEAD = "<tr bgcolor='#040C8B'></font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Name</b> " &_
					"</font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Mailing Address</b></font></td><td align='center'>" & _
					"<font size='1' face='trebuchet ms' color='white'><b>Phone No.</b></font></td><td align='center'><font size='1' " & _
					"face='trebuchet ms' color='white'><b>DOB</b></font></td><td align='center'><font size='1' " & _
					"face='trebuchet ms' color='white'><b>Agency</b></font></td></tr>"
				Set rsTBL = Server.CreateObject("ADODB.RecordSet")
				sqlTBL = "SELECT lname, fname, maddress, mcity, mstate, mzip, DOB, PhoneNo, Consumer_t.Medicaid_Number, CMID FROM Consumer_t, C_Status_t, CMCon_T " & _
					"WHERE CID = Consumer_t.Medicaid_number AND C_Status_t.Medicaid_number = Consumer_t.Medicaid_number AND Active = True ORDER BY lname, fname"
				rsTBL.Open sqlTBL, g_strCONN, 1, 3
				Do Until rsTBL.EOF
					strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("lname") & ", " & _
						rsTBL("fname") & "</font></td><td align='center'><font size='1' face='trebuchet ms'>" & _
						rsTBL("mAddress") & ", " & rsTBL("mCity") & ", " & rsTBL("mState") & ", " & rsTBL("mZip") & "</font></td><td align='center'>" & _
						"<font size='1' face='trebuchet ms'>&nbsp;" & rsTBL("PhoneNo") & "</td><td align='center'><font size='1' " & _
						"face='trebuchet ms'>&nbsp;" & rsTBL("DOB") & "</font></td><td align='center'><font size='1' " & _
						"face='trebuchet ms'>&nbsp;" & GetCMAgency(rsTBL("CMID")) & "</font></td></tr>"
					rsTBL.MoveNext
				Loop
				rsTBL.Close
				Set rsTBL = Nothing
			ElseIf Request("SelRep") = 66 Then	'simulate process items
				PDate = Date
				markerX = 0
				If Request("ToD8") <> "" Then 
					If IsDate(Request("ToD8")) Then
						Pdate = Request("ToD8")
					Else
						Response.Redirect "specrep.asp?err=66"
					End If
				End If 
				''''''''''''
					difwk = DateDiff("ww", wk1, PDate)
		myDATE = PDate
    If difwk >= 0 Then
        wknum = difwk + 1
        If Z_IsOdd2(wknum) Then
            If WeekdayName(Weekday(myDATE), True) = "Sun" Then
                sunDATE = myDATE
                satDATE = DateAdd("d", 13, sunDATE)
            Else
                difDATE = DatePart("w", myDATE)
                sunDATE = DateAdd("d", -CInt(difDATE - 1), myDATE)
                satDATE = DateAdd("d", 13, sunDATE)
            End If
        Else
            If WeekdayName(Weekday(myDATE), True) = "Sun" Then
                sunDATE = DateAdd("d", -7, myDATE)
                satDATE = DateAdd("d", 13, sunDATE)
            Else
                difDATE = DatePart("w", myDATE)
                sunDATE = DateAdd("d", -CInt(difDATE + 6), myDATE)
                satDATE = DateAdd("d", 13, sunDATE)
            End If
        End If
    Else
        wknum = difwk
        If Not Z_IsOdd2(wknum) Then
            If WeekdayName(Weekday(myDATE), True) = "Sun" Then
                sunDATE = myDATE
                satDATE = DateAdd("d", 13, sunDATE)
            Else
                difDATE = DatePart("w", myDATE)
                sunDATE = DateAdd("d", -CInt(difDATE - 1), myDATE)
                satDATE = DateAdd("d", 13, sunDATE)
            End If
        Else
            If WeekdayName(Weekday(myDATE), True) = "Sun" Then
                sunDATE = DateAdd("d", -7, myDATE)
                satDATE = DateAdd("d", 13, sunDATE)
            Else
                difDATE = DatePart("w", myDATE)
                sunDATE = DateAdd("d", -CInt(difDATE + 6), myDATE)
                satDATE = DateAdd("d", 13, sunDATE)
            End If
        End If
    End If
			
		Set rsProc = Server.CreateObject("ADODB.RecordSet")
		sqlProc = "SELECT * FROM [Tsheets_t]"
		Session("MSG") = "Records between " & sunDATE & " - " & satDATE & " has been processed for "
		If Request("seltype4") = 2 Then
			sqlProc = sqlProc & ", consumer_t  WHERE client = medicaid_number AND code = 'M' "
			Session("MSG") = Session("MSG") & "medicaid (simulation)."
		ElseIf Request("seltype4") = 3 Then
			sqlProc = sqlProc & ", consumer_t  WHERE client = medicaid_number AND (code = 'P' OR code = 'C' OR code = 'A') "
			Session("MSG") = Session("MSG") & "private pay (simulation)."
		End If
		sqlProc = sqlProc & "AND date <= #" & satDATE & "# AND date >= #" & sunDATE & "# AND" 
		mySunDate = sunDATE
		If Request("seltype4") = 2 Then
			sqlProc = sqlProc & " IsNull(ProcMed) AND EXT = False ORDER BY lname, fname ASC, date DESC"
		ElseIf Request("seltype4") = 3 Then
			sqlProc = sqlProc & " IsNull(ProcPriv) AND EXT = False ORDER BY lname, fname ASC, date DESC"
		End If
		rsProc.Open sqlProc, g_strCONN, 1, 3
		If Not rsProc.EOF Then
			markerX = 1
			'If Request("seltype4") = 2 Then
			'	strHEAD = "<tr bgcolor='#040C8B'><td align='center' width='100px'><font size='1' face='trebuchet ms' color='white' color='white'>Timesheet Week</font>" & _
			'		"</td><td align='center' width='100px'><font size='1' face='trebuchet ms' color='white' color='white'>Medicaid</font></td><td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>" & _
			'		"Name</font></td><td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>Hours</font></td><td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>" & _
			'		"Units</font></td><td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>" & _
			'		"Amount</font></td><td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>Notes</font></td></tr>"
				'strProcHexp = "Timesheet Week, Medicaid, Last Name, First Name, Hours, Units,Amount, Notes"
			'ElseIf Request("seltype4") = 3 Then
				strHEAD = "<tr bgcolor='#040C8B'><td align='center' width='100px'><font size='1' face='trebuchet ms' color='white' color='white'>Timesheet Week</font>" & _
					"</td><td align='center' width='100px'><font size='1' face='trebuchet ms' color='white' color='white'>Medicaid</font></td><td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>" & _
					"Consumer</font></td><td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>" & _
					"PCSP</font></td><td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>Regular Hours</font></td><td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>" & _
					"Holiday Hours</font></td><td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>" & _
					"Extended Hours</font></td><td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>" & _
					"Rate</font></td><td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>" & _
					"Mileage</font></td><td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>" & _
					"Notes</font></td><td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>User ID</font></td></tr>"
				'strProcHexp = "Timesheet Week, Medicaid, Consumer Last Name, Consumer First Name, PCSP Last Name, PCSP First Name, Regular Hours, Holiday Hours, Extended Hours, Rate, Mileage,Notes, User ID"
			'End If
			Do Until rsProc.EOF
				myNotes = ""
				'Get EXT
				Set rsEXT = Server.CreateObject("ADODB.RecordSet")
				sqlEXT = "SELECT * FROM [Tsheets_t] WHERE date = #" & rsProc("date") & "# AND Client = '" & rsProc("client") & "' AND emp_id = '" & rsProc("emp_ID") & "' " & _
				 	"AND EXT = True AND TimeStamp = #" & rsProc("TimeStamp") & "#"
				rsEXT.Open sqlEXT, g_strCONN, 1, 3
				If Not rsEXT.EOF Then
					extHrs = rsEXT("mon") + rsEXT("tue") + rsEXT("wed") + rsEXT("thu") + rsEXT("fri") + rsEXT("sat") + rsEXT("sun")
					If extHrs <> 0 Then
						myNotes = rsEXT("misc_notes")
					End If
				Else
					extHrs = 0
				End If
				rsEXT.Close
				Set rsEXT = Nothing
				'''''''
				strDate = rsProc("Date") & " - " & DateAdd("d", 6, rsProc("Date"))
				If Request("seltype4") = 2 Then
					regHrs = rsProc("mon") + rsProc("tue") + rsProc("wed") + rsProc("thu") + rsProc("fri") + rsProc("sat") + rsProc("sun")
					holHrs = 0
				ElseIf Request("seltype4") = 3 Then
					Hmon = GetHoliday(rsProc("date"), rsProc("mon"), "MON")
					Htue = GetHoliday(rsProc("date"), rsProc("tue"), "TUE")
					Hwed = GetHoliday(rsProc("date"), rsProc("wed"), "WED")
					Hthur = GetHoliday(rsProc("date"), rsProc("thu"), "THU")
					Hfri = GetHoliday(rsProc("date"), rsProc("fri"), "FRI")
					Hsat = GetHoliday(rsProc("date"), rsProc("sat"), "SAT")
					Hsun = GetHoliday(rsProc("date"), rsProc("sun"), "SUN")
					holHrs = Hmon + Htue + Hwed + Hthur + Hfri + Hsat + Hsun
					tmpregHrs = rsProc("mon") + rsProc("tue") + rsProc("wed") + rsProc("thu") + rsProc("fri") + rsProc("sat") + rsProc("sun")
					regHrs = tmpregHrs - holHrs
				End If 
				myMile = rsProc("mile") + rsProc("amile")
				myNotes = rsProc("misc_notes") & "<br>" & myNotes
				strBOdy = strBody & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & strDate & _
					"</font></td>"
					If Session("lngType") = 1 Or Session("lngType") = 2 Then 
							strBODY = strBODY & "<td align='center'><font size='1' face='trebuchet ms'>" & rsProc("client") & "</font></td>"
					Else
							strBODY = strBODY & "<Td>&nbsp;</td>"
						End If 
						strBODY = strBODY & "<td align='center'><font size='1' face='trebuchet ms'>" & rsProc("lname") & ", " & rsProc("fname") & "</font></td><td align='center'><font size='1' face='trebuchet ms'>" & _
					GetNameWork(rsProc("emp_id")) & "</font></td><td align='center'><font size='1' face='trebuchet ms'>" & _
					Z_FormatNumber(regHrs, 2) & "</font></td><td align='center'><font size='1' face='trebuchet ms'>" & _
					Z_FormatNumber(holHrs, 2) & "</font></td><td align='center'><font size='1' face='trebuchet ms'>" & _
					Z_FormatNumber(extHrs, 2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' >" & _
					Z_FormatNumber(GetPRate2(rsProc("client")),2) & "</font></td><td align='center'><font size='1' face='trebuchet ms'>" & _
					Z_FormatNumber(myMile, 2) & "</font></td><td align='center'><font size='1' face='trebuchet ms'>" & _
					myNotes & "</font></td><td align='center'><font size='1' face='trebuchet ms'>" & _
					GetUser(rsProc("author")) & "</font></td></tr>" & vbCrLf
				rsProc.MoveNExt
			Loop
		Else
			'NO RECORDS FOUND
			If Request("seltype4") = 2 Then
					Session("MSG") = "No medicaid records found on " & sunDATE & " - " & satDATE & " for processing."
			ElseIf Request("seltype4") = 3 Then
					Session("MSG") = "No Private Pay records found on " & sunDATE & " - " & satDATE & " for processing."
			End If
		End If
	rsProc.CLose
	Set rsProc = Nothing	
	'NOT within 2 week period
	Set rsProc2 = Server.CreateObject("ADODB.RecordSet")
	sqlProc2 = "SELECT * FROM tsheets_t"
	If Request("seltype4") = 2 Then
		sqlProc2 = sqlProc2 & ", consumer_t  WHERE code = 'M' AND date < #" & sunDATE & "# AND client = medicaid_number AND IsNull(ProcMed) AND EXT = False ORDER BY lname, fname, Date ASC, client ASC"
	ElseIf Request("seltype4") = 3 Then
		sqlProc2 = sqlProc2 & ", consumer_t  WHERE (code = 'P' OR code = 'C' OR code = 'A') AND date < #" & sunDATE & "# AND client = medicaid_number AND IsNull(ProcPriv) AND EXT = False ORDER BY lname, fname, Date ASC, client ASC"
	End If
	MarkerX = 0
	rsProc2.Open sqlProc2, g_strCONN, 1, 3	
	If Not rsProc2.EOF THen
		strHEAD2 = "<tr bgcolor='#040C8B'><td colspan='11'><font size='1' face='trebuchet ms' color='white'><b>Processed items before the set payroll period</b></font></td></tr>"
		strHEAD2 = strHEAD2 & "<tr bgcolor='#040C8B'><td align='center' width='100px'><font size='1' face='trebuchet ms' color='white' color='white'>Timesheet Week</font>" & _
					"</td><td align='center' width='100px'><font size='1' face='trebuchet ms' color='white' color='white'>Medicaid</font></td><td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>" & _
					"Consumer</font></td><td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>" & _
					"PCSP</font></td><td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>Regular Hours</font></td><td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>" & _
					"Holiday Hours</font></td><td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>" & _
					"Extended Hours</font></td><td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>" & _
					"Rate</font></td><td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>" & _
					"Mileage</font></td><td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>" & _
					"Notes</font></td><td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>User ID</font></td></tr>"
		Do Until rsProc2.EOF
			myNotes = ""
			'Get EXT
			Set rsEXT = Server.CreateObject("ADODB.RecordSet")
			sqlEXT = "SELECT * FROM [Tsheets_t] WHERE date = #" & rsProc2("date") & "# AND Client = '" & rsProc2("client") & "' AND emp_id = '" & rsProc2("emp_ID") & "' " & _
			 	"AND EXT = True AND TimeStamp = #" & rsProc2("TimeStamp") & "#"
			rsEXT.Open sqlEXT, g_strCONN, 1, 3
			If Not rsEXT.EOF Then
				extHrs = rsEXT("mon") + rsEXT("tue") + rsEXT("wed") + rsEXT("thu") + rsEXT("fri") + rsEXT("sat") + rsEXT("sun")
				If extHrs <> 0 Then
					myNotes = rsEXT("misc_notes")
				End If
			Else
				extHrs = 0
			End If
			rsEXT.Close
			Set rsEXT = Nothing
			'''''''
			strDate = rsProc2("Date") & " - " & DateAdd("d", 6, rsProc2("Date"))
			If Request("seltype4") = 2 Then
				regHrs = rsProc2("mon") + rsProc2("tue") + rsProc2("wed") + rsProc2("thu") + rsProc2("fri") + rsProc2("sat") + rsProc2("sun")
				holHrs = 0
			ElseIf Request("seltype4") = 3 Then
				Hmon = GetHoliday(rsProc2("date"), rsProc2("mon"), "MON")
				Htue = GetHoliday(rsProc2("date"), rsProc2("tue"), "TUE")
				Hwed = GetHoliday(rsProc2("date"), rsProc2("wed"), "WED")
				Hthur = GetHoliday(rsProc2("date"), rsProc2("thu"), "THU")
				Hfri = GetHoliday(rsProc2("date"), rsProc2("fri"), "FRI")
				Hsat = GetHoliday(rsProc2("date"), rsProc2("sat"), "SAT")
				Hsun = GetHoliday(rsProc2("date"), rsProc2("sun"), "SUN")
				holHrs = Hmon + Htue + Hwed + Hthur + Hfri + Hsat + Hsun
				tmpregHrs = rsProc2("mon") + rsProc2("tue") + rsProc2("wed") + rsProc2("thu") + rsProc2("fri") + rsProc2("sat") + rsProc2("sun")
				regHrs = tmpregHrs - holHrs
			End If 
			myMile = rsProc2("mile") + rsProc2("amile")
			myNotes = rsProc2("misc_notes") & "<br>" & myNotes
			strBOdy2 = strBody2 & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & strDate & _
				"</font></td>"
				If Session("lngType") = 1 Or Session("lngType") = 2 Then 
							strBOdy2 = strBOdy2 & "<td align='center'><font size='1' face='trebuchet ms'>" & rsProc2("client") & "</font></td>"
					Else
							strBOdy2 = strBOdy2 & "<Td>&nbsp;</td>"
						End If 
						strBOdy2 = strBOdy2 & "<td align='center'><font size='1' face='trebuchet ms'>" & rsProc2("lname") & ", " & rsProc2("fname") & "</font></td><td align='center'><font size='1' face='trebuchet ms'>" & _
				GetNameWork(rsProc2("emp_id")) & "</font></td><td align='center'><font size='1' face='trebuchet ms'>" & _
				Z_FormatNumber(regHrs, 2) & "</font></td><td align='center'><font size='1' face='trebuchet ms'>" & _
				Z_FormatNumber(holHrs, 2) & "</font></td><td align='center'><font size='1' face='trebuchet ms'>" & _
				Z_FormatNumber(extHrs, 2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' >" & _
				Z_FormatNumber(GetPRate2(rsProc2("client")),2) & "</font></td><td align='center'><font size='1' face='trebuchet ms'>" & _
				Z_FormatNumber(myMile, 2) & "</font></td><td align='center'><font size='1' face='trebuchet ms'>" & _
				myNotes & "</font></td><td align='center'><font size='1' face='trebuchet ms'>" & _
				GetUser(rsProc2("author")) & "</font></td></tr>" & vbCrLf
			rsProc2.MoveNext
		Loop
	End If
	rsProc2.Close
	Set rsProc2 = Nothing	
ElseIf Request("SelRep") = 67 Then 'newsletter
			Session("MSG") = "News letter report."
			strHEAD = "<tr bgcolor='#040C8B'></font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Name</b> " &_
					"</font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Mailing Address</b></font></td></tr>"'<td align='center'>" & _
					'"<font size='1' face='trebuchet ms' color='white'><b>Code</b></font></td></tr>"
			Set rsTBL = Server.CreateObject("ADODB.RecordSet")
			sqlTBL = "SELECT * FROM consumer_t, c_status_T where C_Status_t.Medicaid_number = Consumer_t.Medicaid_number AND Active = True and code = 'M' " & _
				"ORDER BY mZip, mAddress"
			rsTBL.Open sqlTBL, g_strCONN, 3, 1
			Do Until rsTBL.EOF
				If rsTBL("mAddress") <> "" Then
					strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("lname") & ", " & _
						rsTBL("fname") & "</font></td><td align='center'><font size='1' face='trebuchet ms'>" & _
						rsTBL("mAddress") & ", " & rsTBL("mCity") & ", " & rsTBL("mState") & ", " & rsTBL("mZip") & "</font></td></tr>"
				End If
				rsTBL.MoveNext
			Loop
			rsTBL.Close
			Set rsTBL = Nothing
			Set rsTBL = Server.CreateObject("ADODB.RecordSet")
			sqlTBL = "SELECT * FROM Representative_t, conrep_T, C_status_T WHERE Cint(RID) = Representative_t.index " & _
				"AND CID = C_status_T.Medicaid_Number AND Active = True " & _
				"ORDER BY zip, address"
			rsTBL.Open sqlTBL, g_strCONN, 3, 1
			Do Until rsTBL.EOF
				If rsTBL("Address") <> "" Then
					strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("lname") & ", " & _
						rsTBL("fname") & "</font></td><td align='center'><font size='1' face='trebuchet ms'>" & _
						rsTBL("Address") & ", " & rsTBL("City") & ", " & rsTBL("State") & ", " & rsTBL("Zip") & "</font></td></tr>"
				End If
				rsTBL.MoveNext
			Loop
			rsTBL.Close
			Set rsTBL = Nothing
			Set rsTBL = Server.CreateObject("ADODB.RecordSet")
			sqlTBL = "SELECT * FROM Worker_t WHERE STATUS = 'Active' ORDER by mzip, mAddress"
			rsTBL.Open sqlTBL, g_strCONN, 3, 1
			Do Until rsTBL.EOF
				If rsTBL("mAddress") <> "" Then
					strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("lname") & ", " & _
						rsTBL("fname") & "</font></td><td align='center'><font size='1' face='trebuchet ms'>" & _
						rsTBL("mAddress") & ", " & rsTBL("mCity") & ", " & rsTBL("mState") & ", " & rsTBL("mZip") & "</font></td></tr>"
				End If
				rsTBL.MoveNext
			Loop
			rsTBL.Close
			Set rsTBL = Nothing
		ElseIf Request("SelRep") = 68 Then 'callerID
			Session("MSG") = "Unapproved Caller ID report."
			strHEAD = "<tr bgcolor='#040C8B'></font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Worker</b> " &_
					"</font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Consumer</b></font></td><td align='center'>" & _
					"<font size='1' face='trebuchet ms' color='white'><b>Date</b></font></td><td align='center'>" & _
					"<font size='1' face='trebuchet ms' color='white'><b>Phone Number</b></font></td><td align='center'>" & _
					"<font size='1' face='trebuchet ms' color='white'><b>RHICC</b></font></td></tr>"
			Set rsTBL = Server.CreateObject("ADODB.RecordSet")
			sqlTBL = "SELECT * " & _ 
				"FROM Tsheets_T, Worker_T, consumer_t " & _ 
				"WHERE Client = Consumer_T.[medicaid_number] " & _ 
				"AND emp_ID = worker_T.[Social_Security_Number] " & _ 
				"AND CallerID <> '' " & _ 
				"ORDER BY consumer_t.[fname], consumer_t.[lname], date "
			
			rsTBL.Open sqlTBL, g_strCONN, 3, 1
			Do Until rsTBL.EOF
				If NOT ApproveNum(rsTBL("CallerID"), rsTBL("client")) Then
					strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("worker_T.lname") & ", " & _
						rsTBL("worker_T.fname") & "</font></td><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("Consumer_T.lname") & ", " & _
						rsTBL("Consumer_T.fname") & "</font></td><td align='center'><font size='1' face='trebuchet ms'>" & _
						rsTBL("Date") & "</font></td><td align='center'><font size='1' face='trebuchet ms'>" & _
						rsTBL("CallerID") & "</font></td><td align='center'><font size='1' face='trebuchet ms'>" & _
						getname3(rsTBL("pm1")) & "</font></td></tr>"
				End If
				rsTBL.MoveNext
			Loop
			rsTBL.Close
			Set rsTBL = Nothing
		ElseIf Request("SelRep") = 69 Then 'medicaid consumer record
			dim myDay(6), tbl(6), myUnits(6), myAC(6), mysun()
			Function Z_FixEOF(rs)
				Z_FixEOF=0
				if not(rs.eof) Then
					Z_FixEOF=rs("val")
					rs.movenext
				End if
			ENd fUnction
			function ACValue(xxx)
				If xxx = "" Then exit function
				Set rsAC = Server.CreateObject("ADODB.RecordSet")
				sqlAC = "SELECT [desc] FROM activity_T WHERE code = " & xxx'strAC(ctr)
				rsAC.Open sqlAC, g_strCONN, 3, 1
				If Not rsAC.EOF Then
					tmpAC = tmpAC & rsAC("desc") & "<br>" 
				End If
				rsAC.Close
				ACValue = tmpAC
				Set rsAC = Nothing
			End function
			
			myDay(0) = "sun"
			myDay(1) = "mon"
			myDay(2) = "tue"
			myDay(3) = "wed"
			myDay(4) = "thu"
			myDay(5) = "fri"
			myDay(6) = "sat"
			
			'get allowed sundays
			If Request("FrmD8") <> "" Then
				If IsDate(Request("FrmD8")) Then
					Redim Preserve mysun(0)
					mysun(0) = GetSun(Request("FrmD8"))
					msgdtefrom =  " from " & Request("FrmD8")
				End If
			End If

			If Request("ToD8") <> "" Then
				If IsDate(Request("ToD8")) Then
					lngI = 0
					Do Until cdate(mysun(lngI)) > cdate(Request("ToD8"))
						lngI = lngI + 1
						Redim Preserve mysun(lngI)
						mysun(lngI) = DateAdd("d", 7, cdate(mysun(lngI - 1)))
					Loop
					msgdteto = " to " & Request("ToD8")
				End If
			End If
			ctrs = 0
			do until ctrs = ubound(mysun)
				'response.write mysun(ctrs) & "<br>"
				closedate = cdate(mysun(ctrs)) 'Request("closedate")
				ctrs = ctrs + 1
			Loop
			''''''''''''''''''''
			'closedate =cdate("9/25/2011")
			'1st week			
			Session("MSG") = "Medicaid consumer Report "
			sqlTBL = "SELECT distinct client, lname, fname, address, city, state, zip FROM Tsheets_T, Consumer_T WHERE client = medicaid_number AND code = 'M' AND " & _
				"date = #" & CDate(closedate) & "# AND EXT = False AND NOT ProcMed IS NULL " 
			If Request("selcon") <> "0" Then 
				sqlTBL = sqlTBL & "AND Client = '" & Request("selcon") & "' "
				Session("MSG") = Session("MSG") & "for " & getname2(Request("selcon")) & " "
			End If	
			sqlTBL = sqlTBL & "ORDER BY lname, fname"
			Session("MSG") = Session("MSG") & msgdtefrom & msgdteto
			Set rsTBL = Server.CreateObject("ADODB.RecordSet")
			rsTBL.Open sqlTBL, g_strCONN, 3, 1
			Do Until rsTBL.EOF
				myTotHours = 0
				strBODY2 = "<table style='border: 1px solid black;'><tr><td align='left' colspan='4'><b>Medicaid Consumer Record</b></td>" & vbCrLf & _
							"<td colspan='4' rowspan='4'><img src='images/printlogo.gif' border='0'></td></tr>" & vbCrLf & _
							"<tr><td align='left' colspan='4'><font size='1' face='trebuchet ms'><b>" & rsTBL("Lname") & ", " & rsTBL("fname") & "</b></font></td></tr>" & vbCrLf & _
							"<tr><td align='left' colspan='4'><font size='1' face='trebuchet ms'><b>" & rsTBL("Address") & "</b></font></td></tr>" & vbCrLf & _
							"<tr><td align='left' colspan='4'><font size='1' face='trebuchet ms'><b>" & rsTBL("City") & ", " & rsTBL("state") & ", " & rsTBL("zip") & "</b></font></td></tr>" & vbCrLf & _
							"<tr><td colspan='8'>&nbsp;</td></tr><tr><td colspan='8'>&nbsp;</td></tr>" & vbCrLf & _
							"<tr><td>&nbsp;</td><td align='center'><font size='1' face='trebuchet ms'>Sunday</font></td><td align='center'><font size='1' face='trebuchet ms'>Monday</font></td>" & _
							"<td align='center'><font size='1' face='trebuchet ms'>Tuesday</font></td><td align='center'><font size='1' face='trebuchet ms'>Wednesday</font></td>" & _
							"<td align='center'><font size='1' face='trebuchet ms'>Thursday</font></td><td align='center'><font size='1' face='trebuchet ms'>Friday</font></td>" & _
							"<td align='center'><font size='1' face='trebuchet ms'>Saturday</font></td></tr>" & vbCrLf & _
							"<tr><td>&nbsp;</td><td align='center'><font size='1' face='trebuchet ms'>" & closedate & "</font></td><td align='center'><font size='1' face='trebuchet ms'>" & DateAdd("d", 1, closedate) & "</font></td>" & _
							"<td align='center'><font size='1' face='trebuchet ms'>" & DateAdd("d", 2, closedate) & "</font></td><td align='center'><font size='1' face='trebuchet ms'>" & DateAdd("d", 3, closedate) & "</font></td>" & _
							"<td align='center'><font size='1' face='trebuchet ms'>" & DateAdd("d", 4, closedate) & "</font></td><td align='center'><font size='1' face='trebuchet ms'>" & DateAdd("d", 5, closedate) & "</font></td>" & _
							"<td align='center'><font size='1' face='trebuchet ms'>" & DateAdd("d", 6, closedate) & "</font></td></tr>" & vbCrLf & _
							"</tr>" & vbCrLf 
				sqlTBLWor = "SELECT distinct emp_ID, lname, fname FROM Tsheets_T, worker_T WHERE emp_ID = Social_Security_Number AND client = '" & rsTBL("client") & "' " & _
					"AND date = #" & CDate(closedate) & "# AND EXT = False AND NOT ProcMed IS NULL ORDER BY lname, fname, emp_ID" 	
				Set rsTBLWor = Server.CreateObject("ADODB.RecordSet")
				rsTBLWor.Open sqlTBLWor, g_strCONN, 3, 1
				Do Until rsTBLWor.EOF
					for lngI = 0 to 6
						sql1 = "SELECT emp_id, " & myDay(lngI) & " AS [val], misc_notes, date FROM tsheets_t WHERE client= '" & rsTBL("client") & "' AND emp_ID= '" & rsTBLWor("emp_ID") & "' AND date = #" & CDate(closedate) & "# and " & myDay(lngI) & " > 0 AND EXT = False AND NOT ProcMed IS NULL ORDER BY timestamp"
						Set tbl(lngI) = Server.CreateObject("ADODB.Recordset")
						tbl(lngI).Open sql1, g_strCONN, 3, 1
					Next
					do while not (tbl(0).eof and tbl(1).eof and tbl(2).eof and tbl(3).eof and tbl(4).eof and tbl(5).eof and tbl(6).eof)
						strBODY2 = strBODY2 & "<tr><td colspan='8'>&nbsp;</td></tr><tr><td align='left' colspan='8'><font size='1' face='trebuchet ms'><b>" & GetNameWork(rsTBLWor("emp_ID")) & "</b></font></td></tr>" & vbCrLf & _
							"<tr><td style='border: 1px solid black; text-align: right;'><font size='1' face='trebuchet ms'>Hours</font></td>"
						for lngI = 0 to 6
							myAC(lngI) = ""
							If Not tbl(lngI).EOF Then
								
								tmpACval =  tbl(lngI)("misc_notes")
								myACval = Split(tmpACval, ",")
								For lngI2 = 0 to Ubound(myACval)
									myAC(lngI) = myAC(lngI) & ACValue(myACval(lngI2))
								Next
							End If
							myHours = Z_FixEOF(tbl(lngI))'ValidDate(Request("FrmD8"), Request("ToD8"), tbl(lngI)("date"), Z_FixEOF(tbl(lngI)), myDay(lngI))
							strBODY2 = strBODY2 & "<td align='center' style='border: 1px solid black; text-align: center;'><font size='1' face='trebuchet ms'>" & myHours & "</font></td>"
							myUnits(lngI) = myHours * 4
							myTotHours = myTotHours + myHours
						next
						strBODY2 = strBODY2 & "</tr>" & vbCrLf & "<tr><td style='border: 1px solid black; text-align: right;'><font size='1' face='trebuchet ms'>Units</font></td>"
							for lngI = 0 to 6
									strBODY2 = strBODY2 & "<td align='center' style='border: 1px solid black; text-align: center;'><font size='1' face='trebuchet ms'>" & myUnits(lngI) & "</font></td>"
							next
							strBODY2 = strBODY2 & "</tr>" & vbCrLf & "<tr><td valign='top' style='border: 1px solid black; text-align: right;'><font size='1' face='trebuchet ms'>Activities</font></td>"
							for lngI = 0 to 6
									strBODY2 = strBODY2 & "<td align='center' style='border: 1px solid black; text-align: center;'><font size='1' face='trebuchet ms'>" & myAC(lngI) & "</font></td>"
							next
							strBODY2 = strBODY2 & "</tr>" & vbCrLf
					Loop
					
					for lngI =0 to 6
						tbl(lngI).Close
						set tbl(lngi) = Nothing
					next
					rsTBLWor.MoveNext
				Loop
				rsTBLWor.Close
				Set rsTBLWor = Nothing
				strBODY2 = strBODY2 & "<tr><td colspan='8'>&nbsp;</td></tr><tr><td colspan='8'>&nbsp;</td></tr><tr><td><font size='1' face='trebuchet ms'>Total Hours</font></td>" & vbCrLf & _
						"<td align='center'><font size='1' face='trebuchet ms'><b><u>" & myTotHours & "</u></b></font></td><td colspan='7'>&nbsp;</td></tr><tr><td><font size='1' face='trebuchet ms'>Total Units</font></td>" & vbCrLf & _
						"<td align='center'><font size='1' face='trebuchet ms'><b><u>" & myTotHours * 4 & "</u></b></font></td><td colspan='7'>&nbsp;</td></tr>" & vbCrLf & _
						"</table><div class='page-break'><br></div>" & vbCrLf
				if myTotHours > 0 Then strBODY = strBODY & strBODY2
				rsTBL.MoveNext
			Loop
			rsTBL.Close
			Set rsTBL = Nothing
		End If
		Session("PrintPrev") = strHEAD & "|" & strBODY & "|" & Session("MSG") & "|" & strBODY2 & "|" & strHEAD2
	End If
	
%>
<html>
	<head>
		<title>LSS - SmartCare - Advance Report</title>
		<link href="styles.css" type="text/css" rel="stylesheet" media="print">
		<link href="CalendarControl.css" type="text/css" rel="stylesheet">
		<script src="CalendarControl.js" language="javascript"></script>
		<script language='JavaScript'>
		function y2k(numinput) {
			var number = new Number(numinput);
			return (number < 1000) ? number + 2000 : number; 
		}

		function dateConv(zdate) {
			var rval, sl1, sl2, strDD = new String(zdate);
			var arrS = strDD.split("/");
			
			var ddtmp = new Date(arrS[0] + "/" + arrS[1] + "/" + y2k(arrS[2]) );
			return(ddtmp);
		}

		function getWeek(inputstr) {
			var strDt = new String(inputstr) //month + "/" + day + "/" + year);
		    var when = dateConv(strDt);
			
			var year = when.getFullYear();
			//alert ("StrDt: " + strDt + "\nDate: " + when + "\n\nYear: " + when.getFullYear());
		    var newYear = new Date(year,0,1);
		    var offset = 7 + 1 - newYear.getDay();
		    if (offset == 8) offset = 1;
		    var daynum = ((Date.UTC(y2k(year),when.getMonth(),when.getDate(),0,0,0) - Date.UTC(y2k(year),0,1,0,0,0)) /1000/60/60/24) + 1;
		    var weeknum = Math.floor((daynum-offset+7)/7);
		    //alert ("pre. week#:" + weeknum);
		    if (weeknum == 0) {
		        year--;
		        var prevNewYear = new Date(year,0,1);
		        var prevOffset = 7 + 1 - prevNewYear.getDay();
		        if (prevOffset == 2 || prevOffset == 8) weeknum = 53; else weeknum = 52;
		    }
		    return weeknum;
		}
		function getWk()
		{
		var sunDATE, satDATE, tmpDateE, weekno, tmpDate = new Date();
		tmpDate = dateConv(document.frmRep.closedate.value);
		weekno = getWeek(document.frmRep.closedate.value);
		var tmpDateE = new Number(tmpDate.getDay());
		if (weekno  % 2 == 0)
			{

				if (tmpDateE == 0)
				{
					sunDATE = new Date (tmpDate);
					satDATE = new Date (tmpDate.setDate(tmpDate.getDate() + 13) );
				}
				else
				{
					//tmpDateE = tmpDate.getDay;
					sunDATE = new Date (tmpDate.setDate(tmpDate.getDate() - (tmpDateE)) ); 
					satDATE = new Date (tmpDate.setDate(tmpDate.getDate() + 13) ); 
				}
			}
		else
			{
				if (tmpDateE == 0)
				{
					sunDATE = new Date (tmpDate.setDate(tmpDate.getDate() - 7));
					satDATE = new Date (tmpDate.setDate(tmpDate.getDate() + 13));
				}
				else
				{
					//tmpDateE = tmpDate.getDay;
					sunDATE = new Date (tmpDate.setDate(tmpDate.getDate() - (tmpDateE + 7)) ); 
					satDATE = new Date(tmpDate.setDate(tmpDate.getDate() + 13) ); 
				}
			} 
		//document.write('WEEK: ' + weekno + '');
		//alert("WEEK: " + weekno + "\nsun: " + sunDATE + "" );
		//document.frmRep.Todate.value = (satDATE.getMonth() + 1) + "/" + satDATE.getDate() + "/" + satDATE.getFullYear();
		//document.frmRep.closedate.value = (sunDATE.getMonth() + 1) + "/" + sunDATE.getDate() + "/" + sunDATE.getFullYear();
		}
		function RemImages()
		{
			document.all.style.display="none"; 
		}
		function LandWarn()
			{
				var ans = window.confirm("Please set page orientation to landscape. Click Ok to continue. Click Cancel to stop.");
				if (ans){
				document.frmRep.action = "Print.asp";
				document.frmRep.submit();
				}
			}
		function SVSort()
			{
				document.frmRep.action = "SpecRep.asp?SelRep=" + <%=typ%> + "&Sort=" + <%=Srt%>
				document.frmRep.submit();
			}
		function ExpCSV(xxx)
		{
			if (xxx == 1) 
				{
				if (document.frmRep.SelRep.value == 19)
					{
					 document.frmRep.action = "Export.asp?sql=3";
					 document.frmRep.submit();
					}
				if (document.frmRep.SelRep.value == 28)
					{
					 document.frmRep.action = "Export.asp?sql=4";
					 document.frmRep.submit();
					}
				if (document.frmRep.SelRep.value == 2)
					{
					 document.frmRep.action = "Export.asp?sql=5";
					 document.frmRep.submit();
					}
				if (document.frmRep.SelRep.value == 3)
					{
					 document.frmRep.action = "Export.asp?sql=6";
					 document.frmRep.submit();
					}
				if (document.frmRep.SelRep.value == 35)
					{
					 document.frmRep.action = "Export.asp?sql=7";
					 document.frmRep.submit();
					}
				if (document.frmRep.SelRep.value == 14)
					{
					 document.frmRep.action = "Export.asp?sql=8";
					 document.frmRep.submit();
					}
				if (document.frmRep.SelRep.value == 20)
					{
					 document.frmRep.action = "Export.asp?sql=9";
					 document.frmRep.submit();
					}
				if (document.frmRep.SelRep.value == 36)
					{
					 document.frmRep.action = "Export.asp?sql=10";
					 document.frmRep.submit();
					}
				if (document.frmRep.SelRep.value == 55)
					{
					 document.frmRep.action = "Export.asp?sql=11&prj=0";
					 document.frmRep.submit();
					}
				if (document.frmRep.SelRep.value == 62)
					{
					 document.frmRep.action = "Export.asp?sql=12";
					 document.frmRep.submit();
					}
				if (document.frmRep.SelRep.value == 60)
					{
					 document.frmRep.action = "Export.asp?sql=13";
					 document.frmRep.submit();
					}
				if (document.frmRep.SelRep.value == 64)
					{
					 document.frmRep.action = "Export.asp?sql=14";
					 document.frmRep.submit();
					}
				if (document.frmRep.SelRep.value == 66)
					{
					 document.frmRep.action = "Export.asp?sql=15";
					 document.frmRep.submit();
					}
				if (document.frmRep.SelRep.value == 67)
					{
					 document.frmRep.action = "Export.asp?sql=16";
					 document.frmRep.submit();
					}
				}
				else
				{
					if (xxx == 2 && document.frmRep.SelRep.value == 55)
					{
						document.frmRep.action = "Export.asp?sql=11&prj=1";
					 document.frmRep.submit();
					}
				}
		}
		function hidetxt()
		{
			if (document.frmRep.SelRep.value == 18 || document.frmRep.SelRep.value == 24 || document.frmRep.SelRep.value == 59 || document.frmRep.SelRep.value == 60 || document.frmRep.SelRep.value == 29 || document.frmRep.SelRep.value == 32 || document.frmRep.SelRep.value == 31 || document.frmRep.SelRep.value == 41 || document.frmRep.SelRep.value == 47 || document.frmRep.SelRep.value == 49 || document.frmRep.SelRep.value == 50 || document.frmRep.SelRep.value == 52 || document.frmRep.SelRep.value == 66 || document.frmRep.SelRep.value == 69)
				{
				 document.frmRep.FrmD8.style.visibility = 'visible';
				 document.frmRep.ToD8.style.visibility = 'visible';
				 document.frmRep.txtFrm.style.visibility = 'visible';
				 document.frmRep.txtTo.style.visibility = 'visible';
				 if (document.frmRep.SelRep.value == 31)
					{
						document.frmRep.seltype2.style.visibility = 'visible';
					}
					else
					{
						document.frmRep.seltype2.style.visibility = 'hidden';
					}
				 if (document.frmRep.SelRep.value == 50)
					{
						 document.frmRep.FrmD8.style.visibility = 'visible';
						 document.frmRep.txtFrm.style.visibility = 'visible';
						 document.frmRep.ToD8.style.visibility = 'hidden';
						 document.frmRep.txtTo.style.visibility = 'hidden';
					}
					if (document.frmRep.SelRep.value == 60)
					{
						 document.frmRep.seltype3.style.visibility = 'visible';
					}
					else
					{
						document.frmRep.seltype3.style.visibility = 'hidden';
					}
					if (document.frmRep.SelRep.value == 66)
					{
						document.frmRep.seltype2.style.visibility = 'hidden';
					 document.frmRep.FrmD8.style.visibility = 'hidden';
					 //document.frmRep.ToD8.style.visibility = 'visible';
					 document.frmRep.txtFrm.style.visibility = 'hidden';
					 //document.frmRep.txtTo.style.visibility = 'visible';
					 document.frmRep.seltype3.style.visibility = 'hidden';
					 document.frmRep.seltype4.style.visibility = 'visible';
					}
				}
			else
				{
				 document.frmRep.seltype2.style.visibility = 'hidden';
				 document.frmRep.FrmD8.style.visibility = 'hidden';
				 document.frmRep.ToD8.style.visibility = 'hidden';
				 document.frmRep.txtFrm.style.visibility = 'hidden';
				 document.frmRep.txtTo.style.visibility = 'hidden';
				  document.frmRep.seltype3.style.visibility = 'hidden';
				  document.frmRep.seltype4.style.visibility = 'hidden';
				}
		}

		function hided8()
		{
			if (document.frmRep.SelRep.value == 27 || document.frmRep.SelRep.value == 39 || document.frmRep.SelRep.value == 44  || document.frmRep.SelRep.value == 55 || document.frmRep.SelRep.value == 51 || document.frmRep.SelRep.value == 64)
				{document.frmRep.closedate.style.visibility = 'visible';
				 document.frmRep.Todate.style.visibility = 'visible';
				 if (document.frmRep.SelRep.value == 27)
				 {document.frmRep.seltype.style.visibility = 'visible';}
				 else
				 {document.frmRep.seltype.style.visibility = 'hidden';}
				 if (document.frmRep.SelRep.value == 55)
				 {
				 	document.frmRep.txtPP.style.visibility = 'visible';
				 		document.frmRep.selopt.style.visibility = 'visible';
				 		document.frmRep.selUri.style.visibility = 'visible';
					}
					else
					{
						document.frmRep.txtPP.style.visibility = 'hidden';
						document.frmRep.selopt.style.visibility = 'hidden';
						document.frmRep.selUri.style.visibility = 'hidden';
					}
					document.frmRep.txtCal.style.visibility = 'visible';
				 document.frmRep.txtTCal.style.visibility = 'visible';
				 document.frmRep.cal1.style.visibility = 'visible';}
			else
				{document.frmRep.Todate.style.visibility = 'hidden';
				 document.frmRep.closedate.style.visibility = 'hidden';
				 document.frmRep.seltype.style.visibility = 'hidden';
				 document.frmRep.txtCal.style.visibility = 'hidden';
				 document.frmRep.txtTCal.style.visibility = 'hidden';
				 document.frmRep.txtPP.style.visibility = 'hidden';
				 document.frmRep.selopt.style.visibility = 'hidden';
				 document.frmRep.selUri.style.visibility = 'hidden';
				 document.frmRep.cal1.style.visibility = 'hidden';}
			}
			function weeknum()
			{
				document.frmRep.action = "weeknum.asp?uri=" + document.frmRep.selUri.value + "&opt=" + document.frmRep.selopt.value + "&tmpdate=" + document.frmRep.closedate.value + "&selcon=" + document.frmRep.SelCon.value;
				document.frmRep.submit();
			}
			function PrintPrev()
			{
				document.frmRep.action = "Print.asp";
				document.frmRep.submit();
			}
			function PrintPrev2()
			{
				document.frmRep.action = "Print2.asp";
				document.frmRep.submit();
			}
			function hideCon()
			{
				if (document.frmRep.SelRep.value == 41 || document.frmRep.SelRep.value == 69)
					{
					document.frmRep.txtCon.style.visibility = 'visible';

					document.frmRep.SelCon.style.visibility = 'visible';
					if(document.frmRep.SelRep.value == 41){
						document.frmRep.SelLog.style.visibility = 'visible';
					}
					else {
						document.frmRep.SelLog.style.visibility = 'hiddem';
					}
				}
				else
					{
					document.frmRep.txtCon.style.visibility = 'hidden';
					document.frmRep.SelCon.style.visibility = 'hidden';
					document.frmRep.SelLog.style.visibility = 'hidden';
					}
			}
			function hideWor()
			{
				if (document.frmRep.SelRep.value == 47 || document.frmRep.SelRep.value == 52)
					{
					document.frmRep.txtWor.style.visibility = 'visible';
					document.frmRep.SelWor.style.visibility = 'visible';
					if (document.frmRep.SelRep.value == 52)
					{
						document.frmRep.SelLog2.style.visibility = 'hidden';
					}
					else
					{
						document.frmRep.SelLog2.style.visibility = 'visible';
					}
					}
				else
					{
					document.frmRep.txtWor.style.visibility = 'hidden';
					document.frmRep.SelWor.style.visibility = 'hidden';
					document.frmRep.SelLog2.style.visibility = 'hidden';
					}
			}
		</script>
		<style>
			Input.btn{
			font-size: 7.5pt;
			font-family: arial;
			color:#000000;
			font-weight:bolder;
			background-color:#d4d0c8;
			border:2px solid;
			text-align: center;
			border-top-color:#d4d0c8;
			border-left-color:#d4d0c8;
			border-right-color:#b6b3ae;
			border-bottom-color:#b6b3ae;
			filter:progid:DXImageTransform.Microsoft.Gradient(GradientType=0,StartColorStr='#ffffffff',EndColorStr='#d4d0c8');
		}
		INPUT.hovbtn{
			font-size: 7.5pt;
			font-family: arial;
			color:#000000;
			font-weight:bolder;
			background-color:#b6b3ae;
			border:2px solid;
			text-align: center;
			border-top-color:#b6b3ae;
			border-left-color:#b6b3ae;
			border-right-color:#d4d0c8;
			border-bottom-color:#d4d0c8;
			filter:progid:DXImageTransform.Microsoft.Gradient(GradientType=0,StartColorStr='#ffffffff',EndColorStr='#b6b3ae');
		}  
		</style>
	</head>
	<body bgcolor='white'  LEFTMARGIN='0' TOPMARGIN='0' onload='hidetxt(); hided8();hideCon(); hideWor();'>
		<!-- #include file="_boxup.asp" -->
		<!-- #include file="_NavHeader.asp" -->

	
		<form method='post' name='frmRep' action="SpecRep.asp">
			<br><br>
			<table cellSpacing='0' cellPadding='0' align='center' border='0'>
				<tr><td colspan='4' align='center'>
				<a href='Report.asp' style='text-decoration: none;'><font size='1' color='blue' face='trebuchet MS'>[General Reports]&nbsp;</font></a>
			
				<font size='2' face='trebuchet MS'>[Advance Reports]</font>&nbsp;&nbsp;
			</td></tr>
				<tr>
					<td align='center'>&nbsp;</font></td>
					<td colspan='3' align='center'>
						<font size='1' face='trebuchet MS'>Type:</font>
						<select name='SelRep'  onchange='hidetxt(); hided8(); hideCon(); hideWor();' >
							<option value='2' <%=Sel2%>>Active Consumers</option>
							<option value='3' <%=Sel3%>>Active PCSP Workers</option>
							<option value='65' <%=Sel65%>>Active Consumers with Case Management Co.</option>
							<option value='36' <%=Sel36%>>Case Manager List</option>
							<option value='14' <%=Sel14%>>Consumers by RIHCC And Town</option>
							<option value='6' <%=Sel6%>>Consumers by Town</option>
							<option value='38' <%=Sel38%>>Consumers Current Care Plan</option>
							<option value='19' <%=Sel19%>>Consumer Date Of Birth</option>
							<option value='48' <%=Sel48%>>Consumer Hours Billed</option>
							<option value='41' <%=Sel41%>>Consumer Logs</option>
							<option value='54' <%=Sel54%>>Consumer Mileage Cap</option>
							<option value='45' <%=Sel45%>>Consumer On Hold</option>
							<option value='29' <%=Sel29%>>Consumer Start and Ammendment Expiration Date</option>
							<option value='40' <%=Sel40%>>Consumer Start and Inactive Date</option>
							<option value='46' <%=Sel46%>>Consumer with No PCSP Worker</option>
							<option value='61' <%=Sel61%>>Consumer with PCSP Worker</option>
							<option value='49' <%=Sel49%>>Finance Consumer List</option>
							<option value='4' <%=Sel4%>>Inactive Consumers with PCSP Worker and Hours</option>
							<option value='67' <%=Sel67%>>Newsletter Labels</option>
							<option value='33' <%=Sel33%>>PCSP Worker by Drivers License Expiration Date</option>
							<option value='15' <%=Sel15%>>PCSP Worker by Insurance Expiration Date</option>
							<option value='20' <%=Sel20%>>PCSP Workers by RIHCC And Town</option>
							<option value='57' <%=Sel57%>>PCSP Workers Contact Info</option>
							<option value='8' <%=Sel8%>>PCSP Workers Date of Hire</option>
							<option value='58' <%=Sel58%>>PCSP Workers Driver Info</option>
							<option value='30' <%=Sel30%>>PCSP Workers Extended Hours</option>
							<option value='9' <%=Sel9%>>PCSP Workers Interested in More Consumers</option>
							<option value='47' <%=Sel47%>>PCSP Workers Logs</option>
							<option value='64' <%=Sel64%>>PCSP Workers Overage Hours</option>
							<option value='50' <%=Sel50%>>PCSP Workers Over 40 Hours (1 week)</option>
							<option value='44' <%=Sel44%>>PCSP Workers Over 80 Hours (2 weeks)</option>
							<option value='24' <%=Sel24%>>PCSP Workers Total Hours</option>
							<option value='59' <%=Sel59%>>PCSP Workers Total Hours (detailed)</option>
							<option value='63' <%=Sel63%>>PCSP Workers Training Logs</option>
							<option value='52' <%=Sel52%>>PCSP Workers Violations</option>
							<option value='10' <%=Sel10%>>PCSP Workers with No Consumer</option>
							<option value='39' <%=Sel39%>>PCSP Workers with Unsubmitted Timesheets</option>
							<option value='11' <%=Sel11%>>Phone Contact for Consumers</option>
							<option value='22' <%=Sel22%>>Phone Contact for PCSP Worker</option>
							<option value='51' <%=Sel51%>>Private Pay Consumers (with Date Range)</option>
							<option value='56' <%=Sel56%>>Private Pay Consumers</option>
							<option value='62' <%=Sel62%>>Private Pay Eligible Worker</option>
							<option value='12' <%=Sel12%>>Referrals</option>
							<option value='35' <%=Sel35%>>Representative List</option>
							<option value='13' <%=Sel13%>>Site Visit for Consumers</option>
							<option value='23' <%=Sel23%>>Site Visit for PCSP Worker</option>
							<option value='27' <%=Sel27%>>* History -  Timesheet / Medicaid </option>
							<option value='31' <%=Sel31%>>* Date Range - Payroll / Medicaid</option>
							<option value='32' <%=Sel32%>>* Date Range - Extended Hours</option>
							<option value='55' <%=Sel55%>>* Date Range - Mileage</option>
							<option value='60' <%=Sel60%>>* Date Range - Total Hours (PCSP Worker/Consumer)</option>
							<option value='66' <%=Sel66%>>* Process Items (Simulation)</option>
							<option value='68' <%=Sel68%>>* Unapproved Caller ID</option>
							<option value='69' <%=Sel69%>>* Medicaid Consumer Record</option>
						</select>
					
						<input type='button' value='Generate Report' class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='document.frmRep.submit();'>
					</td>
				</tr>
				<tr><td>&nbsp;</td></tr>
				<tr>
					<td align='center' colspan='4'>
						<input name='txtFrm' style='width: 40px; border: none;' readonly value='From:'>
						<input name='FrmD8' maxlength='10'>
						<input name='txtTo' style='width: 25px; border: none;' readonly value='To:'>
						<input name='ToD8' maxlength='10'>
						<select name='seltype2'>
							<option value='1' <%=SelPay%>>Payroll</option>
							<option value='2' <%=SelMed%>>Medicaid</option>
							<option value='3' <%=SelOthers%>>Private Pay/Contract/Admin</option>
						</select>
						<select name='seltype3'>
							<option value='1' <%=myPCSP%>>Worker</option>
							<option value='2' <%=myCon%>>Consumer</option>
						</select>
						<select name='seltype4'>
							<option value='2' <%=SelMed%>>Medicaid</option>
							<option value='3' <%=SelOthers%>>Private Pay/Contract/Admin</option>
						</select>
					</td>
				</tr>
				<tr>
					<td>&nbsp;</td>
					<td align='right'>
						<select name='selopt'>
							<option value='0' <%=op1%> >1</option>
							<option value='1' <%=op2%> >2</option>
							<option value='2' <%=op3%> >3</option>
						</select>
						<input name='txtPP' style='width: 80px; border: none;' readonly value='Pay Period/s'>&nbsp;&nbsp;
						<input name='txtCal' style='width: 40px; border: none;' readonly value='From:'>
						<input tabindex="1" name='closedate' style='width:80px;' value='<%=sunDATE%>'
						type="text" onchange='weeknum();' readonly><input tabindex="2" type="button" value="..." name="cal1" style="width: 15px;"
						onclick="showCalendarControl(document.frmRep.closedate);" class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'"> &nbsp;
						<input name='txtTCal' style='width: 40px; border: none;' readonly value='To:'>
						<input tabindex="1" name='Todate' style='width:80px;' readonly value='<%=satDATE%>'
						type="text">&nbsp;&nbsp;
						<select name='selUri'>
							<option value='0' <%=uri1%> >Total</option>
							<option value='1' <%=uri2%> >Detailed</option>
						</select>
						&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
					</td>
					<td>&nbsp;</td>
					<td>
						<select name='seltype'>
							<option value='1' <%=SelPay%>>Payroll</option>
							<option value='2' <%=SelMed%>>Medicaid</option>
						</select>
					</td>
				</tr>
				<tr>
					<td>&nbsp;</td>
					<td align='right'>
						<input name='txtCon' style='width: 70px; border: none;' readonly value='Consumer:'>
						<select name='SelCon'>
							<option value='0'>&nbsp;---All---&nbsp;</option>
							<%=strCON%>
						</select>
						&nbsp;&nbsp;
						<select name='SelLog'>
							<option value='0'>&nbsp;---All---&nbsp;</option>
							<option value='1'>Site Visit</option>
							<option value='2'>Phone Call</option>
							<option value='3'>Misc. Contact</option>
						</select>
					</td>
				</tr>
				<tr>
					<td>&nbsp;</td>
					<td align='right'>
						<input name='txtWor' style='width: 70px; border: none;' readonly value='Worker:'>
						<select name='SelWor'>
							<option value='0'>&nbsp;---All---&nbsp;</option>
							<%=strWOR%>
						</select>
						&nbsp;&nbsp;
						<select name='SelLog2'>
							<option value='0'>&nbsp;---All---&nbsp;</option>
							<option value='1'>Site Visit</option>
							<option value='2'>Phone Call</option>
							<option value='3'>Misc. Contact</option>
						</select>
					</td>
				</tr>
				<tr><td>&nbsp;</td></tr>
				<tr><td colspan='4' align='center'><font color='red' face='trebuchet MS' size='1'>&nbsp;<%=Session("MSG")%>&nbsp;</font></td></tr>
			</table>
			<br>
							<center>
			<% If strBODY <> ""  Or strTBL <> "" Or strProcBX <> "" Or strProcBX2 <> "" Then%>

				<% If Request("SelRep") = 11 Or Request("SelRep") = 13 Or Request("SelRep") = 22 Or Request("SelRep") = 23 Then%>
					<input  style='width: 140px;' align='center' type='button' class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" value='Print Preview' onclick='JavaScript: LandWarn();'>
				<% ElseIf Request("SelRep") = 69 Then%>
					<input  style='width: 140px;' align='center' type='button' class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" value='Print Preview' onclick='PrintPrev2();'>
				<% Else %>
					<input  style='width: 140px;' align='center' type='button' class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" value='Print Preview' onclick='PrintPrev();'>
				<% End If%>
				<% If Request("SelRep") = 19 Or Request("SelRep") = 28 Or Request("SelRep") = 2 Or Request("SelRep") = 3 Or Request("SelRep") = 35 Or Request("SelRep") = 14 Or Request("SelRep") = 20 Or Request("SelRep") = 36 Or Request("SelRep") = 55 Or Request("SelRep") = 62 Or Request("SelRep") = 60 Or Request("SelRep") = 66 Or Request("SelRep") = 67 Then%>
					<input  style='width: 140px;' align='center' type='button' class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" value='Export to CSV' onclick='JavaScript: ExpCSV(1);'>
				<% End If %>
				<% If Request("SelRep") = 55 And Request("selUri") = 0 And (session("UserID") = 893 Or session("UserID") = 67 Or session("UserID") = 2) Then%>
					<input  style='width: 140px;' align='center' type='button' class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" value='Export to PRFJDEPI' onclick='JavaScript: ExpCSV(2);'>
				<% End If %>
		<% End If%>
			<br><br>

				<table cellSpacing='0' cellPadding='0' align='center' border='1'>
					<%=strHEAD%>
					<%=strBODY%>
					<% If strHEAD2 <> "" Then %>
						<tr><td colspan='11'>&nbsp;</td></tr>
					 
						<%=strHEAD2%>
						<%=strBODY2%>
					<% end If %>
				</table>

		</center>
			<input type='hidden' name='fdate' value='<%=sunDATE%>'>
			<input type='hidden' name='tdate' value='<%=satDATE%>'>
			<input type='hidden' name='myUri' value='<%=myUri%>'>
			<input type='hidden' name='seltype32' value='<%=Request("seltype3")%>'>
			<input type='hidden' name='FrmD82' value='<%=Request("FrmD8")%>'>
			<input type='hidden' name='ToD82' value='<%=Request("ToD8")%>'>
			<input type='hidden' name='seltype42' value='<%=Request("seltype4")%>'>
			</form>
		<!-- #include file="_boxdown.asp" -->
	</body>
</html>
<%
Session("MSG") = ""
%>