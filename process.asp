<%Language=VBScript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_security.asp" -->
<%
If UCase(Session("lngType")) = "0" Then
	Session("MSG") = "Invalid User Type. Please Sign In again."
	Response.Redirect "default.asp"
End If
Dim tmpDates(), tmpIDs(), tmpHrs(), tmpName(), tmpEXT(), tmpAmount(), tmpNotes(), tmpWorID(), tmpMax(), tmpMileCap(), tmpCap(), tmpMileOnly()
Dim tmpDates2(), tmpIDs2(), tmpHrs2(), tmpName2(), tmpEXT2(), tmpAmount2(), tmpNotes2(), tmpWorID2(), tmpMax2(),tmpMileCap2(), tmpCap2(), tmpMileOnly2()
Dim tmpamile(), tmpamile2(), tmpHhrs(), tmpHhrs2(), tmpPTO(), tmpPTO2(), tmpEmpID(), tmpPMID(), tmpPMID2(), tmpactcode(), tmpactcode2()

'FUNCTIONS
Function GetSunwk1(xxx)
difwk = DateDiff("ww", wk1, xxx)
		myDATE = xxx
    If difwk >= 0 Then
        wknum = difwk + 1
        If Z_IsOdd2(wknum) Then
            If WeekdayName(Weekday(myDATE), True) = "Sun" Then
                sunDATE = myDATE
                If selopt = 0 Then
                	satDATE = DateAdd("d", 13, sunDATE)
            		ElseIf selopt = 1 Then
            			satDATE = DateAdd("d", 27, sunDATE)
            		ElseIf selopt = 2 Then 	
            			satDATE = DateAdd("d", 41, sunDATE)
            		End If
            Else
                difDATE = DatePart("w", myDATE)
                sunDATE = DateAdd("d", -CInt(difDATE - 1), myDATE)
                If selopt = 0 Then
                	satDATE = DateAdd("d", 13, sunDATE)
            		ElseIf selopt = 1 Then
            			satDATE = DateAdd("d", 27, sunDATE)
            		ElseIf selopt = 2 Then 	
            			satDATE = DateAdd("d", 41, sunDATE)
            		End If
            End If
        Else
            If WeekdayName(Weekday(myDATE), True) = "Sun" Then
                sunDATE = DateAdd("d", -7, myDATE)
               If selopt = 0 Then
                	satDATE = DateAdd("d", 13, sunDATE)
            		ElseIf selopt = 1 Then
            			satDATE = DateAdd("d", 27, sunDATE)
            		ElseIf selopt = 2 Then 	
            			satDATE = DateAdd("d", 41, sunDATE)
            		End If
            Else
                difDATE = DatePart("w", myDATE)
                sunDATE = DateAdd("d", -CInt(difDATE + 6), myDATE)
                If selopt = 0 Then
                	satDATE = DateAdd("d", 13, sunDATE)
            		ElseIf selopt = 1 Then
            			satDATE = DateAdd("d", 27, sunDATE)
            		ElseIf selopt = 2 Then 	
            			satDATE = DateAdd("d", 41, sunDATE)
            		End If
            End If
        End If
    Else
        wknum = difwk
        If Not Z_IsOdd2(wknum) Then
            If WeekdayName(Weekday(myDATE), True) = "Sun" Then
                sunDATE = myDATE
                If selopt = 0 Then
                	satDATE = DateAdd("d", 13, sunDATE)
            		ElseIf selopt = 1 Then
            			satDATE = DateAdd("d", 27, sunDATE)
            		ElseIf selopt = 2 Then 	
            			satDATE = DateAdd("d", 41, sunDATE)
            		End If
            Else
                difDATE = DatePart("w", myDATE)
                sunDATE = DateAdd("d", -CInt(difDATE - 1), myDATE)
                If selopt = 0 Then
                	satDATE = DateAdd("d", 13, sunDATE)
            		ElseIf selopt = 1 Then
            			satDATE = DateAdd("d", 27, sunDATE)
            		ElseIf selopt = 2 Then 	
            			satDATE = DateAdd("d", 41, sunDATE)
            		End If
            End If
        Else
            If WeekdayName(Weekday(myDATE), True) = "Sun" Then
                sunDATE = DateAdd("d", -7, myDATE)
               If selopt = 0 Then
                	satDATE = DateAdd("d", 13, sunDATE)
            		ElseIf selopt = 1 Then
            			satDATE = DateAdd("d", 27, sunDATE)
            		ElseIf selopt = 2 Then 	
            			satDATE = DateAdd("d", 41, sunDATE)
            		End If
            Else
                difDATE = DatePart("w", myDATE)
                sunDATE = DateAdd("d", -CInt(difDATE + 6), myDATE)
               If selopt = 0 Then
                	satDATE = DateAdd("d", 13, sunDATE)
            		ElseIf selopt = 1 Then
            			satDATE = DateAdd("d", 27, sunDATE)
            		ElseIf selopt = 2 Then 	
            			satDATE = DateAdd("d", 41, sunDATE)
            		End If
            End If
        End If
    End If
    GetSunwk1 = sundate
End Function
Function GetSatwk2(xxx)
difwk = DateDiff("ww", wk1, xxx)
		myDATE = xxx
    If difwk >= 0 Then
        wknum = difwk + 1
        If Z_IsOdd2(wknum) Then
            If WeekdayName(Weekday(myDATE), True) = "Sun" Then
                sunDATE = myDATE
                If selopt = 0 Then
                	satDATE = DateAdd("d", 13, sunDATE)
            		ElseIf selopt = 1 Then
            			satDATE = DateAdd("d", 27, sunDATE)
            		ElseIf selopt = 2 Then 	
            			satDATE = DateAdd("d", 41, sunDATE)
            		End If
            Else
                difDATE = DatePart("w", myDATE)
                sunDATE = DateAdd("d", -CInt(difDATE - 1), myDATE)
                If selopt = 0 Then
                	satDATE = DateAdd("d", 13, sunDATE)
            		ElseIf selopt = 1 Then
            			satDATE = DateAdd("d", 27, sunDATE)
            		ElseIf selopt = 2 Then 	
            			satDATE = DateAdd("d", 41, sunDATE)
            		End If
            End If
        Else
            If WeekdayName(Weekday(myDATE), True) = "Sun" Then
                sunDATE = DateAdd("d", -7, myDATE)
               If selopt = 0 Then
                	satDATE = DateAdd("d", 13, sunDATE)
            		ElseIf selopt = 1 Then
            			satDATE = DateAdd("d", 27, sunDATE)
            		ElseIf selopt = 2 Then 	
            			satDATE = DateAdd("d", 41, sunDATE)
            		End If
            Else
                difDATE = DatePart("w", myDATE)
                sunDATE = DateAdd("d", -CInt(difDATE + 6), myDATE)
                If selopt = 0 Then
                	satDATE = DateAdd("d", 13, sunDATE)
            		ElseIf selopt = 1 Then
            			satDATE = DateAdd("d", 27, sunDATE)
            		ElseIf selopt = 2 Then 	
            			satDATE = DateAdd("d", 41, sunDATE)
            		End If
            End If
        End If
    Else
        wknum = difwk
        If Not Z_IsOdd2(wknum) Then
            If WeekdayName(Weekday(myDATE), True) = "Sun" Then
                sunDATE = myDATE
                If selopt = 0 Then
                	satDATE = DateAdd("d", 13, sunDATE)
            		ElseIf selopt = 1 Then
            			satDATE = DateAdd("d", 27, sunDATE)
            		ElseIf selopt = 2 Then 	
            			satDATE = DateAdd("d", 41, sunDATE)
            		End If
            Else
                difDATE = DatePart("w", myDATE)
                sunDATE = DateAdd("d", -CInt(difDATE - 1), myDATE)
                If selopt = 0 Then
                	satDATE = DateAdd("d", 13, sunDATE)
            		ElseIf selopt = 1 Then
            			satDATE = DateAdd("d", 27, sunDATE)
            		ElseIf selopt = 2 Then 	
            			satDATE = DateAdd("d", 41, sunDATE)
            		End If
            End If
        Else
            If WeekdayName(Weekday(myDATE), True) = "Sun" Then
                sunDATE = DateAdd("d", -7, myDATE)
               If selopt = 0 Then
                	satDATE = DateAdd("d", 13, sunDATE)
            		ElseIf selopt = 1 Then
            			satDATE = DateAdd("d", 27, sunDATE)
            		ElseIf selopt = 2 Then 	
            			satDATE = DateAdd("d", 41, sunDATE)
            		End If
            Else
                difDATE = DatePart("w", myDATE)
                sunDATE = DateAdd("d", -CInt(difDATE + 6), myDATE)
               If selopt = 0 Then
                	satDATE = DateAdd("d", 13, sunDATE)
            		ElseIf selopt = 1 Then
            			satDATE = DateAdd("d", 27, sunDATE)
            		ElseIf selopt = 2 Then 	
            			satDATE = DateAdd("d", 41, sunDATE)
            		End If
            End If
        End If
    End If
    GetSatwk2 = satdate
End Function
'GET MILEAGE RATE
Function GetMRate(tmpCDate, tmpTMile)
	If Z_CZero(tmpTMile) = 0 Then
	  GetMRate = 0
	  Exit Function
	End If
	Set rsRateX = Server.CreateObject("ADODB.RecordSet")
	sqlRateX = "SELECT miledate, milerate FROM MileRate_T WHERE miledate <= '" & tmpCDate & "' ORDER BY miledate DESC"
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
'GET RATE FUNCTION
Function GetRateHM(tmpRdate, tmpRhrs, tmpDay)
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
	sqlRateX = "SELECT VHMrate, VHMdate FROM Rate_T WHERE VHMdate <= '" & tmpComDate & "' ORDER BY VHMdate DESC"
	rsRateX.Open sqlRateX, g_strCONN, 3, 1
	If Not rsRateX.EOF Then
		tmpCurrRate = rsRateX("VHMrate")
		tmpCurrDate = rsRateX("VHMdate")
	End If
	rsRateX.Close
	Set rsRateX = Nothing
	'COMPUTE FOR UNIT
	tmpUnit = tmpRhrs * 4
	'COMPUTE FOR RATE
	GetRateHM = tmpUnit * tmpCurrRate
End Function
Function GetRateHA(tmpRdate, tmpRhrs, tmpDay)
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
	sqlRateX = "SELECT VHArate, VHAdate FROM Rate_T WHERE VHAdate <= '" & tmpComDate & "' ORDER BY VHAdate DESC"
	rsRateX.Open sqlRateX, g_strCONN, 3, 1
	If Not rsRateX.EOF Then
		tmpCurrRate = rsRateX("VHArate")
		tmpCurrDate = rsRateX("VHAdate")
	End If
	rsRateX.Close
	Set rsRateX = Nothing
	'COMPUTE FOR UNIT
	tmpUnit = tmpRhrs * 4
	'COMPUTE FOR RATE
	GetRateHA = tmpUnit * tmpCurrRate
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
	sqlRateX = "SELECT rate, rDate FROM Rate_T WHERE rDate <= '" & tmpComDate & "' ORDER BY rDate DESC"
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
	sqlRate = "SELECT rate FROM Consumer_T WHERE medicaid_number = '" & PID & "'"
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
	sqlRate = "SELECT rate FROM Consumer_T WHERE medicaid_number = '" & xxx & "'"
	rsRate.Open sqlRate, g_strCONN, 1, 3
	If Not rsRate.EOF Then
		GetPRate2 = rsRate("rate")
	End If
	rsRate.Close
	Set rsRate = Nothing
End Function
Function GetFileNum(xxx)
	GetFileNum = "*" & xxx
	Set rsFN = Server.CreateObject("ADODB.RecordSet")
	sqlFN = "SELECT FileNum FROM worker_T WHERE social_security_number = '" & xxx & "'"
	rsFN.Open sqlFN, g_strCONN, 3, 1
	If Not rsFN.EOF Then
		If rsFN("FileNum") <> "" Then GetFileNum = rsFN("FileNum")
	End If
	rsFN.Close
	Set rsFN = Nothing
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
Function SearchArrays(strWk, strID, strWID, tmpDates, tmpIDs, tmpWorID)
	' Returns a number (long) equal to the index where
	' strWk and strID is found in the arrays tmpDates
	' and tmpIDs
	' Returns -1 if not found.
	'
	' strWk and strID are strings;
	' tmpDates and tmpIDs are arrays of strings
	
	DIM	lngMax, lngI
	
	' START: validity check
	SearchArrays = -1
	On Error Resume Next	
	lngMax = UBound(tmpDates)
	If Err.Number <> 0 Then Exit Function
	
	For lngI = 0 to lngMax
		If strWID = "" Then
			If tmpDates(lngI) = strWk And tmpIDs(lngI) = strID Then Exit For
		Else
			If tmpDates(lngI) = strWk And tmpIDs(lngI) = strID And tmpWorID(lngI) = srtWID Then Exit For
		End If
	Next
	If lngI > lngMax Then Exit Function
	SearchArrays = lngI
End Function
Function SearchArrays2(strWk, strID, strWID, tmpDates2, tmpIDs2, tmpWorID2)
	' Returns a number (long) equal to the index where
	' strWk and strID is found in the arrays tmpDates
	' and tmpIDs
	' Returns -1 if not found.
	'
	' strWk and strID are strings;
	' tmpDates and tmpIDs are arrays of strings
	
	DIM	lngMax, lngI
	
	' START: validity check
	SearchArrays2 = -1
	On Error Resume Next	
	lngMax = UBound(tmpDates2)
	If Err.Number <> 0 Then Exit Function
	
	For lngI = 0 to lngMax
		If strWID = "" Then
			If tmpDates2(lngI) = strWk And tmpIDs2(lngI) = strID Then Exit For
		Else
			If tmpDates2(lngI) = strWk And tmpIDs2(lngI) = strID And tmpWorID2(lngI) = srtWID Then Exit For
		End If
	Next
	If lngI > lngMax Then Exit Function
	SearchArrays2 = lngI
End Function
Function GetName(zzz)
	GetName = "N/A"
	Set rsName = Server.CreateObject("ADODB.RecordSet")
	sqlName = "SELECT Lname, fname FROM Consumer_t WHERE Medicaid_number = '" & zzz & "' "
	rsName.Open sqlName, g_strCONN, 3, 1
	If Not rsName.EOF Then
		GetName = rsName("Lname") & ", " & rsName("fname")
	End If
	rsName.Close
	Set rsName = Nothing
End Function
Function GetNameCSV(zzz)
	GetNameCSV = "N/A"
	Set rsName = Server.CreateObject("ADODB.RecordSet")
	sqlName = "SELECT Lname, fname FROM Consumer_t WHERE Medicaid_number = '" & zzz & "' "
	rsName.Open sqlName, g_strCONN, 3, 1
	If Not rsName.EOF Then
		GetNameCSV = rsName("Lname") & """,""" & rsName("fname")
	End If
	rsName.Close
	Set rsName = Nothing
End Function
Function GetNameWork(zzz)
	Set rsName = Server.CreateObject("ADODB.RecordSet")
	sqlName = "SELECT Lname, fname FROM Worker_T WHERE Social_Security_Number = '" & zzz & "' "
	rsName.Open sqlName, g_strCONN, 3, 1
	If Not rsName.EOF Then
		GetNameWork = rsName("Lname") & ", " & rsName("fname")
	End If
	rsName.Close
	Set rsName = Nothing
End Function
Function GetNameWorkCSV(zzz)
	Set rsName = Server.CreateObject("ADODB.RecordSet")
	sqlName = "SELECT Lname, fname FROM Worker_T WHERE Social_Security_Number = '" & zzz & "' "
	rsName.Open sqlName, g_strCONN, 3, 1
	If Not rsName.EOF Then
		GetNameWorkCSV = rsName("Lname") & """,""" & rsName("fname")
	End If
	rsName.Close
	Set rsName = Nothing
End Function
Function GetMipID(zzz)
	GetMipID = zzz
	'Set rsName = Server.CreateObject("ADODB.RecordSet")
	'sqlName = "SELECT * FROM Worker_t WHERE Social_Security_Number = '" & zzz & "' "
	'rsName.Open sqlName, g_strCONN, 3, 1
	'If Not rsName.EOF Then
	'	GetMipID = UCase(Left(rsName("Lname") & " " & rsName("fname"), 12))
	'End If
	'rsName.Close
	'Set rsName = Nothing
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
Function SearchArrays3(strwk, strEID, strWID, tmpDates, tmpEmpID, tmpWorID)
		DIM	lngMax, lngI
	
	' START: validity check
	SearchArrays3 = -1
	On Error Resume Next	
	lngMax = UBound(tmpDates)
	If Err.Number <> 0 Then Exit Function
	
	For lngI = 0 to lngMax
		If tmpDates(lngI) = strWk And tmpWorID(lngI) = strWID Then Exit For
	Next
	If lngI > lngMax Then Exit Function
	SearchArrays3 = lngI
End Function
Function SearchArrays4(strWk, strID, strWID, stractcode, tmpDates2, tmpIDs2, tmpWorID2, tmpactcode)
	' Returns a number (long) equal to the index where
	' strWk and strID is found in the arrays tmpDates
	' and tmpIDs
	' Returns -1 if not found.
	'
	' strWk and strID are strings;
	' tmpDates and tmpIDs are arrays of strings
	
	DIM	lngMax, lngI
	
	' START: validity check
	SearchArrays4 = -1
	On Error Resume Next	
	lngMax = UBound(tmpDates2)
	If Err.Number <> 0 Then Exit Function
	
	For lngI = 0 to lngMax
		If strWID = "" Then
			If tmpDates2(lngI) = strWk And tmpIDs2(lngI) = strID And tmpactcode(lngI) = stractcode Then Exit For
		Else
			If tmpDates2(lngI) = strWk And tmpIDs2(lngI) = strID And tmpactcode(lngI) = stractcode And tmpWorID2(lngI) = srtWID Then Exit For
		End If
	Next
	If lngI > lngMax Then Exit Function
	SearchArrays4 = lngI
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
	Function FixDateFormat(xxx)
		FixDateFormat = Right("0" & DatePart("m", xxx), 2) & "/" & Right("0" & DatePart("d", xxx), 2) & "/" & Year(xxx)
	End Function	
Function Z_Pcode(xxx)
	Z_Pcode = ""
	Set rsPcode = Server.CreateObject("ADODB.RecordSet")
	rsPcode.Open "SELECT pcode FROM consumer_T WHERE Medicaid_number = '" & xxx & "' ", g_strCONN, 3, 1
	If Not rsPcode.EOF Then
		Z_Pcode = rsPcode("pcode")
	End If
	rsPcode.Close
	Set rsPcode = Nothing
End Function
''''END FUNCTIONS
	Session("PrintPrev") = ""
	Session("PrintPrevRep") = ""
	Session("PrintPrevPRoc") = ""
	If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
		server.scripttimeout = 600000 '10mins
		PDate = Date
		markerX = 0
		If Request("Payd8") <> "" Then 
			If IsDate(Request("Payd8")) Then
				Pdate = Request("Payd8")
			Else
				Session("MSG") = "Enter valid date."
				Response.Redirect = "Process.asp"
			End If
		End If 
		
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
		If Request("type") = 1 Then
			sqlProc = sqlProc & ", worker_t  WHERE emp_id = social_security_number "
		ElseIf Request("type") = 2 Then
			sqlProc = sqlProc & ", consumer_t  WHERE client = medicaid_number AND code = 'M' "
		ElseIf Request("type") = 3 Then
			sqlProc = sqlProc & ", consumer_t  WHERE client = medicaid_number AND (code = 'P' OR code = 'C' OR code = 'A') "
		ElseIf Request("type") = 4 Then
			sqlProc = sqlProc & ", consumer_t  WHERE client = medicaid_number AND code = 'V' "
		End If
		sqlProc = sqlProc & "AND date <= '" & satDATE & "' AND date >= '" & sunDATE & "' AND" 
		mySunDate = sunDATE
		If Request("type") = 1 Then
			sqlProc = sqlProc & " ProcPay IS NULL ORDER BY lname, fname, date"
		ElseIf Request("type") = 2 Then
			sqlProc = sqlProc & " ProcMed IS NULL AND EXT = 0 ORDER BY lname, fname ASC, date DESC"
		ElseIf Request("type") = 3 Then
			sqlProc = sqlProc & " ProcPriv IS NULL AND EXT = 0 ORDER BY lname, fname ASC, date DESC"
		ElseIf Request("type") = 4 Then
			sqlProc = sqlProc & " ProcVA IS NULL AND EXT = 0 ORDER BY lname, fname ASC, date DESC"
		else
			session("msg") = "please choose between medicaid or private pay or VA."
			response.redirect "process.asp"
		End If
		
		'FOR EXPORT
		Session("sqlVar") = Z_DoEncrypt(Pdate & "|" & Request("type") & "|" & sqlProc)
		'response.write sqlproc
		rsProc.Open sqlProc, g_strCONN, 1, 3
		If Not rsProc.EOF Then
			markerX = 1
			If Request("type") = 1 Then
				strProcH = "<tr bgcolor='#040C8B'><td align='center' width='100px'><font size='1' face='trebuchet ms' color='white' color='white'>Timesheet Week</font>" & _
					"</td><td align='center' width='75px'><font size='1' face='trebuchet ms' color='white' color='white'>FileNumber</font></td><td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>" & _
					"Name</font></td><td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>Reg. Hrs.</font></td>" & _
					"<td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>Holiday Hrs.</font></td>" & _
					"<td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>Total Hrs.</font></td>" & _
					"<td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>PTO</font></td>" & _
					"<td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>Mileage</font></td>" & _
					"<td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>Mileage Amt</font></td></tr>"
				strProcHexp = "Timesheet Week, FileNumber, Last Name, First Name, Reg. Hrs., Holiday Hrs., Total Hrs., PTO, Mileage, Mileage Amt"
			ElseIf Request("type") = 2 Then
				strProcH = "<tr bgcolor='#040C8B'><td align='center' width='100px'><font size='1' face='trebuchet ms' color='white' color='white'>Timesheet Week</font>" & _
					"</td><td align='center' width='100px'><font size='1' face='trebuchet ms' color='white' color='white'>Medicaid</font></td><td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>" & _
					"Name</font></td><td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>Hours</font></td><td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>" & _
					"Units</font></td><td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>" & _
					"Amount</font></td><td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>" & _
					"Pcode</font></td></tr>"
				strProcHexp = "Timesheet Week, Medicaid, Last Name, First Name, Hours, Units,Amount,Pcode"
			ElseIf Request("type") = 3 Then
				strProcH = "<tr bgcolor='#040C8B'><td align='center' width='100px'><font size='1' face='trebuchet ms' color='white' color='white'>Timesheet Week</font>" & _
					"</td><td align='center' width='100px'><font size='1' face='trebuchet ms' color='white' color='white'>Medicaid</font></td><td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>" & _
					"Name</font></td><td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>Regular Hours</font></td><td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>" & _
					"Holiday Hours</font></td><td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>" & _
					"Rate</font></td><td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>" & _
					"Regular Amount</font></td><td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>" & _
					"Holiday Amount</font><td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>" & _
					"Total</font></td><td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>" & _
					"Mileage</font></td></td><td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>" & _
					"Pcode</font></td></td></tr>"
				strProcHexp = "Timesheet Week, Medicaid, Last Name, First Name, Regular Hours, Holiday Hours,Rate, Regualr Amount, Holiday Amount, Total, Mileage,Pcode"
			ElseIf Request("type") = 4 Then
				strProcH = "<tr bgcolor='#040C8B'><td align='center' width='100px'><font size='1' face='trebuchet ms' color='white' color='white'>Timesheet Week</font>" & _
					"</td><td align='center' width='100px'><font size='1' face='trebuchet ms' color='white' color='white'>Medicaid</font></td><td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>" & _
					"Name</font></td><td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>Hours</font></td><td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>" & _
					"Units</font></td><td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>" & _
					"Amount</font></td><td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>" & _
					"Pcode</font></td></tr>"
				strProcHexp = "Timesheet Week, Medicaid, Last Name, First Name, Hours, Units,Amount,Pcode"
			End If
			x = 0
			markme = 0
			ctr = 0
			Do Until rsProc.EOF
				If Request("type") = 2 Then	'medicaid
					'LOOK FOR PREVIOUS ITEMS - ADDITONAL HRS
					Set rsEx = Server.CreateObject("ADODB.RecordSet")
					sqlEx = "SELECT * FROM Tsheets_T WHERE client = '" & rsProc("client") & "' AND emp_id = '" & rsProc("emp_id") & _
						"' AND date =  '" & rsProc("date") & "' AND EXT = 0 AND NOT ProcMed IS NULL ORDER BY Date, timestamp"
					rsEx.Open sqlEx, g_strCONN, 3, 1
					If Not rsEx.EOF Then
						Do Until rsEx.EOF
							TotalHrsEx = rsEx("mon") + rsEx("tue") + rsEx("wed") + rsEx("thu") + rsEx("fri") + rsEx("sat") + rsEx("sun")
							'tmpMileAmtEx = GetMRate(rsEx("date"),rsEx("mile"))
							'tmpMileOnlyEx = rsEx("mile")
							tmpUnitEx = TotalHrsEx * 4
							RmonEx = GetRate(rsEx("date"), rsEx("mon"), "MON")
							RtueEx = GetRate(rsEx("date"), rsEx("tue"), "TUE")
							RwedEx = GetRate(rsEx("date"), rsEx("wed"), "WED")
							RthurEx = GetRate(rsEx("date"), rsEx("thu"), "THU")
							RfriEx = GetRate(rsEx("date"), rsEx("fri"), "FRI")
							RsatEx = GetRate(rsEx("date"), rsEx("sat"), "SAT")
							RsunEx = GetRate(rsEx("date"), rsEx("sun"), "SUN")
							RAmountEx = RmonEx + RtueEx + RwedEx + RthurEx + RfriEx + RsatEx + RsunEx '+ tmpMileAmtEx
							RF = "#000000"
							If rsEx("MAX") = True Then RF = "Red"
							'MC = ""
							'If rsEx("milecap") = True Then MC = "**"
							strTBL2 = strTBL2 & "<tr bgcolor = '#FFFCCC'><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & rsEx("date") & _
							" - " & rsEx("date") + 6 & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & rsEx("client") & _
							"</font><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & GetName(rsEx("client")) & _
							"</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & _
							Z_FormatNumber(TotalHrsEx, 2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & _
							Z_FormatNumber(tmpUnitEx, 2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' " & RF & ">$" & _
							Z_FormatNumber(RAmountEx, 2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' " & RF & ">" & _
							Z_Pcode(rsEx("client")) & "</font></td></tr>" & vbCrLf
							
							strTBLexp = strTBLexp & """" & rsEx("date") & " - " & rsEx("date") + 6 & _
										""",""" & "*" & rsEx("client") & _
										""",""" & GetNameCSV(rsEx("client")) & """,""" & _
										Z_FormatNumber(TotalHrsEx, 2) & """,""" & Z_FormatNumber(tmpUnitEx, 2) & """,""$" & Z_FormatNumber(RAmountEx, 2) & _
										""",""" & Z_Pcode(rsEx("client")) & """" & vbCrLf
							
							rsEx.MoveNext
						Loop
						markeme = 1
						'not yet processed
						TotalHrs = rsProc("mon") + rsProc("tue") + rsProc("wed") + rsProc("thu") + rsProc("fri") + rsProc("sat") + rsProc("sun")
						tmpUnit = TotalHrs * 4
						'tmpMileAmt = GetMRate(rsProc("date"),rsProc("mile"))
						'tmpMilesOnly = rsProc("mile")
						Rmon = GetRate(rsProc("date"), rsProc("mon"), "MON")
						Rtue = GetRate(rsProc("date"), rsProc("tue"), "TUE")
						Rwed = GetRate(rsProc("date"), rsProc("wed"), "WED")
						Rthur = GetRate(rsProc("date"), rsProc("thu"), "THU")
						Rfri = GetRate(rsProc("date"), rsProc("fri"), "FRI")
						Rsat = GetRate(rsProc("date"), rsProc("sat"), "SAT")
						Rsun = GetRate(rsProc("date"), rsProc("sun"), "SUN")
						RAmount = Rmon + Rtue + Rwed + Rthur + Rfri + Rsat + Rsun '+ tmpMileAmt
					  RF = "#000000"
						If rsProc("MAX") = True Then RF = "Red"
						'MC = ""
						'If rsProc("milecap") = True Then MC = "**"
						strTBL2 = strTBL2 & "<tr bgcolor ='#FFFFFF'><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & rsProc("date") & _
							" - " & rsProc("date") + 6 & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & rsProc("client") & _
							"</font><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & GetName(rsProc("client")) & _
							"</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & _
							Z_FormatNumber(TotalHrs, 2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & _
							Z_FormatNumber(tmpUnit, 2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>$" & _
							Z_FormatNumber(RAmount, 2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & _
							Z_Pcode(rsProc("client")) & "</font></td></tr><tr><td colspan='9'>&nbsp;</td></tr>" & vbCrLf
								
						strTBLexp = strTBLexp & """" & rsProc("date") & " - " & rsProc("date") + 6 & _
							""",""" & rsProc("client") & _
							""",""" & GetNameCSV(rsProc("client")) & """,""" & _
							Z_FormatNumber(TotalHrs, 2) & """,""" & Z_FormatNumber(tmpUnit, 2) & """,""$" & Z_FormatNumber(RAmount, 2) &  _
							""",""" & Z_Pcode(rsProc("client")) & """" & vbCrLf	
						
						strBIll = strBILL & """" & rsProc("Client") & """,""" & "001" & """,""" & "IHC" & """,""" & "IHC" & """,""" & "Z598" & """,""" & _
								FixDateFormat(sundate) & """,""" & FixDateFormat(satdate) & """,""" & tmpUnit & """,""" & """" 
							
						if Not rsProc.EOF then strBIll = strBILL & vbCrLf
													 
						rsProc("ProcMed") = Date
						rsProc.Update
						ctr =  ctr + 1
					End If
					rsEx.Close
					Set rsEx = Nothing
				ElseIf Request("type") = 3 Then 'private pay
					'LOOK FOR PREVIOUS ITEMS - ADDITONAL HRS
					Set rsEx = Server.CreateObject("ADODB.RecordSet")
					sqlEx = "SELECT * FROM Tsheets_T WHERE client = '" & rsProc("client") & "' AND emp_id = '" & rsProc("emp_id") & _
						"' AND date =  '" & rsProc("date") & "' AND EXT = 0 AND NOT ProcPriv IS NULL ORDER BY Date, timestamp"
					rsEx.Open sqlEx, g_strCONN, 3, 1
					If Not rsEx.EOF Then
						Do Until rsEx.EOF
							'REGULAR HOURS
							TotalHrsEx = rsEx("mon") + rsEx("tue") + rsEx("wed") + rsEx("thu") + rsEx("fri") + rsEx("sat") + rsEx("sun")
							RAmountEx = GetPRate(TotalHrsEx, rsEx("client"), 0)
							'HOLIDAY HOURS
							Hmon = (Hmon) + (GetHoliday(rsEx("date"), rsEx("mon"), "MON"))
							Htue = (Htue) + (GetHoliday(rsEx("date"), rsEx("tue"), "TUE"))
							Hwed = (Hwed) + (GetHoliday(rsEx("date"), rsEx("wed"), "WED"))
							Hthur = (Hthur) + (GetHoliday(rsEx("date"), rsEx("thu"), "THU"))
							Hfri = (Hfri) + (GetHoliday(rsEx("date"), rsEx("fri"), "FRI"))
							Hsat = (Hsat) + (GetHoliday(rsEx("date"), rsEx("sat"), "SAT"))
							Hsun = (Hsun) + (GetHoliday(rsEx("date"), rsEx("sun"), "SUN"))
							tmpHoldidayHrs = Hmon + Htue + Hwed + Hthur + Hfri + Hsat + Hsun
							RAmountEx2 =  GetPRate(tmpHoldidayHrs, rsEx("client"), 1)
							Hmon = 0
							Htue = 0
							Hwed = 0
							Hthur = 0
							Hfri = 0
							Hsat = 0
							Hsun = 0
							RF = "#000000"
							If rsEx("MAX") = True Then RF = "Red"
							'MC = ""
							'If rsEx("milecap") = True Then MC = "**"
							myRate = GetPRate2(rsEx("client"))
							RMileAmt =  rsEx("mile") + rsEx("amile") 'GetMRate(rsEx("date"),rsEx("mile")) + GetMRate(rsEx("date"),rsEx("amile"))
							TotAmt = RAmountEx2 + RAmountEx
							strTBL2 = strTBL2 & "<tr bgcolor = '#FFFCCC'><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & rsEx("date") & _
								" - " & rsEx("date") + 6 & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & rsEx("client") & _
								"</font><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & GetName(rsEx("client")) & _
								"</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & _
								Z_FormatNumber(TotalHrsEx - tmpHoldidayHrs, 2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & _
								Z_FormatNumber(tmpHoldidayHrs, 2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>$" & _
								Z_FormatNumber(myRate, 2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>$" & _
								Z_FormatNumber(RAmountEx, 2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>$" & _
								Z_FormatNumber(RAmountEx2, 2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>$" & _
								Z_FormatNumber(TotAmt, 2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & _
								Z_FormatNumber(RMileAmt, 2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & _
								Z_Pcode(rsEx("client"))  & "</font></td></tr>" & vbCrLf
							
							strTBLexp = strTBLexp & """" & rsEx("date") & " - " & rsEx("date") + 6 & _
										""",""" & "*" & rsEx("client") & _
										""",""" & GetNameCSV(rsEx("client")) & """,""" & _
										Z_FormatNumber(TotalHrsEx - tmpHoldidayHrs, 2) & """,""" & Z_FormatNumber(tmpHoldidayHrs, 2) & """,""$" & Z_FormatNumber(myRate, 2) & """,""$" & Z_FormatNumber(RAmountEx, 2)& _
										""",""$" & Z_FormatNumber(RAmountEx2, 2) & """,""$" & Z_FormatNumber(TotAmt, 2) & """,""" & Z_FormatNumber(RMileAmt, 2) & _
										""",""" & Z_Pcode(rsEx("client")) & """" & vbCrLf
							rsEx.MoveNext
						Loop
						markeme = 1
						'not yet processed
						'REGULAR HOURS
						TotalHrs = rsProc("mon") + rsProc("tue") + rsProc("wed") + rsProc("thu") + rsProc("fri") + rsProc("sat") + rsProc("sun")
						RAmount = GetPRate(TotalHrs, rsProc("client"), 0)
					  'HOLIDAY HOURS
				  	Hmon = (Hmon) + (GetHoliday(rsProc("date"), rsProc("mon"), "MON"))
						Htue = (Htue) + (GetHoliday(rsProc("date"), rsProc("tue"), "TUE"))
						Hwed = (Hwed) + (GetHoliday(rsProc("date"), rsProc("wed"), "WED"))
						Hthur = (Hthur) + (GetHoliday(rsProc("date"), rsProc("thu"), "THU"))
						Hfri = (Hfri) + (GetHoliday(rsProc("date"), rsProc("fri"), "FRI"))
						Hsat = (Hsat) + (GetHoliday(rsProc("date"), rsProc("sat"), "SAT"))
						Hsun = (Hsun) + (GetHoliday(rsProc("date"), rsProc("sun"), "SUN"))
						tmpHoldidayHrsP = Hmon + Htue + Hwed + Hthur + Hfri + Hsat + Hsun
						RAmountP =  GetPRate(tmpHoldidayHrsP, rsProc("client"), 1)
						Hmon = 0
						Htue = 0
						Hwed = 0
						Hthur = 0
						Hfri = 0
						Hsat = 0
						Hsun = 0
					  RF = "#000000"
						If rsProc("MAX") = True Then RF = "Red"
						'MC = ""
						'If rsProc("milecap") = True Then MC = "**"
						myRate = GetPRate2(rsProc("client"))
						RMileAmt = rsProc("mile") + rsProc("amile") 'GetMRate(rsProc("date"),rsProc("mile")) + GetMRate(rsProc("date"),rsProc("amile"))
						TotAmt = RAmountP + RAmount
						strTBL2 = strTBL2 & "<tr bgcolor = '#FFFFFF'><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & rsProc("date") & _
								" - " & rsProc("date") + 6 & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & rsProc("client") & _
								"</font><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & GetName(rsProc("client")) & _
								"</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & _
								Z_FormatNumber(TotalHrs - tmpHoldidayHrsP, 2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & _
								Z_FormatNumber(tmpHoldidayHrsP, 2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>$" & _
								Z_FormatNumber(myRate, 2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>$" & _
								Z_FormatNumber(RAmount, 2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>$" & _
								Z_FormatNumber(RAmountP, 2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>$" & _
								Z_FormatNumber(TotAmt, 2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & _
								Z_FormatNumber(RMileAmt, 2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & _
								Z_Pcode(rsProc("client")) & "</font></td></tr>" & vbCrLf
								
						strTBLexp = strTBLexp & """" & rsProc("date") & " - " & rsProc("date") + 6 & _
								""",""" & rsProc("client") & _
								""",""" & GetNameCSV(rsProc("client")) & """,""" & _
								Z_FormatNumber(TotalHrs - tmpHoldidayHrsP, 2) & """,""" & Z_FormatNumber(tmpHoldidayHrsP, 2) & """,""$" & Z_FormatNumber(myRate, 2) & """,""$" & Z_FormatNumber(RAmount, 2)& _
								""",""$" & Z_FormatNumber(RAmountP, 2) & """,""$" & Z_FormatNumber(TotAmt, 2) & """,""" & Z_FormatNumber(RMileAmt, 2) & _
										""",""" & Z_Pcode(rsProc("client")) & """" & vbCrLf
						
												 
						rsProc("ProcPriv") = Date
						rsProc.Update
						ctr =  ctr + 1
					End If
					rsEx.Close
					Set rsEx = Nothing
				ElseIf Request("type") = 4 Then 'VA
					'LOOK FOR PREVIOUS ITEMS - ADDITONAL HRS
					Set rsEx = Server.CreateObject("ADODB.RecordSet")
					sqlEx = "SELECT * FROM Tsheets_T WHERE client = '" & rsProc("client") & "' AND emp_id = '" & rsProc("emp_id") & _
						"' AND date =  '" & rsProc("date") & "' AND EXT = 0 AND NOT ProcVA IS NULL ORDER BY actcode, Date, timestamp"
					rsEx.Open sqlEx, g_strCONN, 3, 1
					If Not rsEx.EOF Then
						Do Until rsEx.EOF
							TotalHrsEx = rsEx("mon") + rsEx("tue") + rsEx("wed") + rsEx("thu") + rsEx("fri") + rsEx("sat") + rsEx("sun")
							tmpUnitEx = TotalHrsEx * 4
							If Instr(rsEx("misc_notes"), "80,") > 0 Then
								RmonEx = GetRateHM(rsEx("date"), rsEx("mon"), "MON")
								RtueEx = GetRateHM(rsEx("date"), rsEx("tue"), "TUE")
								RwedEx = GetRateHM(rsEx("date"), rsEx("wed"), "WED")
								RthurEx = GetRateHM(rsEx("date"), rsEx("thu"), "THU")
								RfriEx = GetRateHM(rsEx("date"), rsEx("fri"), "FRI")
								RsatEx = GetRateHM(rsEx("date"), rsEx("sat"), "SAT")
								RsunEx = GetRateHM(rsEx("date"), rsEx("sun"), "SUN")
								RAmountEx = RmonEx + RtueEx + RwedEx + RthurEx + RfriEx + RsatEx + RsunEx '+ tmpMileAmtEx
								
							ElseIf Instr(rsEx("misc_notes"), "82,") > 0 Then
								RmonEx = GetRateHA(rsEx("date"), rsEx("mon"), "MON")
								RtueEx = GetRateHA(rsEx("date"), rsEx("tue"), "TUE")
								RwedEx = GetRateHA(rsEx("date"), rsEx("wed"), "WED")
								RthurEx = GetRateHA(rsEx("date"), rsEx("thu"), "THU")
								RfriEx = GetRateHA(rsEx("date"), rsEx("fri"), "FRI")
								RsatEx = GetRateHA(rsEx("date"), rsEx("sat"), "SAT")
								RsunEx = GetRateHA(rsEx("date"), rsEx("sun"), "SUN")
							End If
							RF = "#000000"
							If rsEx("MAX") = True Then RF = "Red"
							strTBL2 = strTBL2 & "<tr bgcolor = '#FFFCCC'><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & rsEx("date") & _
							" - " & rsEx("date") + 6 & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & rsEx("client") & _
							"</font><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & GetName(rsEx("client")) & _
							"</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & _
							Z_FormatNumber(TotalHrsEx, 2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & _
							Z_FormatNumber(tmpUnitEx, 2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' " & RF & ">$" & _
							Z_FormatNumber(RAmountEx, 2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' " & RF & ">" & _
							Z_Pcode(rsEx("client")) & "</font></td></tr>" & vbCrLf
							
							strTBLexp = strTBLexp & """" & rsEx("date") & " - " & rsEx("date") + 6 & _
								""",""" & "*" & rsEx("client") & _
								""",""" & GetNameCSV(rsEx("client")) & """,""" & _
								Z_FormatNumber(TotalHrsEx, 2) & """,""" & Z_FormatNumber(tmpUnitEx, 2) & """,""$" & Z_FormatNumber(RAmountEx, 2) & _
								""",""" & Z_Pcode(rsEx("client")) & """" & vbCrLf
							rsEx.MoveNext
						Loop
						markeme = 1
						'not yet processed
						TotalHrs = rsProc("mon") + rsProc("tue") + rsProc("wed") + rsProc("thu") + rsProc("fri") + rsProc("sat") + rsProc("sun")
						tmpUnit = TotalHrs * 4
						'tmpMileAmt = GetMRate(rsProc("date"),rsProc("mile"))
						'tmpMilesOnly = rsProc("mile")
						If Instr(rsProc("misc_notes"), "80,") > 0 Then
							Rmon = GetRateHM(rsProc("date"), rsProc("mon"), "MON")
							Rtue = GetRateHM(rsProc("date"), rsProc("tue"), "TUE")
							Rwed = GetRateHM(rsProc("date"), rsProc("wed"), "WED")
							Rthur = GetRateHM(rsProc("date"), rsProc("thu"), "THU")
							Rfri = GetRateHM(rsProc("date"), rsProc("fri"), "FRI")
							Rsat = GetRateHM(rsProc("date"), rsProc("sat"), "SAT")
							Rsun = GetRateHM(rsProc("date"), rsProc("sun"), "SUN")
							RAmount = Rmon + Rtue + Rwed + Rthur + Rfri + Rsat + Rsun '+ tmpMileAmt
							mycode = "VA-HM"
						ElseIf Instr(rsEx("misc_notes"), "82,") > 0 Then
							Rmon = GetRateHA(rsProc("date"), rsProc("mon"), "MON")
							Rtue = GetRateHA(rsProc("date"), rsProc("tue"), "TUE")
							Rwed = GetRateHA(rsProc("date"), rsProc("wed"), "WED")
							Rthur = GetRateHA(rsProc("date"), rsProc("thu"), "THU")
							Rfri = GetRateHA(rsProc("date"), rsProc("fri"), "FRI")
							Rsat = GetRateHA(rsProc("date"), rsProc("sat"), "SAT")
							Rsun = GetRateHA(rsProc("date"), rsProc("sun"), "SUN")
							RAmount = Rmon + Rtue + Rwed + Rthur + Rfri + Rsat + Rsun '+ tmpMileAmt
							mycode = "VA-HA"
						End If
					  RF = "#000000"
						If rsProc("MAX") = True Then RF = "Red"
						'MC = ""
						'If rsProc("milecap") = True Then MC = "**"
						strTBL2 = strTBL2 & "<tr bgcolor ='#FFFFFF'><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & rsProc("date") & _
							" - " & rsProc("date") + 6 & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & rsProc("client") & _
							"</font><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & GetName(rsProc("client")) & _
							"</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & _
							Z_FormatNumber(TotalHrs, 2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & _
							Z_FormatNumber(tmpUnit, 2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>$" & _
							Z_FormatNumber(RAmount, 2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & _
							Z_Pcode(rsProc("client")) & "</font></td></tr><tr><td colspan='9'>&nbsp;</td></tr>" & vbCrLf
								
						strTBLexp = strTBLexp & """" & rsProc("date") & " - " & rsProc("date") + 6 & _
							""",""" & rsProc("client") & _
							""",""" & GetNameCSV(rsProc("client")) & """,""" & _
							Z_FormatNumber(TotalHrs, 2) & """,""" & Z_FormatNumber(tmpUnit, 2) & """,""$" & Z_FormatNumber(RAmount, 2) &  _
							""",""" & Z_Pcode(rsProc("client")) & """" & vbCrLf	
						
						strBIll = strBILL & """" & rsProc("Client") & """,""" & "001" & """,""" & "VA" & """,""" & mycode & """,""" & "Z598" & """,""" & _
								FixDateFormat(sundate) & """,""" & FixDateFormat(satdate) & """,""" & tmpUnit & """,""" & """" 
							
						if Not rsProc.EOF then strBIll = strBILL & vbCrLf
													 
						rsProc("ProcVA") = Date
						rsProc.Update
						ctr =  ctr + 1
					End If
					rsEx.Close
					Set rsEx = Nothing
				End If
				rsProc.MoveNext
			Loop
			rsProc.Close
			Set rsProc = Nothing
			'SET VARIABLES TO 0
			Rmon = 0
			Rtue = 0
			Rwed = 0
			Rthur = 0
			Rfri = 0
			Rsat = 0
			Rsun = 0
			Hmon = 0
			Htue = 0
			Hwed = 0
			Hthur = 0
			Hfri = 0
			Hsat = 0
			Hsun = 0
			'LOOK FOR NEW HRS NOT ADDITIONAL
			Set rsProc2 = Server.CreateObject("ADODB.RecordSet")
			rsProc2.Open sqlProc, g_strCONN, 1, 3
			'response.write sqlPROC
			If Not rsProc2.EOF Then
				Do Until rsProc2.EOF
					myDATE2 = rsProc2("date")
					If Request("type") = 1 Then
						strEmpID = rsProc2("emp_id")
						'rsProcX("ProcPay") = Date	
						RAmount = 0
						RMileAmt = 0
						strWorID = ""
						RMileAmt = GetMRate(rsProc2("date"),rsProc2("mile")) + GetMRate(rsProc2("date"),rsProc2("amile"))
						'response.write GetMRate(rsProc2("date"),rsProc2("mile")) & " + " & GetMRate(rsProc2("date"),rsProc2("amile")) & "<br>"
						RMileOnly = Z_CZero(rsProc2("mile")) + Z_CZero(rsProc2("amile"))
						'response.write Z_CZero(rsProc2("mile")) & " + " & Z_CZero(rsProc2("amile")) & "<br>"
						'RaMileAmt = GetMRate(rsProc2("date"),rsProc2("amile"))
						'RaMileOnly = rsProc2("amile")
					ElseIf Request("type") = 2 Then
						strEmpID = rsProc2("client")
						'strWorID = rsProc2("emp_id")
						'RAmount = 0
						'RMileAmt = 0
						Rmon = (Rmon) + (GetRate(myDATE2, rsProc2("mon"), "MON"))
						Rtue = (Rtue) + (GetRate(myDATE2, rsProc2("tue"), "TUE"))
						Rwed = (Rwed) + (GetRate(myDATE2, rsProc2("wed"), "WED"))
						Rthur = (Rthur) + (GetRate(myDATE2, rsProc2("thu"), "THU"))
						Rfri = (Rfri) + (GetRate(myDATE2, rsProc2("fri"), "FRI"))
						Rsat = (Rsat) + (GetRate(myDATE2, rsProc2("sat"), "SAT"))
						Rsun = (Rsun) + (GetRate(myDATE2, rsProc2("sun"), "SUN"))	
						'RMileAmt = GetMRate(rsProc2("date"),rsProc2("mile"))
						'RMileOnly = rsProc2("mile")
						RAmount = (Rmon) + (Rtue) + (Rwed) + (Rthur) + (Rfri) + (Rsat) + (Rsun) '+ RMileAmt
						Rmon = 0
						Rtue = 0
						Rwed = 0
						Rthur = 0
						Rfri = 0
						Rsat = 0
						Rsun = 0
					ElseIf Request("type") = 3 Then
						strEmpID = rsProc2("client")
						'strWorID = rsProc2("emp_id")
						RAmountP = 0
						RMileAmt = 0
						RMileOnly = 0
						TotalHrs = rsProc2("mon") + rsProc2("tue") + rsProc2("wed") + rsProc2("thu") + rsProc2("fri") + rsProc2("sat") + rsProc2("sun")
						RAmount = GetPRate(TotalHrs, rsProc2("client"), 0)
						Hmon = (Hmon) + (GetHoliday(rsProc2("date"), rsProc2("mon"), "MON"))
						Htue = (Htue) + (GetHoliday(rsProc2("date"), rsProc2("tue"), "TUE"))
						Hwed = (Hwed) + (GetHoliday(rsProc2("date"), rsProc2("wed"), "WED"))
						Hthur = (Hthur) + (GetHoliday(rsProc2("date"), rsProc2("thu"), "THU"))
						Hfri = (Hfri) + (GetHoliday(rsProc2("date"), rsProc2("fri"), "FRI"))
						Hsat = (Hsat) + (GetHoliday(rsProc2("date"), rsProc2("sat"), "SAT"))
						Hsun = (Hsun) + (GetHoliday(rsProc2("date"), rsProc2("sun"), "SUN"))
						tmpHoldidayHrs = Hmon + Htue + Hwed + Hthur + Hfri + Hsat + Hsun
						RHAmount =  GetPRate(tmpHoldidayHrs, rsProc2("client"), 1)
						Hmon = 0
						Htue = 0
						Hwed = 0
						Hthur = 0
						Hfri = 0
						Hsat = 0
						Hsun = 0
						'myRate = GetPRate2(rsProc2("client"))
						RMileAmt = GetMRate(rsProc2("date"),rsProc2("mile")) + GetMRate(rsProc2("date"),rsProc2("amile"))
						RMileOnly = Z_CZero(rsProc2("mile")) + Z_CZero(rsProc2("amile"))
					ElseIf Request("type") = 4 Then
						strEmpID = rsProc2("client")
						If Instr(rsProc2("misc_notes"), "80,") > 0 Then
							Rmon = (Rmon) + (GetRateHM(myDATE2, rsProc2("mon"), "MON"))
							Rtue = (Rtue) + (GetRateHM(myDATE2, rsProc2("tue"), "TUE"))
							Rwed = (Rwed) + (GetRateHM(myDATE2, rsProc2("wed"), "WED"))
							Rthur = (Rthur) + (GetRateHM(myDATE2, rsProc2("thu"), "THU"))
							Rfri = (Rfri) + (GetRateHM(myDATE2, rsProc2("fri"), "FRI"))
							Rsat = (Rsat) + (GetRateHM(myDATE2, rsProc2("sat"), "SAT"))
							Rsun = (Rsun) + (GetRateHM(myDATE2, rsProc2("sun"), "SUN"))	
						ElseIf Instr(rsProc2("misc_notes"), "82,") > 0 Then
							Rmon = (Rmon) + (GetRateHA(myDATE2, rsProc2("mon"), "MON"))
							Rtue = (Rtue) + (GetRateHA(myDATE2, rsProc2("tue"), "TUE"))
							Rwed = (Rwed) + (GetRateHA(myDATE2, rsProc2("wed"), "WED"))
							Rthur = (Rthur) + (GetRateHA(myDATE2, rsProc2("thu"), "THU"))
							Rfri = (Rfri) + (GetRateHA(myDATE2, rsProc2("fri"), "FRI"))
							Rsat = (Rsat) + (GetRateHA(myDATE2, rsProc2("sat"), "SAT"))
							Rsun = (Rsun) + (GetRateHA(myDATE2, rsProc2("sun"), "SUN"))	
						End If
						RAmount = (Rmon) + (Rtue) + (Rwed) + (Rthur) + (Rfri) + (Rsat) + (Rsun) '+ RMileAmt
						Rmon = 0
						Rtue = 0
						Rwed = 0
						Rthur = 0
						Rfri = 0
						Rsat = 0
						Rsun = 0
					End If
					If Not IsNull(rsProc2("lname")) Then
						strName = Replace(rsProc2("lname"),",","") & ", " & rsProc2("fname")
					Else
						strName = rsProc2("lname") & ", " & rsProc2("fname")
					End If
					dblHours = rsProc2("mon") + rsProc2("tue") + rsProc2("wed") + rsProc2("thu") + rsProc2("fri") + rsProc2("sat") + rsProc2("sun")
					
					If Request("type") = 1 Then
						'HOLIDAY HRS
						
						Hmon = (Hmon) + (GetHoliday(myDATE2, rsProc2("mon"), "MON"))
						Htue = (Htue) + (GetHoliday(myDATE2, rsProc2("tue"), "TUE"))
						Hwed = (Hwed) + (GetHoliday(myDATE2, rsProc2("wed"), "WED"))
						Hthur = (Hthur) + (GetHoliday(myDATE2, rsProc2("thu"), "THU"))
						Hfri = (Hfri) + (GetHoliday(myDATE2, rsProc2("fri"), "FRI"))
						Hsat = (Hsat) + (GetHoliday(myDATE2, rsProc2("sat"), "SAT"))
						Hsun = (Hsun) + (GetHoliday(myDATE2, rsProc2("sun"), "SUN"))
						tmpHoldidayHrs = Hmon + Htue + Hwed + Hthur + Hfri + Hsat + Hsun
						Hmon = 0
						Htue = 0
						Hwed = 0
						Hthur = 0
						Hfri = 0
						Hsat = 0
						Hsun = 0
						
						'PTO
						tmpPTOhrs = 0
						Set rsPTO = Server.CreateObject("ADODB.RecordSet")
						sqlPTO = "SELECT * FROM W_PTO_T WHERE WorkerID = '" & strEmpID & "' AND date >= '" & sunDATE & "' AND date <= '" & satDATE & _
							"' AND procitem IS NULL"
							'response.write sqlPTO
						rsPTO.Open sqlPTO, g_strCONN, 1, 3
						Do Until rsPTO.EOF
							tmpPTOhrs = tmpPTOhrs + rsPTO("PTO")
							rsPTO("procitem") = date
							rsPTO.Update
							rsPTO.MoveNext
						Loop
						rsPTO.Close
						Set rsPTO = Nothing
					End If
					strEXT = False
					If dblHours <> 0 And rsProc2("EXT") = True Then strEXT = True
					strMile = RMileAmt
					strCap = False
					If Request("type") = 1 Then
						If rsProc2("milecap") = True Then strCap = True
					Else
						If rsProc2("milecap") = True Then strCap = True
					End If
					strNotes = 	rsProc2("misc_notes") 
					strMax = False
					If rsProc2("MAX") = True Then strMax = True
					stractcode = rsProc2("misc_notes")
					' find the 2-week period
					strWeekLabel = Z_Find2WkPeriod(myDate2)
					' search for it in the arrays
					If Request("type") <> 4 Then
						lngIdx = SearchArrays2(strWeekLabel, strEmpID, strWorID, tmpDates2, tmpIDs2, tmpWorID2)
					Else
						lngIdx = SearchArrays4(strWeekLabel, strEmpID, strWorID, stractcode, tmpDates2, tmpIDs2, tmpWorID2, tmpactcode)
					End If
					If lngIdx < 0 Then ' this is the first time i've encountered the date and id pair, so i make a new entry
						ReDim Preserve tmpDates2(x)
						ReDim Preserve tmpWorID2(x)
						ReDim Preserve tmpIDs2(x)
						ReDim Preserve tmpHrs2(x)
						ReDim Preserve tmpName2(x)
						ReDim Preserve tmpEXT2(x)
						ReDim Preserve tmpAmount2(x)
						ReDim Preserve tmpNotes2(x)
						ReDim Preserve tmpMax2(x)
						Redim Preserve tmpMileCap2(x)
						Redim Preserve tmpCap2(x)
						Redim Preserve tmpMileOnly2(x)
						Redim Preserve tmpHhrs(x)
						ReDim Preserve tmpPTO(x)
						ReDim Preserve tmpActcode(x)

						
						tmpDates2(x) = strWeekLabel
						tmpIDs2(x) = strEmpID
						tmpWorID2(x) = strWorID
						tmpHrs2(x) = dblHours
						tmpName2(x) = strName
						if tmpEXT2(x) = False Then tmpEXT2(x) = strEXT
						tmpAmount2(x) = RAmount
						tmpNotes2(x) = strNotes
						If tmpMax2(x) = False Then tmpMax2(x) = strMax
						tmpMileOnly2(x) = RMileOnly
						tmpMileCap2(x) = strMile
						if tmpCap2(x) = False Then tmpCap2(x) = strCap
						tmpHhrs(x) = tmpHoldidayHrs
						tmpPTO(x) = tmpPTOhrs
						tmpActcode(x) = stractcode
						x = x + 1
					Else
						tmpHhrs(lngIdx) = tmpHhrs(lngIdx) + tmpHoldidayHrs
						tmpMileOnly2(lngIdx) = tmpMileOnly2(lngIdx) + RMileOnly
						tmpHrs2(lngIdx) = tmpHrs2(lngIdx) + dblHours
						tmpAmount2(lngIdx) = tmpAmount2(lngIdx) + RAmount
						tmpMileCap2(lngIdx) = tmpMileCap2(lngIdx) + strMile
						If strNotes <> "" Then tmpNotes2(lngIdx) = tmpNotes2(lngIdx) & "<br>" & strNotes
						If tmpMax2(lngIdx) = False Then tmpMax2(lngIdx) = strMax
						if tmpCap2(lngIdx) = False Then tmpCap2(lngIdx) = strCap
						
					End If
					rsProc2.MoveNext
				Loop
			End If
			rsProc2.Close
			Set rsProc2 = Nothing	
			Session("MSG") = "Records between " & sunDATE & " - " & satDATE & " has been processed for "
			If Request("type") = 1 Then
					Session("MSG") = Session("MSG") & "payroll. <br>* SSN <br>* has extended hours <br>red font - over max hours <br>** over mileage cap"
			ElseIf Request("type") = 2 Then
					Session("MSG") = Session("MSG") & "medicaid.<br>red font - over max hours<br>yellow background - previous entries "
			ElseIf Request("type") = 3 Then
					Session("MSG") = Session("MSG") & "private pay.<br>red font - over max hours<br>yellow background - previous entries "
			ElseIf Request("type") = 4 Then
					Session("MSG") = Session("MSG") & "VA.<br>red font - over max hours<br>yellow background - previous entries "
			End If 
			'PRINT
			If strTBL2 <> "" Then
				strTBL = strTBL2
			End If
				If markerX = 1 Then
					y = 0
					Do Until y = x
						If Request("type") = 1 Then
							IDx = tmpIDs2(y)'Right(tmpIDs2(y), 4)
							strEXT = ""
							If tmpEXT2(y) = True Then strEXT = "*" 
						Else
							strEXT = ""
							IDx = tmpIDs2(y)
						End If 
						'RED FLAG
						Rf = ""	
						If tmpMax2(y) = True Then RF = "Color='Red'"
						'max mile
						MC = ""
						if tmpCap2(y) = true Then MC = "**"
							'response.write tmpCap2(y) & "<br>"
						'CALCULATE BILL TEMP 
						tmpUnit = tmpHrs2(y) * 4
					
						
						If Request("type") = 1 Then
							tmpRegHrs = tmpHrs2(y) - tmpHhrs(y)
							tmpHolHrs = tmpHhrs(y)
							tmpTotHrs = tmpHrs2(y)				
							strTBL = strTBL & "<tr><td align='center'>" & strEXT & "<font size='1' face='trebuchet ms' color='" & RF & "'>" & tmpDates2(y) & _
											"</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & GetFileNum(IDx) & "</font></td>" & _
											"<td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" &tmpName2(y) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & _
											Z_FormatNumber(tmpRegHrs,2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & _
											Z_FormatNumber(tmpHolHrs,2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & _
											Z_FormatNumber(tmpTotHrs,2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & _
											Z_FormatNumber(tmpPTO(y),2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & MC & _
											Z_FormatNumber(tmpMileOnly2(y),2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>$" & _
											Z_FormatNumber(tmpMileCap2(y),2) & "</font></td></tr>" & vbCrLf
											
							strTBLexp = strTBLexp & """" & strEXT & tmpDates2(y) & _
										""",""" & GetFileNum(IDx) & _
										""",""" & GetNameWorkCSV(IDx) & """,""" & Z_FormatNumber(tmpRegHrs,2) & """,""" & Z_FormatNumber(tmpHolHrs,2) & """,""" & _
										Z_FormatNumber(tmpTotHrs,2) & """,""" & Z_FormatNumber(tmpPTO(y),2) & """,""" & Z_FormatNumber(tmpMileOnly2(y),2) & """,""" & Z_FormatNumber(tmpMileCap2(y),2) & """" & vbCrLf
							
							mipID = GetMipID(tmpIDs2(y))			
							strMIPcsv = strMIPcsv & "HTS," & mipID & ",03,R" & vbCrlf & _
								"DTSEARN," & mipID & ",Regular," & Z_FormatNumber(tmpHrs2(y), 2) & ",OQA EX" & vbCrlf & vbCrlf
						ElseIf Request("type") = 2 Then
							'If tmpHrs(y) <> 0 Then
								strTBL = strTBL & "<tr><td align='center'>" & strEXT & "<font size='1' face='trebuchet ms'color=' " & RF & "'>" & tmpDates2(y) & _
												"</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & IDx & "</font></td>" & _
												"<td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" &tmpName2(y) & "</font></td><td align='center'><font size='1' face='trebuchet ms' " & RF & ">" & _
												Z_FormatNumber(tmpHrs2(y),2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & _
												Z_FormatNumber(tmpUnit, 2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>$" & _
												Z_FormatNumber(tmpAmount2(y), 2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & _
												Z_Pcode(IDx) & "</font></td></tr>" & vbCrLf
												
								strTBLexp = strTBLexp & """" & strEXT & tmpDates2(y) & _
											""",""" & IDx & _
											""",""" & GetNameCSV(IDx) & """,""" & _
											tmpHrs2(y) & """,""" & tmpUnit & """,""$" & tmpAmount2(y) & """,""" & Z_Pcode(IDx) & """" & vbCrLf
								
											
								strBIll = strBILL & """" & IDx & """,""" & "001" & """,""" & "IHC" & """,""" & "IHC" & """,""" & "Z598" & """,""" & _
									FixDateFormat(sundate) & """,""" & FixDateFormat(satdate) & """,""" & tmpUnit & """,""" & """" 
									
								if x - 1 <> y then strBIll = strBILL & vbCrLf
							'End If
						ElseIf Request("type") = 3 Then
							tmpRegHrs = tmpHrs2(y) - tmpHhrs(y)
							tmpHolHrs = tmpHhrs(y)
							tmpTotHrs = tmpHrs2(y)	
							myRate = GetPRate2(IDx)
							myAmount = tmpRegHrs * myRate
							myHAmount = tmpHolHrs * myRate * 1.5
							TotAmt = myHAmount + myAmount
							strTBL = strTBL & "<tr><td align='center'>" & strEXT & "<font size='1' face='trebuchet ms' color='" & RF & "'>" & tmpDates2(y) & _
								"</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & IDx & "</font></td>" & _
								"<td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & tmpName2(y) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & _
								Z_FormatNumber(tmpRegHrs,2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & _
								Z_FormatNumber(tmpHolHrs, 2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & _
								Z_FormatNumber(myRate, 2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & _
								Z_FormatNumber(myAmount, 2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & _
								Z_FormatNumber(myHAmount, 2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & _
								Z_FormatNumber(TotAmt, 2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & _
								Z_FormatNumber(tmpMileOnly2(y), 2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & _
								Z_Pcode(IDx) & "</font></td></tr>" & vbCrLf
												
							strTBLexp = strTBLexp & """" & strEXT & tmpDates2(y) & _
								""",""" & IDx & _
								""",""" & GetNameCSV(IDx) & """,""" & _
								Z_FormatNumber(tmpRegHrs, 2) & """,""" & Z_FormatNumber(tmpHolHrs, 2) & """,""$" & Z_FormatNumber(myRate, 2) & """,""$" & _
								Z_FormatNumber(myAmount, 2) & """,""$" & Z_FormatNumber(myHAmount, 2) & """,""$" & _
								Z_FormatNumber(TotAmt, 2) & """,""" & Z_FormatNumber(tmpMileOnly2(y), 2) & """,""" & Z_Pcode(IDx) & """" & vbCrLf '""",""" & tmpNotes2(y) & """" & vbCrLf
						ElseIf Request("type") = 4 Then 'VA
							'If tmpHrs(y) <> 0 Then
								strTBL = strTBL & "<tr><td align='center'>" & strEXT & "<font size='1' face='trebuchet ms'color=' " & RF & "'>" & tmpDates2(y) & _
												"</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & IDx & "</font></td>" & _
												"<td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" &tmpName2(y) & "</font></td><td align='center'><font size='1' face='trebuchet ms' " & RF & ">" & _
												Z_FormatNumber(tmpHrs2(y),2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & _
												Z_FormatNumber(tmpUnit, 2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>$" & _
												Z_FormatNumber(tmpAmount2(y), 2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & _
												Z_Pcode(IDx) & "</font></td></tr>" & vbCrLf
												
								strTBLexp = strTBLexp & """" & strEXT & tmpDates2(y) & _
											""",""" & IDx & _
											""",""" & GetNameCSV(IDx) & """,""" & _
											tmpHrs2(y) & """,""" & tmpUnit & """,""$" & tmpAmount2(y) & """,""" & Z_Pcode(IDx) & """" & vbCrLf
								
								If tmpActcode(y) = 80 Then 
									mycode = "VA-HM"		
								ElseIf tmpActcode(y) = 82 Then
									mycode = "VA-HA"
								End If
								strBIll = strBILL & """" & IDx & """,""" & "001" & """,""" & "VA" & """,""" & mycode & """,""" & "Z598" & """,""" & _
									FixDateFormat(sundate) & """,""" & FixDateFormat(satdate) & """,""" & tmpUnit & """,""" & """" 
									
								if x - 1 <> y then strBIll = strBILL & vbCrLf
						End If
						
											
						
						
						y = y + 1 
					Loop
					If Request("type") = 2 Then 
						strBIll = """" & FixDateFormat(sunDATE) & " - " & FixDateFormat(satDATE) & " for IHC" & """" & vbCrLf & strBILL
					ElseIf Request("type") = 4 Then
						strBIll = """" & FixDateFormat(sunDATE) & " - " & FixDateFormat(satDATE) & " for VA" & """" & vbCrLf & strBILL
					End If
				End If
				
		Else
			'NO RECORDS FOUND
			If Request("type") = 1 Then
					Session("MSG") = "No payroll records found on " & sunDATE & " - " & satDATE & " for processing. <br>* SSN <br>* has extended hours<br>yellow background - previous entries"
			ElseIf Request("type") = 2 Then
					Session("MSG") = "No medicaid records found on " & sunDATE & " - " & satDATE & " for processing.<br>yellow background - previous entries"
			ElseIf Request("type") = 3 Then
					Session("MSG") = "No Private Pay records found on " & sunDATE & " - " & satDATE & " for processing.<br>yellow background - previous entries"
			ElseIf Request("type") = 4 Then
					Session("MSG") = "No VA records found on " & sunDATE & " - " & satDATE & " for processing.<br>yellow background - previous entries"
			End If
		End If
		'LOOK FOR PTO W/O HOURS'''''''''''''''''''''''''''''
		Set rsPTOhrs = Server.CreateObject("ADODB.RecordSet")
		sqlPTOhrs = "SELECT * FROM W_PTO_T, worker_T WHERE Social_Security_Number = WorkerID AND date >= '" & sunDATE & "' " & _
			"AND date <= '" & satDATE & "' AND procitem IS NULL ORDER BY lname, fname, date"
			'response.write sqlPTOhrs
		rsPTOhrs.Open sqlPTOhrs, g_strCONN,1 , 3
		If Not rsPTOhrs.EOF Then
			strTBL = strTBL & "<tr bgcolor='#040C8B'><td colspan='9'><font size='1' face='trebuchet ms' color='white'><b>PTO's w/o hours</b></font></td></tr>"
			strTBL = strTBL & "<tr bgcolor='#040C8B'><td align='center' width='100px'><font size='1' face='trebuchet ms' color='white' color='white'>Timesheet Week</font>" & _
					"</td><td align='center' width='75px'><font size='1' face='trebuchet ms' color='white' color='white'>FileNumber</font></td><td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>" & _
					"Name</font></td><td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>Reg. Hrs.</font></td>" & _
					"<td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>Holiday Hrs.</font></td>" & _
					"<td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>Total Hrs.</font></td>" & _
					"<td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>PTO</font></td>" & _
					"<td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>Mileage</font></td>" & _
					"<td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>Mileage Amt</font></td></tr>"
			strTBLexp = strTBLexp & "PTO's w/o hours" & vbCrLf & "Timesheet Week,FileNumber,Last Name,First Name,Reg. Hrs.,Holiday Hrs.,Total Hrs.,PTO,Mileage,Mileage Amt" & vbCrLf
			Do Until rsPTOhrs.EOF
				
				strTBL = strTBL & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & rsPTOhrs("date") & " - " & DateAdd("d", rsPTOhrs("date"), 6) & _
					"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & GetFileNum(rsPTOhrs("WorkerID")) & "</font></td>" & _
					"<td align='center'><font size='1' face='trebuchet ms'>" & rsPTOhrs("lname") & ", " & rsPTOhrs("fname") & "</font></td>" & _
					"<td align='center'><font size='1' face='trebuchet ms'>0.00</font></td><td align='center'><font size='1' face='trebuchet ms'>0.00</font></td>" & _
					"<td align='center'><font size='1' face='trebuchet ms'>0.00</font></td><td align='center'><font size='1' face='trebuchet ms'>" & _
					Z_FormatNumber(rsPTOhrs("PTO"),2) & "</font></td><td align='center'><font size='1' face='trebuchet ms'>0.00</font></td>" & _
					"<td align='center'><font size='1' face='trebuchet ms'>$0.00</font></td></tr>" & vbCrLf
				strTBLexp = strTBLexp & rsPTOhrs("date") & " - " & DateAdd("d", rsPTOhrs("date"), 6) & "," & GetFileNum(rsPTOhrs("WorkerID")) & _
					"," & rsPTOhrs("lname") & ", " & rsPTOhrs("fname") & "," & "0.00" & "," & "0.00" & "," & "0.00" & _
					"," & Z_FormatNumber(rsPTOhrs("PTO"),2) & "," & "0.00" & "," & "0.00" & vbCrLf
				rsPTOhrs("procitem") = date
				rsPTOhrs.Update
				rsPTOhrs.MoveNext
			Loop
		End If
		rsPTOhrs.Close
		Set rsPTO = Nothing
		'''''''TAG TIMESHEET
		
			Set rsTAG = Server.CreateObject("ADODB.RecordSet")
			rsTAG.Open sqlProc, g_strCONN, 1, 3
			Do until rsTAG.EOF 
				If  Request("type") = 1 Then
					rsTAG("ProcPay") = Date
				ElseIf Request("type") = 2 Then
					rsTAG("ProcMed") = Date
				ElseIf  Request("type") = 3 Then
					rsTAG("ProcPriv") = Date
				ElseIf  Request("type") = 4 Then
					rsTAG("ProcVA") = Date
				End If
				rsTAG.Update	
				rsTAG.MoveNext
			Loop
			rsTAG.Close
			Set rsTAG = Nothing
		
		''''''''''''''''''''''''
		
		
		''''NOT IN 2 WEEK PERIOD 'VA not done
		Set rsProcX = Server.CreateObject("ADODB.RecordSet")
		sqlProcX = "SELECT * FROM tsheets_t"
		If Request("type") = 1 Then
			sqlProcX = sqlProcX & ", worker_t  WHERE date < '" & sunDATE & "' AND emp_id = social_security_number AND ProcPay IS NULL ORDER BY lname, fname, date, emp_id"
		ElseIf Request("type") = 2 Then
			sqlProcX = sqlProcX & ", consumer_t  WHERE code = 'M' AND date < '" & sunDATE & "' AND client = medicaid_number AND ProcMed IS NULL AND EXT = 0 ORDER BY lname, fname, Date ASC, client ASC"
		ElseIf Request("type") = 3 Then
			sqlProcX = sqlProcX & ", consumer_t  WHERE (code = 'P' OR code = 'C' OR code = 'A') AND date < '" & sunDATE & "' AND client = medicaid_number AND ProcPriv IS NULL AND EXT = 0 ORDER BY lname, fname, Date ASC, client ASC"
		ElseIf Request("type") = 4 Then
			sqlProcX = sqlProcX & ", consumer_t  WHERE code = 'V' AND date < '" & sunDATE & "' AND client = medicaid_number AND ProcVA IS NULL AND EXT = 0 ORDER BY lname, fname, Date ASC, client ASC"
		End If
		MarkerX = 0
		rsProcX.Open sqlProcX, g_strCONN, 1, 3
		If Not rsProcX.EOF Then
			'response.write sqlProcX & "<br>"
			markerX = 1
			strMSG = "<tr bgcolor='#040C8B'><td colspan='11'><font size='1' face='trebuchet ms' color='white'><b>Processed items before the set payroll period</b></font></td></tr>"
			strMSGexp = "Processed items before the set payroll period"
			If Request("type") = 1 Then
				strProcHX ="<tr bgcolor='#040C8B'><td align='center' width='100px'><font size='1' face='trebuchet ms' color='white' color='white'>Timesheet Week</font>" & _
					"</td><td align='center' width='75px'><font size='1' face='trebuchet ms' color='white' color='white'>FileNumber</font></td><td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>" & _
					"Name</font></td><td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>Reg. Hrs.</font></td>" & _
					"<td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>Holiday Hrs.</font></td>" & _
					"<td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>Total Hrs.</font></td>" & _
					"<td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>PTO</font></td>" & _
					"<td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>Mileage</font></td>" & _
					"<td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>Mileage Amt</font></td></tr>"
				
				strProcHXexp = "Timesheet Week, FileNumber, Last Name, First Name, Reg. Hrs., Holiday Hrs., Total Hrs., PTO, Mileage, Mileage Amt"
				
			ElseIf Request("type") = 2 Then
				strProcHX = "<tr bgcolor='#040C8B'><td align='center' width='100px'><font size='1' face='trebuchet ms' color='white'>Timesheet Week</font>" & _
				"</td><td align='center' width='100px'><font size='1' face='trebuchet ms' color='white'>Medicaid</font></td><td align='center'><font size='1' face='trebuchet ms' color='white'>" & _
				"Name</font></td><td align='center'><font size='1' face='trebuchet ms' color='white'>Hours</font></td><td align='center'><font size='1' face='trebuchet ms' color='white'>" & _
				"Units</font></td><td align='center'><font size='1' face='trebuchet ms' color='white'>Amount</font></td><td align='center'><font size='1' face='trebuchet ms' color='white'>Pcode</font></td></tr>"
				
				strProcHXexp = "Timesheet Week, Medicaid, Last Name, First Name, Hours, Units, Amount"
				
			ElseIf Request("type") = 3 Then
				strProcHX = "<tr bgcolor='#040C8B'><td align='center' width='100px'><font size='1' face='trebuchet ms' color='white' color='white'>Timesheet Week</font>" & _
					"</td><td align='center' width='100px'><font size='1' face='trebuchet ms' color='white' color='white'>Medicaid</font></td><td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>" & _
					"Name</font></td><td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>Regular Hours</font></td><td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>" & _
					"Holiday Hours</font></td><td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>" & _
					"Rate</font></td><td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>" & _
					"Regular Amount</font></td><td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>" & _
					"Holiday Amount</font><td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>" & _
					"Total</font></td><td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>" & _
					"Mileage</font></td><td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>" & _
					"Pcode</font></td></td></tr>"
			
				strProcHXexp = "Timesheet Week, Medicaid, Last Name, First Name, Regular Hours, Holiday Hours,Rate, Regualr Amount, Holiday Amount, Total, Mileage,Pcode"
			ElseIf Request("type") = 4 Then
				strProcHX = "<tr bgcolor='#040C8B'><td align='center' width='100px'><font size='1' face='trebuchet ms' color='white'>Timesheet Week</font>" & _
				"</td><td align='center' width='100px'><font size='1' face='trebuchet ms' color='white'>Medicaid</font></td><td align='center'><font size='1' face='trebuchet ms' color='white'>" & _
				"Name</font></td><td align='center'><font size='1' face='trebuchet ms' color='white'>Hours</font></td><td align='center'><font size='1' face='trebuchet ms' color='white'>" & _
				"Units</font></td><td align='center'><font size='1' face='trebuchet ms' color='white'>Amount</font></td><td align='center'><font size='1' face='trebuchet ms' color='white'>Pcode</font></td></tr>"
				
				strProcHXexp = "Timesheet Week, Medicaid, Last Name, First Name, Hours, Units, Amount"
			End If
			x = 0
			markme = 0
			Do Until rsProcX.EOF
				If Request("type") = 2 Then	
					'LOOK FOR PREVIOUS ITEMS - ADDITONAL HRS
					Set rsEx = Server.CreateObject("ADODB.RecordSet")
					sqlEx = "SELECT * FROM Tsheets_T WHERE client = '" & rsProcX("client") & "' AND emp_id = '" & rsProcX("emp_id") & _
						"' AND date =  '" & rsProcX("date") & "' AND EXT = 0 AND NOT ProcMed IS NULL ORDER BY Date, timestamp"
						
					rsEx.Open sqlEx, g_strCONN, 3, 1
					If Not rsEx.EOF Then
					
						Do Until rsEx.EOF
							TotalHrsEx = rsEx("mon") + rsEx("tue") + rsEx("wed") + rsEx("thu") + rsEx("fri") + rsEx("sat") + rsEx("sun")
							tmpMileAmtEx = GetMRate(rsEx("date"),rsEx("mile"))
							'tmpMileOnlyEx = rsEx("mile")
							tmpUnitEx = TotalHrsEx * 4
							RmonEx = GetRate(rsEx("date"), rsEx("mon"), "MON")
							RtueEx = GetRate(rsEx("date"), rsEx("tue"), "TUE")
							RwedEx = GetRate(rsEx("date"), rsEx("wed"), "WED")
							RthurEx = GetRate(rsEx("date"), rsEx("thu"), "THU")
							RfriEx = GetRate(rsEx("date"), rsEx("fri"), "FRI")
							RsatEx = GetRate(rsEx("date"), rsEx("sat"), "SAT")
							RsunEx = GetRate(rsEx("date"), rsEx("sun"), "SUN")
							RAmountEx = RmonEx + RtueEx + RwedEx + RthurEx + RfriEx + RsatEx + RsunEx '+ tmpMileAmtEx
							RF = "#000000"
							If rsEx("MAX") = True Then RF = "Red"		
							'MC = ""
							'If rsEx("milecap") = True Then MC = "**"					
							strTBLEx = strTBLEx & "<tr bgcolor = '#FFFCCC'><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & rsEx("date") & _
							" - " & rsEx("date") + 6 & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & rsEx("client") & _
							"</font><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & GetName(rsEx("client")) & _
							"</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & _
							Z_FormatNumber(TotalHrsEx, 2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & _
							Z_FormatNumber(tmpUnitEx, 2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>$" & _
							Z_FormatNumber(RAmountEx, 2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & _
							Z_Pcode(rsEx("client")) & "</font></td></tr>" & vbCrLf
							
						strProcBXexp = strProcBXexp & """" & rsEx("date") & " - " & rsEx("date") + 6 & _
										""",""" & "*" & rsEx("client") & _
										""",""" & GetNameCSV(rsEx("client")) & """,""" & _
										Z_FormatNumber(TotalHrsEx, 2) & """,""" & Z_FormatNumber(tmpUnitEx, 2) & """,""$" & Z_FormatNumber(RAmountEx, 2) & _
										""",""" & Z_Pcode(rsEx("client")) & """" & vbCrLf
							rsEx.MoveNext
						Loop
						markme = 1
							'not yet processed
						TotalHrs = rsProcX("mon") + rsProcX("tue") + rsProcX("wed") + rsProcX("thu") + rsProcX("fri") + rsProcX("sat") + rsProcX("sun")
						tmpUnit = TotalHrs * 4
						'tmpMileAmt = GetMRate(rsProcX("date"),rsProcX("mile"))
						'tmpMilesOnly = rsProcX("mile")
						Rmon = GetRate(rsProcX("date"), rsProcX("mon"), "MON")
						Rtue = GetRate(rsProcX("date"), rsProcX("tue"), "TUE")
						Rwed = GetRate(rsProcX("date"), rsProcX("wed"), "WED")
						Rthur = GetRate(rsProcX("date"), rsProcX("thu"), "THU")
						Rfri = GetRate(rsProcX("date"), rsProcX("fri"), "FRI")
						Rsat = GetRate(rsProcX("date"), rsProcX("sat"), "SAT")
						Rsun = GetRate(rsProcX("date"), rsProcX("sun"), "SUN")
						RAmount = Rmon + Rtue + Rwed + Rthur + Rfri + Rsat + Rsun '+ tmpMileAmt
					  RF = "#000000"
						If rsProcX("MAX") = True Then RF = "Red"
						'MC = ""
						'If rsProc("milecap") = True Then MC = "**"
						strTBLEx = strTBLEx & "<tr bgcolor ='#FFFFFF'><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & rsProcX("date") & _
								" - " & rsProcX("date") + 6 & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & rsProcX("client") & _
								"</font><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & GetName(rsProcX("client")) & _
								"</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & _
								Z_FormatNumber(TotalHrs, 2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & _
								Z_FormatNumber(tmpUnit, 2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & ">$" & _
								Z_FormatNumber(RAmount, 2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & ">" & _
								Z_Pcode(rsProcX("client")) & "</font></td></tr><tr><td colspan='9'>&nbsp;</td></tr>" & vbCrLf
								
						strProcBXexp = strProcBXexp & """" & rsProcX("date") & " - " & rsProcX("date") + 6 & _
										""",""" & rsProcX("client") & _
										""",""" & GetNameCSV(rsProcX("client")) & """,""" & _
										Z_FormatNumber(TotalHrs, 2) & """,""" & Z_FormatNumber(tmpUnit, 2) & """,""$" & Z_FormatNumber(RAmount, 2) & _
										""",""" & Z_Pcode(rsProcX("client")) & """" & vbCrLf
						
						'strBIll2 = strBILL2 & """" & FixDateFormat(rsProcX("date")) & " - " & FixDateFormat(dateadd("d", 6, rsProcX("date"))) & " for IHC" & """" & vbCrLf
						'strBIll2 = strBILL2 & """" & rsProcX("client") & """,""" & "001" & """,""" & "IHC" & """,""" & "IHC" & """,""" & "3440" & """,""" & _
						'			FixDateFormat(GetSunwk1(rsProcX("date"))) & """,""" & FixDateFormat(GetSatwk2(rsProcX("date"))) & """,""" & tmpUnit & """,""" & """" & vbCrLf					
						rsProcX("ProcMed") = Date
						rsProcX.Update
					End If
					rsEx.Close
					Set rsEx = Nothing
				
					'''''''''''''''''''
				ElseIf Request("type") = 3 Then 'private pay
					'LOOK FOR PREVIOUS ITEMS - ADDITONAL HRS
					Set rsEx = Server.CreateObject("ADODB.RecordSet")
					sqlEx = "SELECT * FROM Tsheets_T WHERE client = '" & rsProcX("client") & "' AND emp_id = '" & rsProcX("emp_id") & _
						"' AND date =  '" & rsProcX("date") & "' AND EXT = 0 AND NOT ProcPriv IS NULL ORDER BY Date, timestamp"
						
					rsEx.Open sqlEx, g_strCONN, 3, 1
					If Not rsEx.EOF Then
					
						Do Until rsEx.EOF
							'REGULAR HOURS
							TotalHrsEx = rsEx("mon") + rsEx("tue") + rsEx("wed") + rsEx("thu") + rsEx("fri") + rsEx("sat") + rsEx("sun")
							RAmountEx = GetPRate(TotalHrsEx, rsEx("client"), 0)
							'HOLIDAY HOURS
							Hmon = (Hmon) + (GetHoliday(rsEx("date"), rsEx("mon"), "MON"))
							Htue = (Htue) + (GetHoliday(rsEx("date"), rsEx("tue"), "TUE"))
							Hwed = (Hwed) + (GetHoliday(rsEx("date"), rsEx("wed"), "WED"))
							Hthur = (Hthur) + (GetHoliday(rsEx("date"), rsEx("thu"), "THU"))
							Hfri = (Hfri) + (GetHoliday(rsEx("date"), rsEx("fri"), "FRI"))
							Hsat = (Hsat) + (GetHoliday(rsEx("date"), rsEx("sat"), "SAT"))
							Hsun = (Hsun) + (GetHoliday(rsEx("date"), rsEx("sun"), "SUN"))
							tmpHoldidayHrs = Hmon + Htue + Hwed + Hthur + Hfri + Hsat + Hsun
							RAmountEx2 =  GetPRate(tmpHoldidayHrs, rsEx("client"), 1)
							Hmon = 0
							Htue = 0
							Hwed = 0
							Hthur = 0
							Hfri = 0
							Hsat = 0
							Hsun = 0
							RF = "#000000"
							If rsEx("MAX") = True Then RF = "Red"		
							'MC = ""
							'If rsEx("milecap") = True Then MC = "**"		
							myRate = GetPRate2(rsEx("client"))
							RMileAmt = rsEx("mile") + rsEx("amile") 'GetMRate(rsEx("date"),rsEx("mile")) + GetMRate(rsEx("date"),rsEx("amile"))	
							TotAmt = RAmountEx2 + RAmountEx
							strTBLEx = strTBLEx & "<tr bgcolor = '#FFFCCC'><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & rsEx("date") & _
								" - " & rsEx("date") + 6 & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & rsEx("client") & _
								"</font><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & GetName(rsEx("client")) & _
								"</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & _
								Z_FormatNumber(TotalHrsEx - tmpHoldidayHrs, 2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & _
								Z_FormatNumber(tmpHoldidayHrs, 2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>$" & _
								Z_FormatNumber(myRate, 2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>$" & _
								Z_FormatNumber(RAmountEx, 2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>$" & _
								Z_FormatNumber(RAmountEx2, 2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>$" & _
								Z_FormatNumber(TotAmt, 2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & _
								Z_FormatNumber(RMileAmt, 2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & _
								Z_Pcode(rsEx("client")) & "</font></td></tr>" & vbCrLf
							
						strProcBXexp = strProcBXexp & """" & rsEx("date") & " - " & rsEx("date") + 6 & _
										""",""" & "*" & rsEx("client") & _
										""",""" & GetNameCSV(rsEx("client")) & """,""" & _
										Z_FormatNumber(TotalHrsEx - tmpHoldidayHrs, 2) & """,""" & Z_FormatNumber(tmpHoldidayHrs, 2) & """,""$" & Z_FormatNumber(myRate, 2) & """,""$" & Z_FormatNumber(RAmountEx, 2)& _
										""",""$" & Z_FormatNumber(RAmountEx2, 2) & """,""$" & Z_FormatNumber(TotAmt, 2) & """,""" & Z_FormatNumber(RMileAmt, 2) & _
										""",""" & Z_Pcode(rsEx("client")) & """" & vbCrLf '""",""" & rsEx("misc_notes") & """" & vbCrLf
							rsEx.MoveNext
						Loop
						markme = 1
						'not yet processed
						'REGULAR HOURS
						TotalHrs = rsProcX("mon") + rsProcX("tue") + rsProcX("wed") + rsProcX("thu") + rsProcX("fri") + rsProcX("sat") + rsProcX("sun")
						RAmount = GetPRate(TotalHrs, rsProcX("client"), 0)
					  'HOLIDAY HOURS
				  	Hmon = (Hmon) + (GetHoliday(rsProcX("date"), rsProcX("mon"), "MON"))
						Htue = (Htue) + (GetHoliday(rsProcX("date"), rsProcX("tue"), "TUE"))
						Hwed = (Hwed) + (GetHoliday(rsProcX("date"), rsProcX("wed"), "WED"))
						Hthur = (Hthur) + (GetHoliday(rsProcX("date"), rsProcX("thu"), "THU"))
						Hfri = (Hfri) + (GetHoliday(rsProcX("date"), rsProcX("fri"), "FRI"))
						Hsat = (Hsat) + (GetHoliday(rsProcX("date"), rsProcX("sat"), "SAT"))
						Hsun = (Hsun) + (GetHoliday(rsProcX("date"), rsProcX("sun"), "SUN"))
						tmpHoldidayHrsP = Hmon + Htue + Hwed + Hthur + Hfri + Hsat + Hsun
						RAmountP =  GetPRate(tmpHoldidayHrsP, rsProcX("client"), 1)
						Hmon = 0
						Htue = 0
						Hwed = 0
						Hthur = 0
						Hfri = 0
						Hsat = 0
						Hsun = 0
						RMileAmt = rsProcX("mile") + rsProcX("amile")
					  RF = "#000000"
						If rsProcX("MAX") = True Then RF = "Red"
						'MC = ""
						'If rsProc("milecap") = True Then MC = "**"
						TotAmt = RAmountP + RAmount
						strTBLEx = strTBLEx & "<tr bgcolor = '#FFFFFF'><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & rsProcX("date") & _
								" - " & rsProcX("date") + 6 & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & rsProcX("client") & _
								"</font><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & GetName(rsProcX("client")) & _
								"</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & _
								Z_FormatNumber(TotalHrs - tmpHoldidayHrsP, 2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & _
								Z_FormatNumber(tmpHoldidayHrsP, 2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>$" & _
								Z_FormatNumber(myRate, 2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>$" & _
								Z_FormatNumber(RAmount, 2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>$" & _
								Z_FormatNumber(RAmountP, 2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>$" & _
								Z_FormatNumber(TotAmt, 2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & _
								Z_FormatNumber(RMileAmt, 2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>&nbsp;" & _
								Z_Pcode(rsProcX("client")) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & _
								rsProcX("misc_notes") & "</font></td></tr>" & vbCrLf
								
						strProcBXexp = strProcBXexp & """" & rsProcX("date") & " - " & rsProcX("date") + 6 & _
								""",""" & rsProcX("client") & _
								""",""" & GetNameCSV(rsProcX("client")) & """,""" & _
								Z_FormatNumber(TotalHrs - tmpHoldidayHrsP, 2) & """,""" & Z_FormatNumber(tmpHoldidayHrsP, 2) & """,""$" & Z_FormatNumber(myRate, 2) & """,""$" & Z_FormatNumber(RAmount, 2)& _
								""",""$" & Z_FormatNumber(RAmountP, 2) & """,""$" & Z_FormatNumber(TotAmt, 2) & """,""" & Z_FormatNumber(RMileAmt, 2) & """,""" & _
								Z_Pcode(rsProcX("client")) & """,""" & rsProcX("misc_notes") & """" & vbCrLf
						
											
						rsProcX("ProcPriv") = Date
						rsProcX.Update
					End If
					rsEx.Close
					Set rsEx = Nothing
				ElseIf Request("type") = 4 Then	
					'LOOK FOR PREVIOUS ITEMS - ADDITONAL HRS
					Set rsEx = Server.CreateObject("ADODB.RecordSet")
					sqlEx = "SELECT * FROM Tsheets_T WHERE client = '" & rsProcX("client") & "' AND emp_id = '" & rsProcX("emp_id") & _
						"' AND date =  '" & rsProcX("date") & "' AND EXT = 0 AND NOT ProcVA IS NULL ORDER BY Date, timestamp"
					rsEx.Open sqlEx, g_strCONN, 3, 1
					If Not rsEx.EOF Then
					
						Do Until rsEx.EOF
							TotalHrsEx = rsEx("mon") + rsEx("tue") + rsEx("wed") + rsEx("thu") + rsEx("fri") + rsEx("sat") + rsEx("sun")
							tmpMileAmtEx = GetMRate(rsEx("date"),rsEx("mile"))
							'tmpMileOnlyEx = rsEx("mile")
							tmpUnitEx = TotalHrsEx * 4
							If Instr(rsEx("misc_notes"), "80,") > 0 Then
								RmonEx = GetRateHM(rsEx("date"), rsEx("mon"), "MON")
								RtueEx = GetRateHM(rsEx("date"), rsEx("tue"), "TUE")
								RwedEx = GetRateHM(rsEx("date"), rsEx("wed"), "WED")
								RthurEx = GetRateHM(rsEx("date"), rsEx("thu"), "THU")
								RfriEx = GetRateHM(rsEx("date"), rsEx("fri"), "FRI")
								RsatEx = GetRateHM(rsEx("date"), rsEx("sat"), "SAT")
								RsunEx = GetRateHM(rsEx("date"), rsEx("sun"), "SUN")
								RAmountEx = RmonEx + RtueEx + RwedEx + RthurEx + RfriEx + RsatEx + RsunEx '+ tmpMileAmtEx
								
							ElseIf Instr(rsEx("misc_notes"), "82,") > 0 Then
								RmonEx = GetRateHA(rsEx("date"), rsEx("mon"), "MON")
								RtueEx = GetRateHA(rsEx("date"), rsEx("tue"), "TUE")
								RwedEx = GetRateHA(rsEx("date"), rsEx("wed"), "WED")
								RthurEx = GetRateHA(rsEx("date"), rsEx("thu"), "THU")
								RfriEx = GetRateHA(rsEx("date"), rsEx("fri"), "FRI")
								RsatEx = GetRateHA(rsEx("date"), rsEx("sat"), "SAT")
								RsunEx = GetRateHA(rsEx("date"), rsEx("sun"), "SUN")
							End If
							RF = "#000000"
							If rsEx("MAX") = True Then RF = "Red"		
							'MC = ""
							'If rsEx("milecap") = True Then MC = "**"					
							strTBLEx = strTBLEx & "<tr bgcolor = '#FFFCCC'><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & rsEx("date") & _
							" - " & rsEx("date") + 6 & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & rsEx("client") & _
							"</font><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & GetName(rsEx("client")) & _
							"</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & _
							Z_FormatNumber(TotalHrsEx, 2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & _
							Z_FormatNumber(tmpUnitEx, 2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>$" & _
							Z_FormatNumber(RAmountEx, 2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & _
							Z_Pcode(rsEx("client")) & "</font></td></tr>" & vbCrLf
							
						strProcBXexp = strProcBXexp & """" & rsEx("date") & " - " & rsEx("date") + 6 & _
										""",""" & "*" & rsEx("client") & _
										""",""" & GetNameCSV(rsEx("client")) & """,""" & _
										Z_FormatNumber(TotalHrsEx, 2) & """,""" & Z_FormatNumber(tmpUnitEx, 2) & """,""$" & Z_FormatNumber(RAmountEx, 2) & _
										""",""" & Z_Pcode(rsEx("client")) & """" & vbCrLf
							rsEx.MoveNext
						Loop
						markme = 1
							'not yet processed
						TotalHrs = rsProcX("mon") + rsProcX("tue") + rsProcX("wed") + rsProcX("thu") + rsProcX("fri") + rsProcX("sat") + rsProcX("sun")
						tmpUnit = TotalHrs * 4

						If Instr(rsProcX("misc_notes"), "80,") > 0 Then
							Rmon = GetRateHM(rsProcX("date"), rsProcX("mon"), "MON")
							Rtue = GetRateHM(rsProcX("date"), rsProcX("tue"), "TUE")
							Rwed = GetRateHM(rsProcX("date"), rsProcX("wed"), "WED")
							Rthur = GetRateHM(rsProcX("date"), rsProcX("thu"), "THU")
							Rfri = GetRateHM(rsProcX("date"), rsProcX("fri"), "FRI")
							Rsat = GetRateHM(rsProcX("date"), rsProcX("sat"), "SAT")
							Rsun = GetRateHM(rsProcX("date"), rsProcX("sun"), "SUN")
							RAmount = Rmon + Rtue + Rwed + Rthur + Rfri + Rsat + Rsun '+ tmpMileAmt
						ElseIf Instr(rsProcX("misc_notes"), "82,") > 0 Then
							Rmon = GetRateHA(rsProcX("date"), rsProcX("mon"), "MON")
							Rtue = GetRateHA(rsProcX("date"), rsProcX("tue"), "TUE")
							Rwed = GetRateHA(rsProcX("date"), rsProcX("wed"), "WED")
							Rthur = GetRateHA(rsProcX("date"), rsProcX("thu"), "THU")
							Rfri = GetRateHA(rsProcX("date"), rsProcX("fri"), "FRI")
							Rsat = GetRateHA(rsProcX("date"), rsProcX("sat"), "SAT")
							Rsun = GetRateHA(rsProcX("date"), rsProcX("sun"), "SUN")
							RAmount = Rmon + Rtue + Rwed + Rthur + Rfri + Rsat + Rsun '+ tmpMileAmt
						End If
					  RF = "#000000"
						If rsProcX("MAX") = True Then RF = "Red"
						'MC = ""
						'If rsProc("milecap") = True Then MC = "**"
						strTBLEx = strTBLEx & "<tr bgcolor ='#FFFFFF'><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & rsProcX("date") & _
								" - " & rsProcX("date") + 6 & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & rsProcX("client") & _
								"</font><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & GetName(rsProcX("client")) & _
								"</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & _
								Z_FormatNumber(TotalHrs, 2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & _
								Z_FormatNumber(tmpUnit, 2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & ">$" & _
								Z_FormatNumber(RAmount, 2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & ">" & _
								Z_Pcode(rsProcX("client")) & "</font></td></tr><tr><td colspan='9'>&nbsp;</td></tr>" & vbCrLf
								
						strProcBXexp = strProcBXexp & """" & rsProcX("date") & " - " & rsProcX("date") + 6 & _
										""",""" & rsProcX("client") & _
										""",""" & GetNameCSV(rsProcX("client")) & """,""" & _
										Z_FormatNumber(TotalHrs, 2) & """,""" & Z_FormatNumber(tmpUnit, 2) & """,""$" & Z_FormatNumber(RAmount, 2) & _
										""",""" & Z_Pcode(rsProcX("client")) & """" & vbCrLf
						
						'strBIll2 = strBILL2 & """" & FixDateFormat(rsProcX("date")) & " - " & FixDateFormat(dateadd("d", 6, rsProcX("date"))) & " for IHC" & """" & vbCrLf
						'strBIll2 = strBILL2 & """" & rsProcX("client") & """,""" & "001" & """,""" & "IHC" & """,""" & "IHC" & """,""" & "3440" & """,""" & _
						'			FixDateFormat(GetSunwk1(rsProcX("date"))) & """,""" & FixDateFormat(GetSatwk2(rsProcX("date"))) & """,""" & tmpUnit & """,""" & """" & vbCrLf					
						rsProcX("ProcVA") = Date
						rsProcX.Update
					End If
					rsEx.Close
					Set rsEx = Nothing
				
					'''''''''''''''''''
				End If
			rsProcX.MoveNext
		Loop	
	End If
	rsProcX.Close
	Set rsProcX = Nothing
	'SET VARIABLES TO 0
	Rmon = 0
	Rtue = 0
	Rwed = 0
	Rthur = 0
	Rfri = 0
	Rsat = 0
	Rsun = 0
	Hmon = 0
	Htue = 0
	Hwed = 0
	Hthur = 0
	Hfri = 0
	Hsat = 0
	Hsun = 0
	'LOOK FOR NEW HRS NOT ADDITIONAL
	Set rsProcX2 = Server.CreateObject("ADODB.RecordSet")
	rsProcX2.Open sqlProcX, g_strCONN, 1, 3
	response.write sqlProcX & "<br>"
	If Not rsProcX2.EOF Then
		Do Until rsProcX2.EOF
			myDATE = rsProcX2("date")
			If Request("type") = 1 Then
				strEmpID = rsProcX2("emp_id")
				'rsProcX("ProcPay") = Date	
				RAmount = 0
				strWorID = ""
				RMileAmt = GetMRate(rsProcX2("date"),rsProcX2("mile")) + GetMRate(rsProcX2("date"),rsProcX2("amile"))
				RMileOnly = Z_CZero(rsProcX2("mile")) + Z_CZero(rsProcX2("amile"))
			ElseIf Request("type") = 2 Then 
				strEmpID = rsProcX2("client")
				'strWorID = rsProcX2("emp_id")
				'RAmount = 0
				'RMileAmt = 0
				Rmon = (Rmon) + (GetRate(myDATE, rsProcX2("mon"), "MON"))
				Rtue = (Rtue) + (GetRate(myDATE, rsProcX2("tue"), "TUE"))
				Rwed = (Rwed) + (GetRate(myDATE, rsProcX2("wed"), "WED"))
				Rthur = (Rthur) + (GetRate(myDATE, rsProcX2("thu"), "THU"))
				Rfri = (Rfri) + (GetRate(myDATE, rsProcX2("fri"), "FRI"))
				Rsat = (Rsat) + (GetRate(myDATE, rsProcX2("sat"), "SAT"))
				Rsun = (Rsun) + (GetRate(myDATE, rsProcX2("sun"), "SUN"))	
				'RMileAmt = GetMRate(rsProcX2("date"),rsProcX2("mile"))
				'RMileOnly = rsProcX2("mile")
				RAmount = (Rmon) + (Rtue) + (Rwed) + (Rthur) + (Rfri) + (Rsat) + (Rsun) '+ RMileAmt
				Rmon = 0
				Rtue = 0
				Rwed = 0
				Rthur = 0
				Rfri = 0
				Rsat = 0
				Rsun = 0
			ElseIf Request("type") = 3 Then
				strEmpID = rsProcX2("client")
				'strWorID = rsProcX2("emp_id")
				RAmountP = 0
				RMileAmt = 0
				RMileOnly = 0
				TotalHrs = rsProcX2("mon") + rsProcX2("tue") + rsProcX2("wed") + rsProcX2("thu") + rsProcX2("fri") + rsProcX2("sat") + rsProcX2("sun")
				RAmount = GetPRate(TotalHrs, rsProcX2("client"), 0)
				Hmon = (Hmon) + (GetHoliday(myDATE, rsProcX2("mon"), "MON"))
				Htue = (Htue) + (GetHoliday(myDATE, rsProcX2("tue"), "TUE"))
				Hwed = (Hwed) + (GetHoliday(myDATE, rsProcX2("wed"), "WED"))
				Hthur = (Hthur) + (GetHoliday(myDATE, rsProcX2("thu"), "THU"))
				Hfri = (Hfri) + (GetHoliday(myDATE, rsProcX2("fri"), "FRI"))
				Hsat = (Hsat) + (GetHoliday(myDATE, rsProcX2("sat"), "SAT"))
				Hsun = (Hsun) + (GetHoliday(myDATE, rsProcX2("sun"), "SUN"))
				tmpHoldidayHrs = Hmon + Htue + Hwed + Hthur + Hfri + Hsat + Hsun
				RAmountP =  GetPRate(tmpHoldidayHrs, rsProcX2("client"), 1)
				Hmon = 0
				Htue = 0
				Hwed = 0
				Hthur = 0
				Hfri = 0
				Hsat = 0
				Hsun = 0
				'myRate = GetPRate2(rsProc2("client"))
				RMileAmt = GetMRate(rsProcX2("date"),rsProcX2("mile")) + GetMRate(rsProcX2("date"),rsProcX2("amile"))
				RMileOnly = Z_CZero(rsProcX2("mile")) + Z_CZero(rsProcX2("amile")) 
				'response.write RMileOnly & "<br>"
			ElseIf Request("type") = 4 Then
						strEmpID = rsProcX2("client")
						If Instr(rsProcX2("misc_notes"), "80,") > 0 Then
							Rmon = (Rmon) + (GetRateHM(myDATE2, rsProcX2("mon"), "MON"))
							Rtue = (Rtue) + (GetRateHM(myDATE2, rsProcX2("tue"), "TUE"))
							Rwed = (Rwed) + (GetRateHM(myDATE2, rsProcX2("wed"), "WED"))
							Rthur = (Rthur) + (GetRateHM(myDATE2, rsProcX2("thu"), "THU"))
							Rfri = (Rfri) + (GetRateHM(myDATE2, rsProcX2("fri"), "FRI"))
							Rsat = (Rsat) + (GetRateHM(myDATE2, rsProcX2("sat"), "SAT"))
							Rsun = (Rsun) + (GetRateHM(myDATE2, rsProcX2("sun"), "SUN"))	
						ElseIf Instr(rsProcX2("misc_notes"), "82,") > 0 Then
							Rmon = (Rmon) + (GetRateHA(myDATE2, rsProcX2("mon"), "MON"))
							Rtue = (Rtue) + (GetRateHA(myDATE2, rsProcX2("tue"), "TUE"))
							Rwed = (Rwed) + (GetRateHA(myDATE2, rsProcX2("wed"), "WED"))
							Rthur = (Rthur) + (GetRateHA(myDATE2, rsProcX2("thu"), "THU"))
							Rfri = (Rfri) + (GetRateHA(myDATE2, rsProcX2("fri"), "FRI"))
							Rsat = (Rsat) + (GetRateHA(myDATE2, rsProcX2("sat"), "SAT"))
							Rsun = (Rsun) + (GetRateHA(myDATE2, rsProcX2("sun"), "SUN"))	
						End If
						RAmount = (Rmon) + (Rtue) + (Rwed) + (Rthur) + (Rfri) + (Rsat) + (Rsun) '+ RMileAmt
						Rmon = 0
						Rtue = 0
						Rwed = 0
						Rthur = 0
						Rfri = 0
						Rsat = 0
						Rsun = 0
					End If
			If Not IsNull(rsProcX2("lname")) Then
				strName = Replace(rsProcX2("lname"),",", "") & ", " & rsProcX2("fname")
			Else
				strName = rsProcX2("lname") & ", " & rsProcX2("fname")
			End If
			dblHours = rsProcX2("mon") + rsProcX2("tue") + rsProcX2("wed") + rsProcX2("thu") + rsProcX2("fri") + rsProcX2("sat") + rsProcX2("sun")
			If Request("type") = 1 Then
				'HOLIDAY HRS
				
				Hmon = (Hmon) + (GetHoliday(myDATE, rsProcX2("mon"), "MON"))
				Htue = (Htue) + (GetHoliday(myDATE, rsProcX2("tue"), "TUE"))
				Hwed = (Hwed) + (GetHoliday(myDATE, rsProcX2("wed"), "WED"))
				Hthur = (Hthur) + (GetHoliday(myDATE, rsProcX2("thu"), "THU"))
				Hfri = (Hfri) + (GetHoliday(myDATE, rsProcX2("fri"), "FRI"))
				Hsat = (Hsat) + (GetHoliday(myDATE, rsProcX2("sat"), "SAT"))
				Hsun = (Hsun) + (GetHoliday(myDATE, rsProcX2("sun"), "SUN"))
				tmpHoldidayHrs = Hmon + Htue + Hwed + Hthur + Hfri + Hsat + Hsun
				Hmon = 0
				Htue = 0
				Hwed = 0
				Hthur = 0
				Hfri = 0
				Hsat = 0
				Hsun = 0
				'PTO
				tmpPTOhrs2 = 0
				Set rsPTO = Server.CreateObject("ADODB.RecordSet")
				sqlPTO = "SELECT * FROM W_PTO_T WHERE WorkerID = '" & strEmpID & "' AND date >= '" & myDATE & "' AND date <= '" & _
					GetNextSat(myDate,wk1) & "' AND procitem IS NULL"
				'response.write sqlPTO
				rsPTO.Open sqlPTO, g_strCONN, 1, 3
				Do Until rsPTO.EOF
					tmpPTOhrs2 = tmpPTOhrs2 + rsPTO("PTO")
					rsPTO("procitem") = date
					rsPTO.update
					rsPTO.MoveNext
				Loop
				rsPTO.Close
				Set rsPTO = Nothing
			End If
			strEXT = False
			If dblHours <> 0 And rsProcX2("EXT") = True Then strEXT = True
			strMile = RMileAmt
			strCap = False
			If Request("type") = 1 Then
				If rsProcX2("milecap") = True Then strCap = True
			Else
				If rsProcX2("milecap") = True Then strCap = True
			End If
			strNotes = 	rsProcX2("misc_notes") 
			strMax = False
			If rsProcX2("MAX") = True Then strMax = True
			stractcode = rsProcX2("misc_notes")
			' find the 2-week period
			strWeekLabel = Z_Find2WkPeriod(myDate)
			' search for it in the arrays
			'lngIdx = SearchArrays(strWeekLabel, strEmpID, strWorID, tmpDates, tmpIDs, tmpWorID)
			If Request("type") <> 4 Then
						lngIdx = SearchArrays2(strWeekLabel, strEmpID, strWorID, tmpDates2, tmpIDs2, tmpWorID2)
					Else
						lngIdx = SearchArrays4(strWeekLabel, strEmpID, strWorID, stractcode, tmpDates2, tmpIDs2, tmpWorID2, tmpactcode)
					End If
			If lngIdx < 0 Then ' this is the first time i've encountered the date and id pair, so i make a new entry
				ReDim Preserve tmpDates(x)
				ReDim Preserve tmpWorID(x)
				ReDim Preserve tmpIDs(x)
				ReDim Preserve tmpHrs(x)
				ReDim Preserve tmpName(x)
				ReDim Preserve tmpEXT(x)
				ReDim Preserve tmpAmount(x)
				ReDim Preserve tmpNotes(x)
				ReDim Preserve tmpMax(x)
				Redim Preserve tmpMileCap(x)
				Redim Preserve tmpCap(x)
				Redim Preserve tmpMileOnly(x)
				Redim Preserve tmpHhrs2(x)
				ReDim Preserve tmpPTO2(x)
				ReDim Preserve tmpActcode2(x)
				
				tmpDates(x) = strWeekLabel
				tmpIDs(x) = strEmpID
				tmpWorID(x) = strWorID
				tmpHrs(x) = dblHours
				tmpName(x) = strName
				if tmpEXT(x) = False Then tmpEXT(x) = strEXT
				tmpAmount(x) = RAmount
				tmpNotes(x) = strNotes
				If tmpMax(x) = False Then tmpMax(x) = strMax
				tmpMileOnly(x) = RMileOnly
				tmpMileCap(x) = strMile
				if tmpCap(x) = False Then tmpCap(x) = strCap
				tmpHhrs2(x) = tmpHoldidayHrs
				tmpPTO2(x) = tmpPTOhrs2
				tmpActcode2(x) = stractcode
				x = x + 1
			Else
				'tmpPTO2(lngIdx) = tmpPTO2(lngIdx) + tmpPTOhrs2
				tmpHhrs2(lngIdx) = tmpHhrs2(lngIdx) + tmpHoldidayHrs
				tmpMileOnly(lngIdx) = tmpMileOnly(lngIdx) + RMileOnly
				tmpHrs(lngIdx) = tmpHrs(lngIdx) + dblHours
				tmpAmount(lngIdx) = tmpAmount(lngIdx) + RAmount
				tmpMileCap(lngIdx) = tmpMileCap(lngIdx) + strMile
				If strNotes <> "" Then tmpNotes(lngIdx) = tmpNotes(lngIdx) & "<br>" & strNotes
				If tmpMax(lngIdx) = False Then tmpMax(lngIdx) = strMax
				if tmpCap(lngIdx) = False Then tmpCap(lngIdx) = strCap
			End If
			rsProcX2.MoveNext
		Loop
	End If
	rsProcX2.Close
	Set rsProcX2 = Nothing	
	'TAG
	Set rsTagX = Server.CreateObject("ADODB.RecordSet")
	rsTagX.Open sqlProcX, g_strCONN, 1, 3
	Do Until rsTagX.EOF
		If Request("type") = 1 Then
			rsTagX("ProcPay") = Date	
		ElseIf Request("type") = 2 Then
			rsTagX("ProcMed") = Date
		ElseIf Request("type") = 3 Then
			rsTagX("ProcPriv") = Date
		ElseIf  Request("type") = 4 Then
			rsTAGX("ProcVA") = Date
		End If
		rsTagX.Update
		rsTagX.MoveNext
	Loop
	rsTagX.Close
	Set rsTagX = Nothing
	''''
	If strTBLEx <> "" Then
			strProcBX = strTBLEx
		End If
			If markerX = 1 Then
				y = 0
				Do Until y = x
					If Request("type") = 1 Then
						IDx = tmpIDs(y)'Right(tmpIDs(y), 4)
						strEXT = ""
						If tmpEXT(y) = True Then strEXT = "*" 
						
					Else
						strEXT = ""
						IDx = tmpIDs(y)
					End If 
					'RED FLAG
					Rf = ""	
					If tmpMax(y) = True Then RF = "Color='Red'"
					'max mile
					MC = ""
					if tmpCap(y) = true Then MC = "**"
					'CALCULATE BILL TEMP 
					tmpUnit = tmpHrs(y) * 4
				
					
					If Request("type") = 1 Then
						tmpRegHrs = tmpHrs(y) - tmpHhrs2(y)
						tmpHolHrs = tmpHhrs2(y)
						tmpTotHrs = tmpHrs(y)				
						strProcBX = strProcBX & "<tr><td align='center'>" & strEXT & "<font size='1' face='trebuchet ms' color='" & RF & "'>" & tmpDates(y) & _
										"</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & GetFileNum(IDx) & "</font></td>" & _
										"<td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" &tmpName(y) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & _
										Z_FormatNumber(tmpRegHrs,2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & _
										Z_FormatNumber(tmpHolHrs,2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & _
										Z_FormatNumber(tmpTotHrs,2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & _
										Z_FormatNumber(tmpPTO2(y),2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & MC & _
										Z_FormatNumber(tmpMileOnly(y),2) & "</font></td><td align='center'><font size='1' face='trebuchet ms'color='" & RF & "'>$" & _
										Z_FormatNumber(tmpMileCap(y),2) & "</font></td></tr>" & vbCrLf
										
						strProcBXexp = strProcBXexp & """" & strEXT & tmpDates(y) & _
									""",""" & GetFileNum(IDx) & _
									""",""" & GetNameWorkCSV(IDx) & """,""" & Z_FormatNumber(tmpRegHrs,2) & """,""" & Z_FormatNumber(tmpHolHrs,2) & """,""" & _
									Z_FormatNumber(tmpTotHrs,2) & """,""" & Z_FormatNumber(tmpPTO2(y),2) & """,""" & Z_FormatNumber(tmpMileOnly(y),2) & """,""" & Z_FormatNumber(tmpMileCap(y),2) & vbCrLf
							
							mipID = GetMipID(tmpIDs(y))		
							strMIPcsv = strMIPcsv & "HTS," & mipID & ",03,R" & vbCrlf & _
								 "DTSEARN," & mipID & ",Regular," & Z_FormatNumber(tmpHrs(y), 2) & ",OQA EX" & vbCrlf & vbCrlf
					ElseIf Request("type") = 2 Then
						'If tmpHrs(y) <> 0 Then
							strProcBX = strProcBX & "<tr><td align='center'>" & strEXT & "<font size='1' face='trebuchet ms' color='" & RF & "'>" & tmpDates(y) & _
											"</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & IDx & "</font></td>" & _
											"<td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" &tmpName(y) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & _
											Z_FormatNumber(tmpHrs(y),2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & _
											Z_FormatNumber(tmpUnit, 2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>$" & _
											Z_FormatNumber(tmpAmount(y), 2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & _
											Z_Pcode(IDx) & "</font></td></tr>" & vbCrLf
											
							strProcBXexp = strProcBXexp & """" & strEXT & tmpDates(y) & _
										""",""" & IDx & _
										""",""" & GetNameCSV(IDx) & """,""" & _
										tmpHrs(y) & """,""" & tmpUnit & """,""$" & tmpAmount(y) & """,""" & Z_Pcode(IDx) & """" & vbCrLf
										
						'strBIll2 = strBILL2 & """" & FixDateFormat(myDATE) & " - " & FixDateFormat(dateadd("d", 6, myDATE)) & " for IHC" & """" & vbCrLf
						'strBIll2 = strBILL2 & """" & IDx & """,""" & "001" & """,""" & "IHC" & """,""" & "IHC" & """,""" & "3440" & """,""" & _
						'		FixDateFormat(GetSunwk1(mydate)) & """,""" & FixDateFormat(GetSatwk2(mydate)) & """,""" & tmpUnit & """,""" & """" & vbCrLf	
						'End If
					ElseIf Request("type") = 3 Then
						'response.write "PASOK"
						tmpRegHrs = tmpHrs(y) - tmpHhrs2(y)
						tmpHolHrs = tmpHhrs2(y)
						tmpTotHrs = tmpHrs(y)	
						myRate = GetPRate2(IDx)
						myAmount = Z_Czero(tmpRegHrs * myRate)
						myHAmount = Z_Czero(tmpHolHrs * myRate * 1.5)
						TotAmt = Z_Czero(myHAmount + myAmount)
						strProcBX = strProcBX & "<tr><td align='center'>" & strEXT & "<font size='1' face='trebuchet ms' color='" & RF & "'>" & tmpDates(y) & _
							"</font></td><td align='center'><font size='1' face='trebuchet ms' " & RF & ">" & IDx & "</font></td>" & _
							"<td align='center'><font size='1' face='trebuchet ms' " & RF & ">" & tmpName(y) & "</font></td><td align='center'><font size='1' face='trebuchet ms' " & RF & ">" & _
							Z_FormatNumber(tmpRegHrs,2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & _
							Z_FormatNumber(tmpHolHrs, 2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & _
							Z_FormatNumber(myRate, 2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & _
							Z_FormatNumber(myAmount, 2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & _
							Z_FormatNumber(myHAmount, 2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & _
							Z_FormatNumber(TotAmt, 2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & _
							Z_FormatNumber(tmpMileOnly(y), 2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & _
							Z_Pcode(IDx) & "</font></td></tr>" & vbCrLf
											
						strProcBXexp = strProcBXexp & """" & strEXT & tmpDates(y) & _
							""",""" & IDx & _
							""",""" & GetNameCSV(IDx) & """,""" & _
							Z_FormatNumber(tmpRegHrs, 2) & """,""" & Z_FormatNumber(tmpHolHrs, 2) & """,""$" & Z_FormatNumber(myRate, 2) & """,""$" & _
							Z_FormatNumber(myAmount, 2) & """,""$" & Z_FormatNumber(myHAmount, 2) & """,""$" & _
							Z_FormatNumber(TotAmt, 2) & """,""" & Z_FormatNumber(tmpMileOnly(y), 2) & """,""" & Z_Pcode(IDx) & """" & vbCrLf '""",""" & tmpNotes(y) & """" & vbCrLf
					ElseIf Request("type") = 4 Then 'VA
							strProcBX = strProcBX & "<tr><td align='center'>" & strEXT & "<font size='1' face='trebuchet ms' color='" & RF & "'>" & tmpDates(y) & _
											"</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & IDx & "</font></td>" & _
											"<td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" &tmpName(y) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & _
											Z_FormatNumber(tmpHrs(y),2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & _
											Z_FormatNumber(tmpUnit, 2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>$" & _
											Z_FormatNumber(tmpAmount(y), 2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & RF & "'>" & _
											Z_Pcode(IDx) & "</font></td></tr>" & vbCrLf
											
							strProcBXexp = strProcBXexp & """" & strEXT & tmpDates(y) & _
										""",""" & IDx & _
										""",""" & GetNameCSV(IDx) & """,""" & _
										tmpHrs(y) & """,""" & tmpUnit & """,""$" & tmpAmount(y) & """,""" & Z_Pcode(IDx) & """" & vbCrLf
						End If
										
					
					
					y = y + 1 
				Loop
			End If
		'LOOK FOR PTO W/O HOURS'''''''''''''''''''''''''''''
		Set rsPTOhrs = Server.CreateObject("ADODB.RecordSet")
		sqlPTOhrs = "SELECT * FROM W_PTO_T, worker_T WHERE Social_Security_Number = WorkerID AND date < '" & sunDATE & "' " & _
			"AND procItem IS NULL ORDER BY lname, fname, date"
			'response.write sqlPTOhrs
		rsPTOhrs.Open sqlPTOhrs, g_strCONN,1 , 3
		If Not rsPTOhrs.EOF Then
			If strMSG = "" Then strMSG = "<tr bgcolor='#040C8B'><td colspan='9'><font size='1' face='trebuchet ms' color='white'><b>Processed items before the set payroll period</b></font></td></tr>"
			If strMSGexp = "" Then strMSGexp = "Processed items before the set payroll period"
			strProcBX = strProcBX & "<tr bgcolor='#040C8B'><td colspan='9'><font size='1' face='trebuchet ms' color='white'><b>PTO's w/o hours</b></font></td></tr>"
			strProcBXexp = strProcBXexp & "PTO's w/o hours" & vbCrLf
			strProcBX = strProcBX & "<tr bgcolor='#040C8B'><td align='center' width='100px'><font size='1' face='trebuchet ms' color='white' color='white'>Timesheet Week</font>" & _
					"</td><td align='center' width='75px'><font size='1' face='trebuchet ms' color='white' color='white'>FileNumber</font></td><td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>" & _
					"Name</font></td><td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>Reg. Hrs.</font></td>" & _
					"<td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>Holiday Hrs.</font></td>" & _
					"<td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>Total Hrs.</font></td>" & _
					"<td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>PTO</font></td>" & _
					"<td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>Mileage</font></td>" & _
					"<td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>Mileage Amt</font></td></tr>"
			strProcBXexp = strProcBXexp & "Timesheet Week, FileNumber, Last Name, First Name, Reg. Hrs., Holiday Hrs., Total Hrs., PTO, Mileage, Mileage Amt" & vbCrLf
			Do Until rsPTOhrs.EOF
				strProcBX = strProcBX & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & rsPTOhrs("date") & " - " & DateAdd("d", rsPTOhrs("date"), 6) & _
					"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & GetFileNum(rsPTOhrs("WorkerID")) & "</font></td>" & _
					"<td align='center'><font size='1' face='trebuchet ms'>" & rsPTOhrs("lname") & ", " & rsPTOhrs("fname") & "</font></td>" & _
					"<td align='center'><font size='1' face='trebuchet ms'>0.00</font></td><td align='center'><font size='1' face='trebuchet ms'>0.00</font></td>" & _
					"<td align='center'><font size='1' face='trebuchet ms'>0.00</font></td><td align='center'><font size='1' face='trebuchet ms'>" & _
					Z_FormatNumber(rsPTOhrs("PTO"),2) & "</font></td><td align='center'><font size='1' face='trebuchet ms'>0.00</font></td>" & _
					"<td align='center'><font size='1' face='trebuchet ms'>$0.00</font></td></tr>" & vbCrLf
				strProcBXexp = strProcBXexp & """" & rsPTOhrs("date") & " - " & DateAdd("d", rsPTOhrs("date"), 6) & """,""" & GetFileNum(rsPTOhrs("WorkerID")) & _
					""",""" & rsPTOhrs("lname") & """,""" & rsPTOhrs("fname") & """,""" & "0.00" & "," & "0.00" & "," & "0.00" & "," & _
					Z_FormatNumber(rsPTOhrs("PTO"),2) & """," & "0.00" & "," & "0.00" & vbCrLf
				rsPTOhrs("procitem") = date
				rsPTOhrs.Update
				rsPTOhrs.MoveNext
			Loop
		End If
		rsPTOhrs.Close
		Set rsPTO = Nothing
		
		''''OVER 40 HOURS
		If Request("type") = 1 Then
			'1st week
			strMSG2 = "<tr bgcolor='#040C8B'><td colspan='9'><font size='1' face='trebuchet ms' color='white'><b>PCSP Workers over 40 hours from " & mySunDate & " - " & DateAdd("d", 6, mySunDate) & "</b></font></td></tr>"
			strMSG2exp = "PCSP Workers over 40 hours from " & mySunDate & " - " & DateAdd("d", 6, mySunDate)
			strProcHX2 = "<tr bgcolor='#040C8B'><td align='center'><font size='1' face='trebuchet ms' color='white'><b>PCSP Worker Name</b>" & _
						"</font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Hours</b></font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>RIHCC</b></font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>RIHCC2</b></font></td></tr>" 
			strProcHX2exp = "Worker Last Name, Worker First Name, Hours, RIHCC, RIHCC2"
			
			Set rsTBL = Server.CreateObject("ADODB.RecordSet")
					sqlTBL  = "SELECT * FROM Tsheets_t, worker_T, proj_man_T, consumer_T  WHERE consumer_T.Medicaid_Number = client AND " & _
						"PMID = Proj_Man_T.ID AND emp_id = worker_T.social_security_number AND date >= '" & mySunDate & "' AND date <= '" & DateAdd("d", 6, mySunDate) & "'" & _
						" ORDER BY proj_man_T.Lname, proj_man_T.Fname, worker_T.Lname, worker_T.Fname"
					'response.write sqltBl
					rsTBL.Open sqlTBL, g_strCONN, 1, 3	
					Erase tmpDates
					Erase tmpWorID
					Erase tmpEmpID
					Erase tmpHrs
					Erase tmpPMID
					Erase tmpPMID2
					x = 0
					Do Until rsTBL.EOF
						strEmpID = rsTBL("client")
						strWorID = rsTBL("emp_id")
						dblHours = rsTBL("mon") + rsTBL("tue") + rsTBL("wed") + rsTBL("thu") + rsTBL("fri") + rsTBL("sat") + rsTBL("sun")
						strDate = rsTBL("date")
						strPMID = rsTBL("PM1")
						strPMID2 = rsTBL("PM2")
						strWeekLabel = GetSun(Request("frmd8"))
						lngIdx = SearchArrays3(strWeekLabel,  strEmpID, strWorID, tmpDates, tmpEmpID, tmpWorID)
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
							strProcBX2 = strProcBX2 & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & GetNameWork(tmpWorID(y)) & _
								"</font></td>" & _
								"<td align='center'><font size='1' face='trebuchet ms'>" & tmpHrs(y) & "</font></td>" & _
								"<td align='center'><font size='1' face='trebuchet ms'>" & GetName3(tmpPMID(y)) & "</font></td><td align='center'><font size='1' face='trebuchet ms'>" & GetName3(tmpPMID2(y)) & "</font></td></tr>"
							strProcBX2exp = strProcBX2exp & GetNameWork(tmpWorID(y)) & "," & tmpHrs(y) & ",""" & GetName3(tmpPMID(y)) & """,""" & GetName3(tmpPMID2(y)) & """" & vbCrLf
						End If
						y = y + 1
					Loop 
				
				'2nd week
				sunDATE = DateAdd("d", 7, mySunDate)
				satDATE = DateAdd("d", 6, sunDATE)
				strMSG3 = "<tr bgcolor='#040C8B'><td colspan='9'><font size='1' face='trebuchet ms' color='white'><b>PCSP Workers over 40 hours from " & sunDATE & " - " & satDATE & "</b></font></td></tr>"
				strMSG3exp = "PCSP Workers over 40 hours from " & sunDATE & " - " & satDATE
			strProcHX3 = "<tr bgcolor='#040C8B'><td align='center'><font size='1' face='trebuchet ms' color='white'><b>PCSP Worker Name</b>" & _
						"</font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Hours</b></font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>RIHCC</b></font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>RIHCC2</b></font></td></tr>" 
			strProcHX3exp = "Worker Last Name, Worker First Name, Hours, RIHCC, RIHCC2"
			Set rsTBL = Server.CreateObject("ADODB.RecordSet")
					sqlTBL  = "SELECT * FROM Tsheets_t, worker_T, proj_man_T, consumer_T  WHERE consumer_T.Medicaid_Number = client AND " & _
						"PMID = Proj_Man_T.ID AND emp_id = worker_T.social_security_number AND date >= '" & sunDATE & "' AND date <= '" & satDATE & "'" & _
						" ORDER BY proj_man_T.Lname, proj_man_T.Fname, worker_T.Lname, worker_T.Fname"
					rsTBL.Open sqlTBL, g_strCONN, 1, 3	
					Erase tmpDates
					Erase tmpWorID
					Erase tmpEmpID
					Erase tmpHrs
					Erase tmpPMID
					Erase tmpPMID2
					x = 0
					Do Until rsTBL.EOF
						strEmpID = rsTBL("client")
						strWorID = rsTBL("emp_id")
						dblHours = rsTBL("mon") + rsTBL("tue") + rsTBL("wed") + rsTBL("thu") + rsTBL("fri") + rsTBL("sat") + rsTBL("sun")
						strDate = rsTBL("date")
						strPMID = rsTBL("PM1")
						strPMID2 = rsTBL("PM2")
						strWeekLabel = GetSun(Request("frmd8"))
						lngIdx = SearchArrays3(strWeekLabel,  strEmpID, strWorID, tmpDates, tmpEmpID, tmpWorID)
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
							strProcBX3 = strProcBX3 & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & GetNameWork(tmpWorID(y)) & _
								"</font></td>" & _
								"<td align='center'><font size='1' face='trebuchet ms'>" & tmpHrs(y) & "</font></td>" & _
								"<td align='center'><font size='1' face='trebuchet ms'>" & GetName3(tmpPMID(y)) & "</font></td><td align='center'><font size='1' face='trebuchet ms'>" & GetName3(tmpPMID2(y)) & "</font></td></tr>"
							strProcBX3exp = strProcBX3exp & GetNameWork(tmpWorID(y)) & "," & tmpHrs(y) & ",""" & GetName3(tmpPMID(y)) & """,""" & GetName3(tmpPMID2(y)) & """" & vbCrLf
						End If
						y = y + 1
					Loop 
				End If
		'''''''''''''''''
		'FOR PRINT PREVIEW
		Session("PrintPrevPRoc") = strProcH & "|" & strTBL & "|" & strMSG & "|" & strProcHX & "|" & strProcBX  & "|" & Session("MSG") & _
			"|" & strMSG2 & "|" & strProcHX2 & "|" & strProcBX2 & "|" & strMSG3 & "|" & strProcHX3 & "|" & strProcBX3
		'FOR CSV
		Set fso = CreateObject("Scripting.FileSystemObject")
		a = 1
		tmpdate = replace(date, "/", "")
		ProcCSV = "C:\Work\lss-dbvortex\export\ProcItem" & tmpdate & ".csv"
		BillCSV = "C:\Work\lss-dbvortex\export\BillItem" & tmpdate & ".csv"
		if fso.FileExists(ProcCSV) THen
			Do
				if fso.FileExists("C:\Work\lss-dbvortex\export\ProcItem" & tmpdate & "-" & a & ".csv") Then
					a = a + 1
				Else
					ProcCSV = "C:\Work\lss-dbvortex\export\ProcItem" & tmpdate & "-" & a & ".csv"
					Exit Do
				End If
			Loop		 
		End If
		if fso.FileExists(BillCSV) THen
			Do
				if fso.FileExists("C:\Work\lss-dbvortex\export\BillItem" & tmpdate & "-" & a & ".csv") Then
					a = a + 1
				Else
					BillCSV = "C:\Work\lss-dbvortex\export\BillItem" & tmpdate & "-" & a & ".csv"
					Exit Do
				End If
			Loop		 
		End If
		Set Prt = fso.CreateTextFile(ProcCSV)
		Prt.WriteLine strProcHexp
		Prt.WriteLine strTBLexp
		
		Prt.WriteLine vbCrLf
		Prt.WriteLine strMSGexp
		Prt.WriteLine strProcHXexp
		Prt.WriteLine strProcBXexp
		
		Prt.WriteLine vbCrLf
		Prt.WriteLine strMSG2exp
		Prt.WriteLine strProcHX2exp
		Prt.WriteLine strProcBX2exp
		
		Prt.WriteLine vbCrLf
		Prt.WriteLine strMSG3exp
		Prt.WriteLine strProcHX3exp
		Prt.WriteLine strProcBX3exp
		If Request("type") = 1 Then Prt.WriteLine "* has extended hours"
			
		copypath = copyfile & Z_GetFilename(ProcCSV)
		Set Prtx = fso.CreateTextFile(copypath, True)
		Prtx.WriteLine strProcHexp
		Prtx.WriteLine strTBLexp
		
		Prtx.WriteLine vbCrLf
		Prtx.WriteLine strMSGexp
		Prtx.WriteLine strProcHXexp
		Prtx.WriteLine strProcBXexp
		
		Prtx.WriteLine vbCrLf
		Prtx.WriteLine strMSG2exp
		Prtx.WriteLine strProcHX2exp
		Prtx.WriteLine strProcBX2exp
		
		Prtx.WriteLine vbCrLf
		Prtx.WriteLine strMSG3exp
		Prtx.WriteLine strProcHX3exp
		Prtx.WriteLine strProcBX3exp
		If Request("type") = 1 Then Prtx.WriteLine "* has extended hours"
		Session("ProcCSV") = Z_DoEncrypt(Z_GetFilename(ProcCSV))	
		response.write Z_GetFilename(ProcCSV)
		'Session("ProcCSV") = ProcCSV
		
		'NEW BILLING
		Set Prt2 = fso.CreateTextFile(BillCSV)
		'if Left(strBILL,Len(strBILL)-1) = vbCrLf Then
		'		intLength = Len(strBILL)
		'		strEnd = Right(strBILL, 2)
				
		'		strBILL = Left(strBILL, intLength - 1)
		'End If
		Prt2.Write strBILL
		'Prt2.WriteLine strBILL2
		copypath = copyfile & Z_GetFilename(BillCSV)
		Set Prtx2 = fso.CreateTextFile(copypath, True)
		Prtx2.WriteLine strBILL
		Session("BillCSV") = Z_DoEncrypt(Z_GetFilename(BillCSV))
		'Session("BillCSV") = BillCSV
		response.write Z_GetFilename(BillCSV)
		Set fso = CreateObject("Scripting.FileSystemObject")
		Set ALog = fso.OpenTextFile(AdminLog, 8, True)
		Alog.WriteLine Now & ":: Process Item ran by: UID: " & Session("UserID") & vbCrLf
		Set Alog = Nothing
		'Set fso = Nothing 
		
		'FOR MIP
		'If strMIPcsv <> "" Then
		'	Set fso = CreateObject("Scripting.FileSystemObject")
		'	Set Mip = fso.CreateTextFile(mipCSV)
		'	Mip.WriteLine strMIPcsv
		'End If
		
		Set fso = Nothing
End If
				
				
					
%>
<html>
	<head>
		<title>Timesheet - Process Items</title>
		<link href="styles.css" type="text/css" rel="stylesheet" media="print">
		<script language='JavaScript'>
			function ExCSV()
			{
				document.frmProc.action = "Export.asp?sql=2";
				document.frmProc.submit();
			}
			function PrintPrev()
			{
				document.frmProc.action = "Print.asp";
				document.frmProc.submit();
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
	<body bgcolor='white'  LEFTMARGIN='0' TOPMARGIN='0'>
		
				<!-- #include file='_boxup.asp' -->
				<!-- #include file='_NavHeader.asp' -->
	<br>
	<form name='frmProc' method='post' action='Process.asp'>
		<center>
		<table cellSpacing='0' cellPadding='0' align='center' border='0'>
			<tr>
				<td align='center'>
					<font face='trebuchet MS' size='1'>Last Payroll Day:&nbsp;</font>
					<input maxlength='10' name='Payd8' style='font-size: 10px; height: 20px; width: 80px;'>
					<font face='trebuchet MS' size='1'>(mm/dd/yyyy)&nbsp;</font>
				</td>
				<td align='left'>
					<!--<input type='radio' checked name='type' value='1'>
					<font face='trebuchet MS' size='1'>-Payroll&nbsp;</font>
					<br>-->
					<input type='radio' name='type' value='2'>
					<font face='trebuchet MS' size='1'>-Medicaid&nbsp;</font>
					<br>
					<input type='radio' name='type' value='3'>
					<font face='trebuchet MS' size='1'>-Private Pay&nbsp;</font>
					<br>
					<input type='radio' name='type' value='4'>
					<font face='trebuchet MS' size='1'>-VA&nbsp;</font>
				</td>
				<td>
					<input type='button' class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" value='Process' onclick='document.frmProc.submit();'>
				</td>
				
			</tr>
			<tr>
				<td align='center' colspan='4'><font color='red' face='trebuchet MS' size='2'>&nbsp;<%=Session("MSG")%>&nbsp;</font></td>
			</tr>
			<tr><td>&nbsp;</td></tr>
			<tr><td align='center' colspan='4'>
							<% If strTBL <> "" Or strProcBX <> "" Then %>
				<input  style='width: 140px;' align='center' type='button' class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" value='Print Preview' onclick='PrintPrev()'>&nbsp;
				<input  style='width: 140px;' align='center' type='button' class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" value='Export table to CSV' onclick='JavaScript: ExCSV();'>
				<br>
			<% End If %>

				<table border='1'>
					<%=strProcH%>
					<%=strTBL%>
				</table>
				<br>
				<table border='1'>
					<%=strMSG%>
					<%=strProcHX%>
					<%=strProcBX%>	
				</table>
				<br>
				<table border='1'>
					<%=strMSG2%>
					<%=strProcHX2%>
					<%=strProcBX2%>	
				</table>
				<br>
				<table border='1'>
					<%=strMSG3%>
					<%=strProcHX3%>
					<%=strProcBX3%>	
				</table>
				<br>
				<center>
					
			</td></tr>
		</table>
	</form>
	<!-- #include file='_boxdown.asp' -->
	</body>
</html>
<% Session("MSG") = "" %>	