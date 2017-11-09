<%
	DIM tmp(), xtmp(), arrWor(), arrPer()
	Dim tmpDates(), tmpWorID(), tmpIDs(),  tmpHrs(), tmpPMID(), tmpEmpID(), tmpPMID2()
	dim tmpemp(), tmpcon(), tmpmile(), tmpmilecap(), tmpHrsExt(), tmpID(), tmpCID()
	dim myDay(6), tbl(6), myUnits(6), myAC(6), mysun(), rsHrs(6), sundatex(2), myacode()
	Function Z_IncludeDateHrs(ammenddte, mydte, dayname, hrs)
		Z_IncludeDateHrs = 0
		If hrs = 0 Then Exit Function
		If Z_FixNull(ammenddte) = "" Then Exit Function
		tsdate = Z_GetDate(mydte, dayname)
		If cdate(ammenddte) < cdate(tsdate) Then	Z_IncludeDateHrs = hrs
	End Function
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
	function ACValue(xxx)
			If xxx = "" Then exit function
			'response.write "Xxx: " & xxx & "<br>"
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
Function Z_MatchSkills(con, wor)
	Z_MatchSkills = 0
	Set rsCon = Server.CreateObject("ADODB.RecordSet")
	rsCon.Open "SELECT top 1 * FROM c_health_T WHERE medicaid_number = '" & con & "' ORDER BY [datestamp] desc", g_strCONN, 3, 1
	sklctr = 0
	sklmtch = 0
	If Not rsCon.EOF Then
		Set rsWor = Server.CreateObject("ADODB.RecordSet")
		rsWor.Open "SELECT * FROM w_skills_T WHERE wid = '" & wor & "' ", g_strCONN, 3, 1
		If Not rsWor.EOF Then
			If rsCon("housekeep") Then 
				sklctr = sklctr + 1
				If rsWor("housekeep") Then sklmtch = sklmtch + 1
			End If
			If rsCon("laundry") Then 
				sklctr = sklctr + 1
				If rsWor("laundry") Then sklmtch = sklmtch + 1
			End If
			If rsCon("meal") Then 
				sklctr = sklctr + 1
				If rsWor("meal") Then sklmtch = sklmtch + 1	
			End If
			If rsCon("grocery") Then 
				sklctr = sklctr + 1
				If rsWor("grocery") Then sklmtch = sklmtch + 1	
			End If
			If rsCon("dress") Then 
				sklctr = sklctr + 1
				If rsWor("dress") Then sklmtch = sklmtch + 1	
			End If
			If rsCon("eat") Then 
				sklctr = sklctr + 1
				If rsWor("eat") Then sklmtch = sklmtch + 1	
			End If
			If rsCon("asstwalk") Then 
				sklctr = sklctr + 1
				If rsWor("asstwalk") Then sklmtch = sklmtch + 1	
			End If
			If rsCon("asstwheel") Then 
				sklctr = sklctr + 1
				If rsWor("asstwheel") Then sklmtch = sklmtch + 1	
			End If
			If rsCon("asstmotor") Then 
				sklctr = sklctr + 1
				If rsWor("asstmotor") Then sklmtch = sklmtch + 1	
			End If
			If rsCon("commeal") Then 
				sklctr = sklctr + 1
				If rsWor("commeal") Then sklmtch = sklmtch + 1	
			End If
			If rsCon("medical") Then 
				sklctr = sklctr + 1
				If rsWor("medical") Then sklmtch = sklmtch + 1	
			End If
			If rsCon("shower") Then 
				sklctr = sklctr + 1
				If rsWor("shower") Then sklmtch = sklmtch + 1	
			End If
			If rsCon("tub") Then 
				sklctr = sklctr + 1
				If rsWor("tub") Then sklmtch = sklmtch + 1	
			End If
			If rsCon("oral") Then 
				sklctr = sklctr + 1
				If rsWor("oral") Then sklmtch = sklmtch + 1	
			End If
			If rsCon("commode") Then 
				sklctr = sklctr + 1
				If rsWor("commode") Then sklmtch = sklmtch + 1	
			End If
			If rsCon("sit") Then 
				sklctr = sklctr + 1
				If rsWor("sit") Then sklmtch = sklmtch + 1	
			End If
			If rsCon("medication") Then 
				sklctr = sklctr + 1
				If rsWor("medication") Then sklmtch = sklmtch + 1	
			End If
			If rsCon("undress") Then 
				sklctr = sklctr + 1
				If rsWor("undress") Then sklmtch = sklmtch + 1	
			End If
			If rsCon("shampoosink") Then 
				sklctr = sklctr + 1
				If rsWor("shampoosink") Then sklmtch = sklmtch + 1	
			End If
			If rsCon("oralcare") Then 
				sklctr = sklctr + 1
				If rsWor("oralcare") Then sklmtch = sklmtch + 1	
			End If
			If rsCon("massage") Then 
				sklctr = sklctr + 1
				If rsWor("massage") Then sklmtch = sklmtch + 1	
			End If
			If rsCon("shampoobed") Then 
				sklctr = sklctr + 1
				If rsWor("shampoobed") Then sklmtch = sklmtch + 1	
			End If
			If rsCon("shave") Then 
				sklctr = sklctr + 1
				If rsWor("shave") Then sklmtch = sklmtch + 1	
			End If
			If rsCon("bedbath") Then 
				sklctr = sklctr + 1
				If rsWor("bedbath") Then sklmtch = sklmtch + 1	
			End If
			If rsCon("bedpan") Then 
				sklctr = sklctr + 1
				If rsWor("bedpan") Then sklmtch = sklmtch + 1	
			End If
			If rsCon("ptexer") Then 
				sklctr = sklctr + 1
				If rsWor("ptexer") Then sklmtch = sklmtch + 1	
			End If
			If rsCon("hoyer") Then 
				sklctr = sklctr + 1
				If rsWor("hoyer") Then sklmtch = sklmtch + 1	
			End If
			If rsCon("eye") Then 
				sklctr = sklctr + 1
				If rsWor("eye") Then sklmtch = sklmtch + 1	
			End If
			If rsCon("transferbelt") Then 
				sklctr = sklctr + 1
				If rsWor("transferbelt") Then sklmtch = sklmtch + 1	
			End If
			If rsCon("alz") Then 
				sklctr = sklctr + 1
				If rsWor("alz") Then sklmtch = sklmtch + 1	
			End If
			If rsCon("Incontinence") Then 
				sklctr = sklctr + 1
				If rsWor("Incontinence") Then sklmtch = sklmtch + 1	
			End If
			If rsCon("Hospice") Then 
				sklctr = sklctr + 1
				If rsWor("Hospice") Then sklmtch = sklmtch + 1	
			End If
			If rsCon("lotion") Then 
				sklctr = sklctr + 1
				If rsWor("lotion") Then sklmtch = sklmtch + 1	
			End If
		End If
		rsWor.Close
		Set rsWor = Nothing
	End If
	rsCon.Close
	Set rsCon = Nothing
	If sklctr > 0 Then Z_MatchSkills = sklmtch / sklctr
End Function
Function ChkConList(con, conlist)
	ChkConList = True
	If InStr(conlist, "|" & con & "|") > 0 Then ChkConList = False 
End Function
Function GetConList(conList)
	ListCon = Split(conList, "|")
	lst = 1
	Do Until lst = Ubound(ListCon)
		GetConList = GetConList & GetName2(ListCon(lst)) & "<br>"
		lst = lst + 1
	Loop
End Function
Function GetBadge(strSSN)
	Set rsWID = Server.CreateObject("ADODB.RecordSet")
	sqlWID = "SELECT uBadge FROM Worker_T WHERE [Social_Security_Number] = '" & strSSN & "'"
	rsWID.Open sqlWID, g_strCONN, 3, 1
	If Not rsWID.EOF Then
		GetBadge = rsWID("uBadge")
	End If
	rsWID.Close
	Set rsWID = Nothing
End Function
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
	sqlRateX = "SELECT * FROM MileRate_T WHERE miledate <= '" & tmpCDate & "' ORDER BY miledate DESC"
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
	sqlRateX = "SELECT * FROM Rate_T WHERE rDate <= '" & tmpComDate & "' ORDER BY rDate DESC"
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
	Function GetCM(xxx)
		GetCM = "N/A"
		Set rsName = Server.CreateObject("ADODB.RecordSet")
		sqlName = "SELECT * FROM Proj_Man_T WHERE [ID] = " & xxx
		rsName.Open sqlName, g_strCONN, 3, 1
		If Not rsName.EOF Then
			GetCM = rsName("lname") & ", " & rsName("fname")
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
			Set rsNamed = Server.CreateObject("ADODB.RecordSet")
			sqlName = "SELECT [lname], [fname] FROM Proj_Man_T WHERE ID = " & zzz 
			rsNamed.Open sqlName, g_strCONN, 3, 1
			If Not rsNamed.EOF Then
				GetName3 = rsNamed("Lname") & ", " & rsNamed("fname")
			End If
			rsNamed.Close
			Set rsNamed = Nothing
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
	Function ValidDate2(dt1, dt2, xdate, xday)
        ValidDate2 = False

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
Function SearchArrays5(strwk, tmpDates, strCID, tmpCID)
		DIM	lngMax, lngI
	
	' START: validity check
	SearchArrays5 = -1
	On Error Resume Next	
	lngMax = UBound(tmpDates)
	If Err.Number <> 0 Then Exit Function
	
	For lngI = 0 to lngMax
		If tmpDates(lngI) = strWk And tmpCID(lngI) = strCID Then Exit For
	Next
	If lngI > lngMax Then Exit Function
	SearchArrays5 = lngI
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
Function GetCMName(xxx)
	GetCMName = "N/A"
	If xxx = "" Then Exit Function
	Set rsCode = Server.CreateObject("ADODB.RecordSet")
	sqlCode = "SELECT lname, fname From Case_Manager_t WHERE [index] = " & xxx
	rsCode.Open sqlCode, g_strCONN, 3, 1
	If Not rsCode.EOF Then GetCMName = rsCode("lname") & ", " & rsCode("fname")
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
%>