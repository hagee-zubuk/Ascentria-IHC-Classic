<%@Language=VBScript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_security.asp" -->
<%
Function IsFloat(xxx)
	IsFloat = False
	Set rsFloat = Server.CreateObject("ADODB.RecordSet")
	sqlFloat = "SELECT flt FROM Worker_T WHERE Social_Security_Number = '" & xxx & "'"
	rsFloat.Open sqlFloat, g_strCONN, 3, 1
	If Not rsFloat.EOF Then
		If rsFloat("flt") Then IsFloat = True
	End If
	rsFloat.Close
	Set rsFloat = Nothing
End Function
Function RoundTime(strTime)
	mtime = strTime 'FormatDateTime(strTime, 4)
	strHrs = DatePart("h", mtime)
	strMins = Cstr(DatePart("n", mtime))
	If strMins >= 0 And strMins < 8 Then
		strMins = "00"
	ElseIf strMins >= 8 And strMins < 22 Then
		strMins = "15"
	ElseIf strMins >= 22 And strMins < 37 Then
		strMins = "30"
	ElseIf strMins >= 37 And strMins < 52 Then
		strMins = "45"
	ElseIf strMins >= 52 And strMins < 60 Then
		strHrs = strHrs + 1
		strMins = "00"
	Else
		strMins = "00"
	End If
	'check if 24:00
	tmpRoundTime = strHrs & ":" & strMins
	If tmpRoundTime = "24:00" Then tmpRoundTime = "23:59"
	RoundTime = tmpRoundTime
End Function
Function GetHrs(strTimeIn, strTimeOut)
	strMin = DateDiff("n", strTimeIn, strTimeOut)
	If strMin > 0 Then 
		tmpBillHrs = strMin / 60
		tmpBillMHrs = Int(tmpBillHrs)
		tmpLen = Len(tmpBillHrs)
		tmpPosDec = Instr(tmpBillHrs, ".")
		tmpBillMMin = Z_FormatNumber("0" & Right(tmpBillHrs, tmpLen - (tmpPosDec - 1)), 2)
		'response.write tmpBillMMin & "<br>"
		If Cdbl(tmpBillMMin) >= 0.00 And  Cdbl(tmpBillMMin) < 0.125 Then
			GetHrs = tmpBillMHrs
		ElseIf  Cdbl(tmpBillMMin) >= 0.125 And  Cdbl(tmpBillMMin) < 0.375 Then
			GetHrs = tmpBillMHrs + 0.25
		ElseIf  Cdbl(tmpBillMMin) >= 0.375 And  Cdbl(tmpBillMMin) < 0.625 Then
			GetHrs = tmpBillMHrs + 0.5
		ElseIf  Cdbl(tmpBillMMin) >= 0.625 And  Cdbl(tmpBillMMin) < 0.875 Then
			GetHrs = tmpBillMHrs + 0.75
		ElseIf  Cdbl(tmpBillMMin) >= 0.875 And  Cdbl(tmpBillMMin) < 1 Then
			GetHrs = tmpBillMHrs + 1
		Else
				GetHrs = tmpBillMHrs
		End If
	Else
		'strMin = 1440 - Mid(strMin, 2)
		If Cdbl(strMin) >= 0.00 And  Cdbl(strMin) < 0.125 Then
			GetHrs = 0
		ElseIf  Cdbl(strMin) >= 0.125 And  Cdbl(strMin) < 0.375 Then
			GetHrs = 0.25
		ElseIf  Cdbl(strMin) >= 0.375 And  Cdbl(strMin) < 0.625 Then
			GetHrs = 0.5
		ElseIf  Cdbl(strMin) >= 0.625 And  Cdbl(strMin) < 0.875 Then
			GetHrs = 0.75
		ElseIf  Cdbl(strMin) >= 0.875 And  Cdbl(strMin) < 1 Then
			GetHrs = 1
		End If
	End If
End Function
Function GetTime(strTimeIn, strTimeOut)
	tmpBillMin = DateDiff("n", strTimeIn, strTimeOut)
	If tmpBillMin < 0 Then tmpBillMin = 1440 - Mid(tmpBillMin, 2)
	tmpBillHrs = tmpBillMin / 60
	tmpBillMHrs = Int(tmpBillHrs)
	tmpLen = Len(tmpBillHrs)
	tmpPosDec = Instr(tmpBillHrs, ".")
	If Instr(tmpBillHrs, ".") > 1 Then
		myTime = tmpBillMHrs + Z_FormatNumber("0" & Right(tmpBillHrs, tmpLen - (tmpPosDec - 1)), 2)
	Else
		myTime = tmpBillMHrs
	End If
	GetTime = myTime
End Function
Function GetSundayTS(strDate)
	If WeekDay(strDate) = 1 Then 
		GetSundayTS = strDate
	Else
		tmpDate = strDate
		Do Until Weekday(tmpDate) = 1
			tmpDate = DateAdd("d", -1, tmpDate)
		Loop
		GetSundayTS = tmpDate
	End If
End Function
Function MakeNewFileName()
	strNow = Now
	strNow = Replace(strNow, "/", "")
	strNow = Replace(strNow, ":", "")
	strNow = Replace(strNow, " ", "")
	MakeNewFileName = strNow
End Function
Function GetWIDBadge(strBadge)
	Set rsWID = Server.CreateObject("ADODB.RecordSet")
	sqlWID = "SELECT [Social_Security_Number] FROM Worker_T WHERE Badge = '" & strBadge & "'"
	rsWID.Open sqlWID, g_strCONN, 3, 1
	If Not rsWID.EOF Then
		GetWIDBadge = rsWID("Social_Security_Number")
	End If
	rsWID.Close
	Set rsWID = Nothing
End Function
Function GetCIDCustID(strCustID)
	Set rsWIDs = Server.CreateObject("ADODB.RecordSet")
	sqlWID = "SELECT [Medicaid_Number] FROM Consumer_T WHERE CliID = " & strCustID 
	rsWIDs.Open sqlWID, g_strCONN, 3, 1
	If Not rsWIDs.EOF Then
		GetCIDCustID = rsWIDs("Medicaid_Number")
	End If
	rsWIDs.Close
	Set rsWIDs = Nothing
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
Function IsWorker(xxx, yyy)
	IsWorker = False
	Set rsWor = Server.CreateObject("ADODB.RecordSet")
	sqlWor = "SELECT * FROM ConWork_T WHERE CID = '" & xxx & "' AND WID = '" & yyy & "'"
	rsWor.Open sqlWor, g_strCONN, 3, 1
	If Not rsWor.EOF Then
		IsWorker = True
	End If
	rsWor.Close
	Set rsWor = Nothing
End Function
Function GetWIDSSN(strSSN)
	Set rsWID = Server.CreateObject("ADODB.RecordSet")
	sqlWID = "SELECT [index] FROM Worker_T WHERE Social_Security_Number = '" & strSSN & "'"
	rsWID.Open sqlWID, g_strCONN, 3, 1
	If Not rsWID.EOF Then
		GetWIDSSN = rsWID("index")
	End If
	rsWID.Close
	Set rsWID = Nothing
End Function
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
	Server.ScriptTimeout = 10800
	set oUpload = Server.CreateObject("Chilkat.UploadRcv")
	oUpload.SaveToUploadDir = 0
	oUpload.UploadDir = uploadpath
	success = oUpload.Consume()
	If success = 1 Then
		nFileName = MakeNewFileName() & ".CSV"
		oUpload.SetFilename 0, nFileName
		oUpload.SaveNthToUploadDir(0)
		'read file
		TimeStamp = Now
		Set fso = CreateObject("Scripting.FileSystemObject")
		Set oCSV = fso.OpenTextFile(uploadPath & nFileName, 1)'nFileName fso.OpenTextFile(uploadPath & "test.csv", 1)
		Set oLine = fso.CreateTextFile("C:\Work\lss-dbvortex\log\importlines.txt", true)
		ctr = 1
		Do Until oCSV.AtEndofStream
			'read csv
			strLine = oCSV.ReadLine
			If Trim(strLine) <> "" Then
				csvValue = Split(strLine, ",")
				'clean values
				ctr2 = 0
				Do Until ctr2 = Ubound(csvValue) + 1
					csvValue(ctr2) = Replace(csvValue(ctr2),chr(34), "")
					ctr2 = ctr2 + 1
				Loop
				'CHECK VALUES
				If Ubound(csvValue) < 9 Then
					Session("MSG") = Session("MSG") & "ERROR: Incorrect CSV format (LINE: " & ctr & "). Succeeding rows were not saved.<br>"
					Response.Redirect "import.asp"
				End If
				If csvValue(0) = "" Then 
					Session("MSG") = Session("MSG") & "ERROR: First value of row is invalid (LINE: " & ctr & "). Succeeding rows were not saved.<br>"
					Response.Redirect "import.asp"
				End If
				If Not IsDate(csvValue(1)) Then
					Session("MSG") = Session("MSG") & "ERROR: Second value of row is invalid (LINE: " & ctr & "). Succeeding rows were not saved.<br>"
					Response.Redirect "import.asp"
				End If
				If Not IsDate(csvValue(2)) Then
					Session("MSG") = Session("MSG") & "ERROR: Third value of row is invalid (LINE: " & ctr & "). Succeeding rows were not saved.<br>"
					Response.Redirect "import.asp"
				End If
				If Not IsDate(csvValue(4)) Then
					Session("MSG") = Session("MSG") & "ERROR: Fifth value of row is invalid (LINE: " & ctr & "). Succeeding rows were not saved.<br>"
					Response.Redirect "import.asp"
				End If
				If csvValue(5) = "" Then 
					Session("MSG") = Session("MSG") & "ERROR: Sixth value of row is invalid (LINE: " & ctr & "). Succeeding rows were not saved.<br>"
					Response.Redirect "import.asp"
				End If
				If csvValue(7) = "" Then 
					Session("MSG") = Session("MSG") & "ERROR: Eighth value of row is invalid (LINE: " & ctr & "). Succeeding rows were not saved.<br>"
					Response.Redirect "import.asp"
				End If
				'''''''''''''
				WID = GetWIDBadge(csvValue(0))
				TSDate = CDate(csvValue(1))
				TSTimeIn = csvValue(1) & " " & RoundTime(CDate(csvValue(2)))
				TSTimeOut = csvValue(3) & " " & RoundTime(CDate(csvValue(4)))
				CID = GetCIDCustID(csvValue(5))
				PhoneNum = csvValue(7)
				PhoneNum2 = csvValue(8)
				ActvityCode = ""
				
				'response.write "CID: " & CID & " CSV:" & csvValue(5) & "<br>"
				ctr3 = 9
				Do Until ctr3 = Ubound(csvValue) + 1
					If csvValue(ctr3) <> "" Then ActvityCode = ActvityCode & csvValue(ctr3) & ","
					ctr3 = ctr3 + 1
				Loop
				If 	IsWorker(CID, GetWIDSSN(WID)) Or IsFloat(WID) Then
					'Get Total Hours for week 
					Set tblCon = Server.CreateObject("ADODB.RecordSet")
					sqlCon = "SELECT MaxHrs FROM Consumer_t WHERE medicaid_number = '" & CID & "' "
					tblCon.Open sqlCon, g_strCONN, 1, 3
					If Not tblCon.EOF Then
						tmpMax = Z_CZero(tblCon("MaxHrs"))
					End If
					tblCon.Close
					Set tblCon = Nothing
					'GET EXISTING HOURS
					Set rsChk = Server.CreateObject("ADODB.RecordSet")
					sqlChk = "SELECT * FROM tSHEETS_T WHERE date = '" & GetSundayTS(TSDate) & "' AND client = '" & _
						CID & "' AND Ext = 0"
					rsChk.Open sqlChk, g_strCONN, 3, 1
					HrsCon = 0
					If Not rsChk.EOF Then
						Do Until rsChk.EOF
							HrsCon = HrsCon + Z_CZero(rsChk("mon")) + Z_CZero(rsChk("tue")) + Z_CZero(rsChk("wed")) + Z_CZero(rsChk("thu")) + Z_CZero(rsChk("fri")) + _
								Z_CZero(rsChk("sat")) + Z_CZero(rsChk("sun"))
							rsChk.MoveNext
						Loop
					End If	
					rsChk.Close
					Set rsChk = Nothing
					tmpHRS = 0
					'Session("MSG") = Session("MSG") & "LINE: " & ctr & " -- Tin: " & TSTimeIn & " Tout: " & TSTimeOut & "<br>"
					tmpHRS = GetHrs(TSTimeIn, TSTimeOut) + HrsCon
					If tmpHRS > tmpMax Then
						'save db EXTENDED
						remHrs = tmpMax - HrsCon ' remaining hrs available
						Set rsTS = Server.CreateObject("ADODB.RecordSet")
						sqlTS = "SELECT * FROM Tsheets_T WHERE timestamp = '" & Now & "'"
						rsTS.Open sqlTS, g_strCONN, 1, 3
						rsTS.AddNew 'main TS
						rsTS("emp_ID") = WID
						rsTS("Client") = CID
						rsTS("Date") = GetSundayTS(TSDate)
						If WeekDay(TSDate) = 2 Then rsTS("mon") = remHrs'GetHrs(TSTimeIn, TSTimeOut)
						If WeekDay(TSDate) = 3 Then rsTS("tue") = remHrs'GetHrs(TSTimeIn, TSTimeOut)
						If WeekDay(TSDate) = 4 Then rsTS("wed") = remHrs'GetHrs(TSTimeIn, TSTimeOut)
						If WeekDay(TSDate) = 5 Then rsTS("thu") = remHrs'GetHrs(TSTimeIn, TSTimeOut)
						If WeekDay(TSDate) = 6 Then rsTS("fri") = remHrs'GetHrs(TSTimeIn, TSTimeOut)
						If WeekDay(TSDate) = 7 Then rsTS("sat") = remHrs'GetHrs(TSTimeIn, TSTimeOut)
						If WeekDay(TSDate) = 1 Then rsTS("sun") = remHrs'GetHrs(TSTimeIn, TSTimeOut)
						rsTS("author") = Session("UserID")
						rsTS("EXT") = False
						rsTS("Timestamp") = TimeStamp
						rsTS("misc_notes") = Trim(ActvityCode)
						rsTS("CallerID") = PhoneNum
						rsTS("CallerID2") = PhoneNum2
						rsTS("timein") = TSTimeIn
						rsTS("timeout") = TSTimeOut
						rsTS.Update
						rsTS.Close
						Set rsTS = Nothing
						extHrs = GetHrs(TSTimeIn, TSTimeOut) - remHrs
						Set rsTS = Server.CreateObject("ADODB.RecordSet")
						sqlTS = "SELECT * FROM Tsheets_T WHERE timestamp = '" & Now & "'"
						rsTS.Open sqlTS, g_strCONN, 1, 3
						rsTS.AddNew 'EXT TS
						rsTS("emp_ID") = WID
						rsTS("Client") = CID
						rsTS("Date") = GetSundayTS(TSDate)
						If WeekDay(TSDate) = 2 Then rsTS("mon") = extHrs'GetHrs(TSTimeIn, TSTimeOut)
						If WeekDay(TSDate) = 3 Then rsTS("tue") = extHrs'GetHrs(TSTimeIn, TSTimeOut)
						If WeekDay(TSDate) = 4 Then rsTS("wed") = extHrs'GetHrs(TSTimeIn, TSTimeOut)
						If WeekDay(TSDate) = 5 Then rsTS("thu") = extHrs'GetHrs(TSTimeIn, TSTimeOut)
						If WeekDay(TSDate) = 6 Then rsTS("fri") = extHrs'GetHrs(TSTimeIn, TSTimeOut)
						If WeekDay(TSDate) = 7 Then rsTS("sat") = extHrs'GetHrs(TSTimeIn, TSTimeOut)
						If WeekDay(TSDate) = 1 Then rsTS("sun") = extHrs'GetHrs(TSTimeIn, TSTimeOut)
						rsTS("author") = Session("UserID")
						rsTS("EXT") = True
						rsTS("Timestamp") = TimeStamp
						rsTS("misc_notes") = Trim(ActvityCode)
						rsTS("CallerID") = PhoneNum
						rsTS("CallerID2") = PhoneNum2
						rsTS("timein") = TSTimeIn
						rsTS("timeout") = TSTimeOut
						rsTS.Update
						rsTS.Close
						Set rsTS = Nothing
					Else
						'save db main TS
						Set rsTS = Server.CreateObject("ADODB.RecordSet")
						sqlTS = "SELECT * FROM Tsheets_T WHERE timestamp = '" & Now & "'"
						rsTS.Open sqlTS, g_strCONN, 1, 3
						rsTS.AddNew 'main TS
						rsTS("emp_ID") = WID
						rsTS("Client") = CID
						rsTS("Date") = GetSundayTS(TSDate)
						If WeekDay(TSDate) = 2 Then rsTS("mon") = GetHrs(TSTimeIn, TSTimeOut)
						If WeekDay(TSDate) = 3 Then rsTS("tue") = GetHrs(TSTimeIn, TSTimeOut)
						If WeekDay(TSDate) = 4 Then rsTS("wed") = GetHrs(TSTimeIn, TSTimeOut)
						If WeekDay(TSDate) = 5 Then rsTS("thu") = GetHrs(TSTimeIn, TSTimeOut)
						If WeekDay(TSDate) = 6 Then rsTS("fri") = GetHrs(TSTimeIn, TSTimeOut)
						If WeekDay(TSDate) = 7 Then rsTS("sat") = GetHrs(TSTimeIn, TSTimeOut)
						If WeekDay(TSDate) = 1 Then rsTS("sun") = GetHrs(TSTimeIn, TSTimeOut)
						rsTS("author") = Session("UserID")
						rsTS("EXT") = False
						rsTS("Timestamp") = TimeStamp
						rsTS("misc_notes") = Trim(ActvityCode)
						rsTS("CallerID") = PhoneNum
						rsTS("CallerID2") = PhoneNum2
						rsTS("timein") = TSTimeIn
						rsTS("timeout") = TSTimeOut
						rsTS.Update
						rsTS.Close
						Set rsTS = Nothing
						Set rsTS = Server.CreateObject("ADODB.RecordSet")
						sqlTS = "SELECT * FROM Tsheets_T WHERE timestamp = '" & Now & "'"
						rsTS.Open sqlTS, g_strCONN, 1, 3
						rsTS.AddNew 'EXT TS
						rsTS("emp_ID") = WID
						rsTS("Client") = CID
						rsTS("Date") = GetSundayTS(TSDate)
						rsTS("author") = Session("UserID")
						rsTS("EXT") = True
						rsTS("Timestamp") = TimeStamp
						rsTS("misc_notes") = Trim(ActvityCode)
						rsTS("CallerID") = PhoneNum
						rsTS("CallerID2") = PhoneNum2
						rsTS("timein") = TSTimeIn
						rsTS("timeout") = TSTimeOut
						rsTS.Update
						rsTS.Close
						Set rsTS = Nothing
					End If
				Else
					Session("MSG") = Session("MSG") & "ERROR: Worker is not associated with consumer or not a float worker (LINE: " & ctr & ").<br>"
				End If
			End If
			oLine.WriteLine Now & ":: " & ctr
			ctr = ctr + 1
		Loop	
		
		'move file for backup
		'fso.MoveFile uploadPath & nFileName, BackupDest
		Session("MSG") = Session("MSG") & "Data Saved."
		oLine.WriteLine Now & ":: " & Session("MSG")
		oLine.Close
		Set oLine = Nothing
		Set fso = Nothing

	else
		 Response.Write oUpload.LastErrorHtml
	End If
	Set oUpload = Nothing
	
End If
%>
<html>
	<head>
		<title>LSS - In-Home Care - Vortex Import</title>
		<script language='JavaScript'>
			
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
	<body bgcolor='white' LEFTMARGIN='0' TOPMARGIN='0'>
		<form method="POST" enctype="multipart/form-data">
			<!-- #include file="_boxup.asp" -->
			<!-- #include file="_NavHeader.asp" -->
			<table border='0' align='center'>
				<tr><td colspan='2' align='center'><font color='red'  face='trebuchet MS' size='1'>&nbsp;<%=Session("MSG")%>&nbsp;</font></td></tr>
				<tr>
					<td align='center'>
						<input type="file" name="F1" size="20">
					</td>
				</tr>
				<td align='center'>
						<input type='submit' value='Import Timesheets' style='width: 200px;' class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'">
					</td>
				</tr>
			</table>
			<!-- #include file="_boxdown.asp" -->
		</form>
	</body>
</html>
<%
Session("MSG") = "" 
%>