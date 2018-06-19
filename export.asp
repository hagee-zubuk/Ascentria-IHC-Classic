<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="export_helper.asp" -->
<%

Server.ScriptTimeout = 360000

If Request("exp") = 1 Then
	temp = Z_DoDeCrypt(Request("emp"))
	tclient = Z_DoDeCrypt(Request("cli"))
	lname = split(temp,",")
	Set fso = CreateObject("Scripting.FileSystemObject")
	tmpdate = split(datevalue(date),"/")
	If len(tmpdate(0)) = 1 Then
		tmpMonth = 0 & tmpdate(0)
	Else
		tmpMonth = tmpdate(0)
	End If
	If len(tmpdate(1)) = 1 Then
		tmpDay = 0 & tmpdate(1)
	Else
		tmpDay = tmpdate(1)
	End if
	rep = PayCSV & lname(0) & tmpMonth & tmpDay & tmpdate(2) & ".csv"
	Set Prt = fso.CreateTextFile(rep, True)
	Prt.WriteLine "HTS, " & lname(0) & ", User, D"
	Prt.WriteLine "DTSEARN, " & lname(0) & ", Wage, " & Request("thrs")
	Prt.Close
	Set fso = Nothing
	Session("MSG") = "Report exported."
	copypath = copyfile & Z_GetFilename(rep)
	Set Prt2 = fso.CreateTextFile(copypath, True)
	Prt2.WriteLine "HTS, " & lname(0) & ", User, D"
	Prt2.WriteLine "DTSEARN, " & lname(0) & ", Wage, " & Request("thrs")
	Session("dload") = Z_DoEncrypt(Z_GetFilename(rep))
	Response.Redirect "extra2.asp"
ElseIf Request("sql") = 1 Then
	set tblrep = Server.CreateObject("ADODB.RecordSet")
	Set fso = CreateObject("Scripting.FileSystemObject")
	Wk = Session("Wk")
	sqlRep = Z_DoDecrypt(Session("tmpSQL"))
	
	tblREP.Open sqlREP, g_strCONN, 1, 3
	tmpID2 = ""
	marker = 0
	If Not tblREP.EOF Then
		SHours = 0
		tblREP.MoveFirst
		Do Until tblREP.EOF
			If Session("type") = 1 Then
				tmpID = tblREP("emp_id")
				If tmpID <> tmpID2 Or tblREP.EOF Then
					THours = (Tmon) + (Ttue) + (Twed) + (Tthur) + (Tfri) + (Tsat) + (Tsun)
					Tmon = 0
					Ttue = 0
					Twed = 0
					Tthur = 0
					Tfri = 0
					Tsat = 0
					Tsun = 0
					If tmpID2 <> "" Then 
						Set tblname = Server.CreateObject("ADODB.RecordSet")
						sqlName = "SELECT * FROM Worker_t"
						tblname.Open sqlName, g_strCONN, 1, 3
						strTmp = "Social_Security_Number= '" & tmpID2 & "' "
						tblname.Find(strTmp)
						If Not tblname.EOF Then name = tblname("lname") & ", "	& tblname("fname")
						tblname.Close
						set tblname = Nothing	
								strTBL = strTBL & Wk & "," & Right(tmpID2, 4) & "," & _
									name & ";" & _
									Z_FormatNumber(THours,2) & vbCrLf
									SHours = Shours + THours
					End If
					tmpID2 = tmpID
				End If
				Tmon = (Tmon) + (tblREP("mon"))
				Ttue = (Ttue) + (tblREP("tue"))
				Twed = (Twed) + (tblREP("wed"))
				Tthur = (Tthur) + (tblREP("thu"))
				Tfri = (Tfri) + (tblREP("fri"))
				Tsat = (Tsat) + (tblREP("sat"))
				Tsun = (Tsun) + (tblREP("sun"))
			ElseIf Session("type") = 2 Then
				'tmpID2 = "0"
				tmpID = tblREP("medicaid_number")
				If tmpID <> tmpID2 Or tblREP.EOF Then
					THours = (Tmon) + (Ttue) + (Twed) + (Tthur) + (Tfri) + (Tsat) + (Tsun)
					Tmon = 0
					Ttue = 0
					Twed = 0
					Tthur = 0
					Tfri = 0
					Tsat = 0
					Tsun = 0
					'response.write "ID: " & tmpID2 & "<br>"
					'if tmpID2 = "" Then tmpID2 = "0"
					If tmpID2 <> "" Then 
						'If THours <> 0 Then
							Set tblname = Server.CreateObject("ADODB.RecordSet")
							sqlName = "SELECT * FROM Consumer_t"
							tblname.Open sqlName, g_strCONN, 1, 3
							strTmp = "medicaid_number= '" & tmpID2 & "' "
							tblname.Find(strTmp)
							If Not tblname.EOF Then name = tblname("lname") & ", "	& tblname("fname")
							tblname.Close
							set tblname = Nothing
									strTBL = strTBL & wk & "," & tmpID2 & "," & name & "," & Z_FormatNumber(THours,2) & vbCrLf
							SHours = Shours + THours
						'End If
					End If
					tmpID2 = tmpID
				End If
				Tmon = (Tmon) + (tblREP("mon"))
				Ttue = (Ttue) + (tblREP("tue"))
				Twed = (Twed) + (tblREP("wed"))
				Tthur = (Tthur) + (tblREP("thu"))
				Tfri = (Tfri) + (tblREP("fri"))
				Tsat = (Tsat) + (tblREP("sat"))
				Tsun = (Tsun) + (tblREP("sun"))
			Else
				Tmon = (Tmon) + (tblREP("mon"))
				Ttue = (Ttue) + (tblREP("tue"))
				Twed = (Twed) + (tblREP("wed"))
				Tthur = (Tthur) + (tblREP("thu"))
				Tfri = (Tfri) + (tblREP("fri"))
				Tsat = (Tsat) + (tblREP("sat"))
				Tsun = (Tsun) + (tblREP("sun"))
			End If
			tblREP.MoveNext
		Loop 
		THours = (Tmon) + (Ttue) + (Twed) + (Tthur) + (Tfri) + (Tsat) + (Tsun)
		Tmon = 0
		Ttue = 0
		Twed = 0
		Tthur = 0
		Tfri = 0
		Tsat = 0
		Tsun = 0
		
		If Session("type") = 1 Then 
			Set tblname = Server.CreateObject("ADODB.RecordSet")
						sqlName = "SELECT * FROM Worker_t"
						tblname.Open sqlName, g_strCONN, 1, 3
						strTmp = "Social_Security_Number= '" & tmpID2 & "' "
						tblname.Find(strTmp)
						If Not tblname.EOF Then name = tblname("lname") & ", "	& tblname("fname")
						tblname.Close
						set tblname = Nothing
				strTBL = strTBL & wk & "," & right(tmpID2, 4) & "," & name & "," & _
							Z_FormatNumber(THours,2) & vbCrLf
							SHours = Shours + THours
							strHEAD = "Week, SSN, Last Name, First Name, Hours" & vbcrlf
		ElseIf Session("type") = 2 Then
				Set tblname = Server.CreateObject("ADODB.RecordSet")
						sqlName = "SELECT * FROM Consumer_t"
						tblname.Open sqlName, g_strCONN, 1, 3
						strTmp = "medicaid_number= '" & tmpID & "' "
						tblname.Find(strTmp)
						If Not tblname.EOF Then name = tblname("lname") & ", "	& tblname("fname")
						tblname.Close
						set tblname = Nothing
						If Session("lngType") = 1 Or Session("lngType") = 2 Then 
							strTBL = strTBL & wk & "," & tmpID & "," & name & "," & Z_FormatNumber(THours,2) & vbCrLf
						Else
							strTBL = strTBL & wk & ",," & name & "," & Z_FormatNumber(THours,2) & vbCrLf
						End If	
							SHours = Shours + THours
							strHEAD = "Week, Medicaid #, Last Name, First Name, Hours" & vbcrlf
			'End If
		Else
			marker = 1
		End If
		'strTBL = strTBL & "<tr><td align='right' colspan='3'><b>Total Hours:&nbsp;</b></td><td align='center'><b>" & _
		'	Z_FormatNumber(SHours,2) & "</b></td><tr>"
		
	End If
	tblREP.Close
	set tblREP = Nothing	
	Set Prt = fso.CreateTextFile(RepCSV, True)
	Prt.WriteLine strHEAD
	Prt.WriteLine strTBL
	Prt.Close
	
	copypath = copyfile & Z_GetFilename(RepCSV)
	Set Prt2 = fso.CreateTextFile(copypath, True)
	Prt2.WriteLine strHEAD
	Prt2.WriteLine strTBL
	Session("dload") = Z_DoEncrypt(Z_GetFilename(RepCSV))
	
	Set fso = Nothing
	Session("MSG") = "Table exported."
	Response.Redirect "extra2.asp"
ElseIf Request("sql") = 2 Then
		Set fso = CreateObject("Scripting.FileSystemObject")
		
		'Session("MSG") = "Table exported."
		Session("dload") = Session("ProcCSV")
		Session("dload2") = Session("BillCSV")
		Response.Redirect "extra.asp"
		
	ElseIf Request("sql") = 3 Then
		Set fso = CreateObject("Scripting.FileSystemObject")
		Set rsTBL = Server.CreateObject("ADODB.RecordSet")
		sqlTBL = "SELECT DOB, fname, lname, address, city, state, active FROM consumer_t, c_status_t " & _
			"WHERE Consumer_t.medicaid_number = c_status_t.medicaid_number AND Active = 1 ORDER BY Month(DOB), Day(DOB)"
		rsTBL.Open sqlTBL, g_strCONN, 1, 3
		strProcH = "Date OF Birth, Last Name, First Name, Address, City, State"
		Do Until rsTBL.EOF	
			strProcB = strProcB & rsTBL("DOB") & ", """ & rsTBL("lname") & _
				""",""" & rsTBL("fname") & """,""" & rsTBL("Address") & """, " & rsTBL("City") & ", " & rsTBL("State") & vbcrlf
			rsTBL.MoveNext
		Loop
		Set Prt = fso.CreateTextFile(ConDOB, True)
		Prt.WriteLine strProcH
		Prt.WriteLine strProcB
		
		copypath = copyfile & Z_GetFilename(ConDOB)
		Set Prt2 = fso.CreateTextFile(copypath, True)
		Prt2.WriteLine strProcH
		Prt2.WriteLine strProcB
		Session("dload") = Z_DoEncrypt(Z_GetFilename(ConDOB))
		rsTBL.Close
		Set rsTBL = Nothing
		'Session("dload") = Z_DoEncrypt(conDOB)
		Response.Redirect "extra2.asp"
	ElseIf Request("sql") = 4 Then
		Set fso = CreateObject("Scripting.FileSystemObject")
		Set rsTBL = Server.CreateObject("ADODB.RecordSet")
		sqlTBL = "SELECT fname, lname, maddress, mcity, mstate, mzip FROM worker_T " & _
			"WHERE Status = 'Active' ORDER BY lname, fname"
		rsTBL.Open sqlTBL, g_strCONN, 1, 3
		strProcH = "Last Name, First Name, Address, City, State, Zip"
		Do Until rsTBL.EOF	
			strProcB = strProcB & """" & rsTBL("lname") & _
				""", " & rsTBL("fname") & ",""" & rsTBL("mAddress") & """," & rsTBL("mCity") & ", " & rsTBL("mState") & ", " & rsTBL("mZip") & vbcrlf
			rsTBL.MoveNext
		Loop
		Set Prt = fso.CreateTextFile(WorMailAdr, True)
		Prt.WriteLine strProcH
		Prt.WriteLine strProcB
		rsTBL.Close
		Set rsTBL = Nothing
		copypath = copyfile & Z_GetFilename(WorMailAdr)
		Set Prt2 = fso.CreateTextFile(copypath, True)
		Prt2.WriteLine strProcH
		Prt2.WriteLine strProcB
		Session("dload") = Z_DoEncrypt(Z_GetFilename(WorMailAdr))
		'Session("dload") = Z_DoEncrypt(WorMailAdr)
		Response.Redirect "extra2.asp"
ElseIf Request("sql") = 5 Then
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set rsTBL = Server.CreateObject("ADODB.RecordSet")
	sqlTBL = "SELECT lname, fname, maddress, mcity, mstate, mzip, DOB, PhoneNo, c.[Medicaid_Number], gender, TermDate, l.[Language] " & _
			"FROM Consumer_t AS c " & _
			"INNER JOIN C_Status_t AS s ON c.[Medicaid_number] = s.[Medicaid_number] " & _
			"INNER JOIN language_T AS l ON c.[langid]=l.[index] " & _
			"WHERE Active = 1 " & _
			"AND onHold <> 1 " & _
			"AND TermDate IS NULL "
	If Request("selrh") > 0 Then sqlTBL = sqlTBL & "AND PMID = " & Request("selrh") & " "
	sqlTBL = sqlTBL & "ORDER BY lname, fname"
	rsTBL.Open sqlTBL, g_strCONN, 1, 3
	strProcH = "Last Name, First Name, Mailing Address, City, State, Zip, Phone, DOB, Medicaid Number, Gender, Language"
	Do Until rsTBL.EOF	
		If Session("lngType") = 1 Or Session("lngType") = 2 Then 
			strProcB = strProcB & """" & rsTBL("lname") & _
				""",""" & rsTBL("fname") & """,""" & rsTBL("mAddress") & """,""" & rsTBL("mCity") & """,""" & rsTBL("mState") & """,""" & rsTBL("mZIP") & """,""" & rsTBL("PhoneNo") & """,""" & rsTBL("DOB") & """,""" & rsTBL("Medicaid_Number") & """,""" & rsTBL("Gender") & _
					""",""" & rsTBL("language") & """" & vbcrlf
		Else
			strProcB = strProcB & """" & rsTBL("lname") & _
				""",""" & rsTBL("fname") & """,""" & rsTBL("mAddress") & """,""" & rsTBL("mCity") & """,""" & rsTBL("mState") & """,""" & rsTBL("mZIP") & """,""" & rsTBL("PhoneNo") & """,""" & rsTBL("DOB") & """,""" & "" & """,""" & rsTBL("Gender") & """,""" & _
					rsTBL("language") & """" & vbcrlf
		End If
		rsTBL.MoveNext
	Loop
	Set Prt = fso.CreateTextFile(ActiveCon, True)
	Prt.WriteLine strProcH
	Prt.WriteLine strProcB
	rsTBL.Close
	Set rsTBL = Nothing
	copypath = copyfile & Z_GetFilename(ActiveCon)
	Set Prt2 = fso.CreateTextFile(copypath, True)
	Prt2.WriteLine strProcH
	Prt2.WriteLine strProcB
	Session("dload") = Z_DoEncrypt(Z_GetFilename(ActiveCon))
	'Session("dload") = Z_DoEncrypt(ActiveCon)
	Response.Redirect "extra2.asp"	
ElseIf Request("sql") = 6 Then
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set rsTBL = Server.CreateObject("ADODB.RecordSet")
	sqlTBL = "SELECT lname, fname, maddress, mcity, mstate, mzip, DOB, PhoneNo, FileNum, language " & _
			"FROM Worker_t AS w " & _
			"INNER JOIN language_T AS l ON w.[langid]=l.[index] " & _
			"WHERE status = 'Active' AND term_date IS NULL "
	If Request("selrh") > 0 Then sqlTBL = sqlTBL & "AND PM1 = " & Request("selrh") & " "
	sqlTBL = sqlTBL & "ORDER BY lname, fname"
	rsTBL.Open sqlTBL, g_strCONN, 1, 3
	strProcH = "Last Name, First Name, Mailing Address, City, State, Zip, Phone No., DOB, File Number, Language"
	Do Until rsTBL.EOF	
		strProcB = strProcB & """" & rsTBL("lname") & _
				""",""" & rsTBL("fname") & """,""" & rsTBL("mAddress") & """, " & rsTBL("mCity") & ", " & rsTBL("mState") & _
				", " & rsTBL("mZIP") & ", " & rsTBL("PhoneNo") & ", " & rsTBL("DOB") & ", " & rsTBL("FileNum") & ",""" & _
				rsTbl("Language") & """" & vbCrLf
		rsTBL.MoveNext
	Loop
	Set Prt = fso.CreateTextFile(ActiveWor, True)
	Prt.WriteLine strProcH
	Prt.WriteLine strProcB
	rsTBL.Close
	Set rsTBL = Nothing
	copypath = copyfile & Z_GetFilename(ActiveWor)
	Set Prt2 = fso.CreateTextFile(copypath, True)
	Prt2.WriteLine strProcH
	Prt2.WriteLine strProcB
	Session("dload") = Z_DoEncrypt(Z_GetFilename(ActiveWor))
	'Session("dload") = Z_DoEncrypt(ActiveWor)
	Response.Redirect "extra2.asp"	
	ElseIf Request("sql") = 7 Then
		Set fso = CreateObject("Scripting.FileSystemObject")
		Set rsTBL = Server.CreateObject("ADODB.RecordSet")
		sqlTBL = "SELECT Representative_T.lname as rlname, Representative_T.fname as rfname, Representative_T.address as raddress, Representative_T.city as rcity, Representative_T.state as rstate, Representative_T.zip as rzip, Consumer_T.lname as clname, Consumer_T.fname as cfname, pmid " & _
					"FROM Representative_T, consumer_T, Conrep_T, C_Status_t " & _
					"WHERE RID = Representative_T.[index] " & _
					"AND CID = Consumer_t.medicaid_Number " & _ 
					"AND C_Status_t.Medicaid_number = Consumer_t.Medicaid_number AND Active = 1 " & _
					"ORDER BY Representative_T.Lname, Representative_T.Fname"
		rsTBL.Open sqlTBL, g_strCONN, 1, 3
		strProcH = "Representative Last Name, Representative First Name, Address, City, State, Zip, Consumer Last Name, Consumer First Name,RCC Last Name, RCC First Name"
		Do Until rsTBL.EOF	
			strProcB = strProcB & """" & rsTBL("rlname") & _
				""",""" & rsTBL("rfname") & """,""" & rsTBL("rAddress") & """,""" & rsTBL("rCity") & """,""" & rsTBL("rState") & """,""" & rsTBL("rZIP") & """,""" & rsTBL("clname") & """,""" & rsTBL("cfname") & """,""" & GetCM(rsTBL("PMID")) & """" & vbcrlf
			rsTBL.MoveNext
		Loop
		Set Prt = fso.CreateTextFile(RepList, True)
		Prt.WriteLine strProcH
		Prt.WriteLine strProcB
		rsTBL.Close
		Set rsTBL = Nothing
		copypath = copyfile & Z_GetFilename(RepList)
		Set Prt2 = fso.CreateTextFile(copypath, True)
		Prt2.WriteLine strProcH
		Prt2.WriteLine strProcB
		Session("dload") = Z_DoEncrypt(Z_GetFilename(RepList))
		'Session("dload") = Z_DoEncrypt(RepList)
		Response.Redirect "extra2.asp"	
ElseIf Request("sql") = 8 Then
	Set fso = CreateObject("Scripting.FileSystemObject")
	strProcH = "Last Name,First Name,Mailing Address,City,State,Zip,Phone Num.,RIHCC Last Name, RIHCC First Name, Language"
	Set rsTBL = CreateObject("ADODB.RecordSet")
	sqlTBL = "SELECT c.lname as clname, c.fname as cfname, maddress, mcity, language" & _
			", mstate, mzip, phoneno, pmid, p.[lname] AS plname, p.[fname] AS pfname, p.[id] AS pid " & _
			"FROM Consumer_t AS c " & _
			"INNER JOIN C_Status_t AS s ON c.[medicaid_number] = s.[medicaid_number] " & _
			"INNER JOIN Proj_man_T AS p ON c.[PMID] = p.[ID] " & _ 
			"INNER JOIN language_T AS l ON c.[langid]=l.[index] " & _
			"WHERE Active=1 "
	If Request("selrh") > 0 Then sqlTBL = sqlTBL & "AND PMID = " & Request("selrh") & " "
	sqlTBL = sqlTBL & "ORDER BY p.Lname, p.Fname, c.lname, c.fname, City"
	rsTBL.Open sqlTBL, g_strCONN, 1, 3
	Do Until rsTBL.EOF
		strProcB = strProcB & """" & rsTBL("clname") & """,""" & rsTBL("cfname") & """,""" & rsTBL("mAddress") & """,""" & _
			rsTBL("mCity") & """,""" & rsTBL("mState") & """,""" & rsTBL("mZip") & """,""" & rsTBL("PhoneNo") & """,""" & _
			rsTBL("plname") & """,""" & rsTBL("pfname") & """,""" & rsTBL("language") & """" & vbCrLf
		rsTBL.MoveNext
	Loop
	rsTBL.Close
	Set rsTBL = Nothing
	Set Prt = fso.CreateTextFile(ConPMTown, True)
	Prt.WriteLine strProcH
	Prt.WriteLine strProcB
	copypath = copyfile & Z_GetFilename(ConPMTown)
	Set Prt2 = fso.CreateTextFile(copypath, True)
	Prt2.WriteLine strProcH
	Prt2.WriteLine strProcB
	Session("dload") = Z_DoEncrypt(Z_GetFilename(ConPMTown))
	'Session("dload") = Z_DoEncrypt(ConPMTown)
	Response.Redirect "extra2.asp"
ElseIf Request("sql") = 9 Then
	Set fso = CreateObject("Scripting.FileSystemObject")
	strProcH = "Last Name,First Name,Address,City,State,Zip,Phone Num.,RIHCC, RIHCC2, Language"
	Set rsTBL = CreateObject("ADODB.RecordSet")
	sqlTBL = "SELECT w.[fname], w.[lname], [maddress], [mcity], [mstate], [mzip], [phoneno], [language]" & _
			", COALESCE(p1.[lname], 'N/A') AS [p1lname], COALESCE(p1.[fname], '') AS [p1fname] " & _
			", COALESCE(p2.[lname], 'N/A') AS [p2lname], COALESCE(p2.[fname], '') AS [p2fname] " & _
			"FROM worker_t AS w " & _
			"INNER JOIN language_t AS l ON w.[langid]=l.[index] " & _
			"LEFT JOIN Proj_Man_T AS p1 ON w.[pm1]=p1.[id] " & _
			"LEFT JOIN Proj_Man_T AS p2 ON w.[pm2]=p2.[id] " & _
			"WHERE status = 'Active' "
	If Request("selrh") > 0 Then sqlTBL = sqlTBL & "AND PM1 = " & Request("selrh") & " "
	sqlTBL = sqlTBL & "ORDER BY lname, fname, mCity"
	rsTBL.Open sqlTBL, g_strCONN, 1, 3
	Do Until rsTBL.EOF
	
		strProcB = strProcB & """" & rsTBL("lname") & """,""" & rsTBL("fname") & """,""" & rsTBL("mAddress") & """,""" & _
			rsTBL("mCity") & """,""" & rsTBL("mState") & """,""" & rsTBL("mZip") & """,""" & rsTBL("PhoneNo") & """,""" & _
			rsTBL("p1lname") & ", " & rsTBL("p1fname") & """,""" & _
			rsTBL("p2lname") & ", " & rsTBL("p2fname") & """,""" & rsTBL("language") & """" & vbCrLf
		rsTBL.MoveNext
	Loop
	rsTBL.Close
	Set rsTBL = Nothing
	Set Prt = fso.CreateTextFile(WorPMTown, True)
	Prt.WriteLine strProcH
	Prt.WriteLine strProcB
	copypath = copyfile & Z_GetFilename(WorPMTown)
	Set Prt2 = fso.CreateTextFile(copypath, True)
	Prt2.WriteLine strProcH
	Prt2.WriteLine strProcB
	Session("dload") = Z_DoEncrypt(Z_GetFilename(WorPMTown))
	'Session("dload") = Z_DoEncrypt(WorPMTown)
	Response.Redirect "extra2.asp"
	ElseIf Request("sql") = 10 Then
		Set fso = CreateObject("Scripting.FileSystemObject")
		strProcH = "Last Name,First Name,Address,City,State,Zip,Agency,Consumer First Name,Consumer Last Name,RIHCC"
		Set rsTBL = CreateObject("ADODB.RecordSet")
		sqlTBL = "SELECT Case_manager_T.lname as cmlname, Case_manager_T.fname as cmfname, Case_manager_T.address as cmaddress, " & _
					"Case_manager_T.city as cmcity, Case_manager_T.state as cmstate, Case_manager_T.zip as cmzip, Case_manager_T.email as cmemail, faxno, agency, " & _
					"Consumer_T.lname as clname, Consumer_T.fname as cfname, pmid " & _
					"FROM Case_manager_T, consumer_T, CMCon_T " & _
					"WHERE CMID = Case_manager_T.[index] " & _
					"AND CID = Consumer_t.medicaid_Number " & _ 
					"ORDER BY Case_manager_T.Lname, Case_manager_T.Fname, consumer_T.lname, consumer_t.fname"
		rsTBL.Open sqlTBL, g_strCONN, 1, 3
		Do Until rsTBL.EOF
			strProcB = strProcB & """" & rsTBL("cmlname") & """,""" & rsTBL("cmfname") & """,""" & rsTBL("cmAddress") & """,""" & _
				rsTBL("cmCity") & """,""" & rsTBL("cmState") & """,""" & rsTBL("cmZip") & """,""" & rsTBL("Agency") & """,""" & rsTBL("clname") & """,""" & _
				rsTBL("clname") & """,""" & GetName3(rsTBL("PMID")) & """" & vbCrLf
			rsTBL.MoveNext
		Loop
		rsTBL.Close
		Set rsTBL = Nothing
		Set Prt = fso.CreateTextFile(CaseListCSV, True)
		Prt.WriteLine strProcH
		Prt.WriteLine strProcB
		copypath = copyfile & Z_GetFilename(CaseListCSV)
		Set Prt2 = fso.CreateTextFile(copypath, True)
		Prt2.WriteLine strProcH
		Prt2.WriteLine strProcB
		Session("dload") = Z_DoEncrypt(Z_GetFilename(CaseListCSV))
		'Session("dload") = Z_DoEncrypt(CaseListCSV)
		Response.Redirect "extra2.asp"	
	ElseIf Request("sql") = 11 Then
		Set fso = CreateObject("Scripting.FileSystemObject")
		If Request("myUri") = 0 Then
			strProcH = "Worker Name, File Number, RIHCC 1, RIHCC 2, Timesheet Week, Total Miles"
		Else
			strProcH = "Consumer Name, Mileage Cap, RIHCC,Worker Name, File Number, Timesheet Week, Total Miles"
		End If
		strProcH2 = "Co Code,Batch ID,File #,temp dept,temp rate,reg hours,o/t hours,hours 3 code,hours 3 amount,hours 4 code,hours 4 amount,earnings 3 code,earnings 3 amount,earnings 4 code,earnings 4 amount,earnings 5 code,earnings 5 amount,memo code,memo amount"
		Set rsTBL = CreateObject("ADODB.RecordSet")
		sqlTBL = "SELECT * FROM Tsheets_T, consumer_T, worker_T WHERE Medicaid_Number = client AND emp_ID = Worker_T.Social_Security_Number"
		If Request("fdate") <> "" Then
			sqlTBL = sqlTBL & " AND date >= '" & CDate(Request("fdate")) & "'" 
		End If
		If Request("tdate") <> "" Then
			sqlTBL = sqlTBL & " AND date  <= '" & CDate(Request("tdate")) & "'" 
		End If
		sqlTBL = sqlTBL & " ORDER BY consumer_T.lname, consumer_T.fname"
		rsTBL.Open sqlTBL, g_strCONN, 1, 3
		x= 0
		Do Until rsTBL.EOF
			totmile = Z_CZero(rsTBL("mile")) + Z_CZero(rsTBL("amile"))
				If totmile <> 0 Then
					If Request("myUri") = 0 Then
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
						'If Z_fixnull(rsTBL("ProcMile")) = "" Then
						'	rsTBL("ProcMile") = Date
						'	rsTBL.Update
						'End If
					Else
						conname = rsTBL("consumer_T.lname") & ", " & rsTBL("consumer_T.fname")
						worname = GetName(rsTBL("emp_ID"))
						APM = GetAPM2(rsTBL("client"))
						tmpTSWk1 = rsTBL("date")
						tmpTSWk2 = dateadd("d", 6, rsTBL("date"))'Cdate(rsTBL("date")) + 6
						tmpFileNum = GetFileNum(rsTBL("emp_id"))
						strProcB = strProcB & """" &  conname & """,""" & Z_CZero(rsTBL("consumer_T.milecap")) & """,""" & APM & _
							""",""" & worname & """,""" & tmpFileNum & """,""" & tmpTSWk1 & " - " & tmpTSWk2 & """,""" & totmile & """" & vbCrlf
						'If Z_fixnull(rsTBL("ProcMile")) = "" Then
						'	mileAmt = 0
						'	mileAmt = totmile * 0.47 'change to db rate
						'	strProcB2 = strProcB2 & "FJD,50," & tmpFileNum & ",,,,,,,,,,,EXP," & mileAmt & ",,,," & vbCrlf
						'	rsTBL("ProcMile") = Date
						'	rsTBL.Update
						'End If
					End If
				End If
			rsTBL.MoveNext
		Loop
		rsTBL.Close
		Set rsTBL = Nothing
		'not w/n 2 week
		If Request("fdate") <> "" Then
			Set rsMile = Server.CreateObject("ADODB.RecordSet")
			sqlMile = "SELECT * FROM Tsheets_T, consumer_T, worker_T WHERE Medicaid_Number = client AND emp_ID = Worker_T.Social_Security_Number " & _
				 "AND date < '" & CDate(Request("fdate")) & "' AND ProcMile IS NULL"
			If Request("myUri") = 0 Then
				sqlMile = sqlMile & " ORDER BY worker_T.lname, worker_T.fname"
				strProcH3 = "Worker Name, File Number, RIHCC 1, RIHCC 2, Timesheet Week, Total Miles"
			Else
				sqlMile = sqlMile & " ORDER BY consumer_T.lname, consumer_T.fname"
				strProcH3 = "Consumer Name, Mileage Cap, RIHCC,Worker Name, File Number, Timesheet Week, Total Miles"
			End If
		End If
		'sqlMile = sqlMile & " ORDER BY consumer_T.lname, consumer_T.fname"
		rsMile.Open sqlMile, g_strCONN, 1, 3
		x2 = 0
		'lngIdx = 0
		Do Until rsMile.eof
		totmile = Z_CZero(rsMile("mile")) + Z_CZero(rsMile("amile"))
			If totmile <> 0 Then
				If Request("myUri") = 0 Then
					myempid = rsMile("emp_ID")
					mycliid = rsMile("client")
					mydate = rsMile("date")
					mymiles = totmile
					mymilecap = rsMile("consumer_T.milecap")
					lngIdx = SearchArrays32(myempid)
					If lngIdx < 0 Then
						ReDim Preserve tmpemp2(x)
						ReDim Preserve tmpcon2(x)
						ReDim Preserve tmpDates2(x)
						ReDim Preserve tmpmile2(x)
						ReDim Preserve tmpmilecap2(x)

							
						tmpemp2(x2) = myempid
						tmpcon2(x2) = mycliid
						tmpDates2(x2) = mydate
						tmpmile2(x2) = mymiles
						tmpmilecap2(x2) = mymilecap
	
						x2 = x2 + 1
					Else
						tmpmile2(lngIdx) = tmpmile2(lngIdx) + mymiles
					End If
					'If Z_fixnull(rsTBL("ProcMile")) = "" Then
					'	rsTBL("ProcMile") = Date
					'	rsTBL.Update
					'End If
				Else
					conname = rsMile("consumer_T.lname") & ", " & rsMile("consumer_T.fname")
					worname = GetName(rsMile("emp_ID"))
					APM = GetAPM2(rsMile("client"))
					tmpTSWk1 = rsMile("date")
					tmpTSWk2 = dateadd("d", 6, rsMile("date"))'Cdate(rsTBL("date")) + 6
					tmpFileNum = GetFileNum(rsMile("emp_id"))
					strProcB3 = strProcB3 & """" &  conname & """,""" & Z_CZero(rsMile("consumer_T.milecap")) & """,""" & APM & _
						""",""" & worname & """,""" & tmpFileNum & """,""" & tmpTSWk1 & " - " & tmpTSWk2 & """,""" & totmile & """" & vbCrlf
					'If Z_fixnull(rsTBL("ProcMile")) = "" Then
					'	mileAmt = 0
					'	mileAmt = totmile * 0.47 'change to db rate
					'	strProcB2 = strProcB2 & "FJD,50," & tmpFileNum & ",,,,,,,,,,,EXP," & mileAmt & ",,,," & vbCrlf
					'	rsTBL("ProcMile") = Date
					'	rsTBL.Update
					'End If
				End If
			End If
			rsMile.MoveNext
		Loop
		rsMile.Close
		Set rsMile = Nothing
		'''''''
		If Request("myUri") = 0 Then
			y = 0
			Do Until y = x 
				tmpTSWk1 = Request("tdate")'tmpDates(y)
				tmpTSWk2 = Request("fdate") 'dateadd("d", 6, tmpDates(y))'Cdate(rsTBL("date")) + 6
				strProcB = strProcB & """" &  GetName(tmpemp(y)) & """,""" & GetFileNum(tmpemp(y)) & """,""" & GetName3(GetPM1(tmpemp(y))) & _
						""",""" & GetName3(GetPM2(tmpemp(y))) & """,""" & tmpTSWk1 & " - " & tmpTSWk2 & """,""" & tmpmile(y) & """" & vbCrlf
				'If rsTBL("ProcMile") = "" Then
				'	mileAmt = 0
				'	mileAmt = tmpmile(y) * 0.47 'change to db rate
				'	strProcB2 = strProcB2 & "FJD,50," & GetFileNum(tmpemp(y)) & ",,,,,,,,,,,EXP," & mileAmt & ",,,," & vbCrlf
				
				y = y + 1
			Loop 
			y = 0
			Do Until y = x2
				tmpTSWk1 = tmpDates2(y)
				tmpTSWk2 = dateadd("d", 6, tmpDates2(y))'Cdate(rsTBL("date")) + 6
				strProcB3 = strProcB3 & """" &  GetName(tmpemp2(y)) & """,""" & GetFileNum(tmpemp2(y)) & """,""" & GetName3(GetPM1(tmpemp2(y))) & _
						""",""" & GetName3(GetPM2(tmpemp2(y))) & """,""" & tmpTSWk1 & " - " & tmpTSWk2 & """,""" & tmpmile2(y) & """" & vbCrlf
				'If rsTBL("ProcMile") = "" Then
				'	mileAmt = 0
				'	mileAmt = tmpmile(y) * 0.47 'change to db rate
				'	strProcB2 = strProcB2 & "FJD,50," & GetFileNum(tmpemp(y)) & ",,,,,,,,,,,EXP," & mileAmt & ",,,," & vbCrlf
				
				y = y + 1
			Loop 
		End If
		If Request("prj") = 1 Then
			'ADP
			If Request("myUri") = 0 Then
				Set rsTBL = CreateObject("ADODB.RecordSet")
				sqlTBL = "SELECT * FROM Tsheets_T, consumer_T, worker_T WHERE procMile is null AND Medicaid_Number = client AND emp_ID = Worker_T.Social_Security_Number"
				If Request("fdate") <> "" Then
					sqlTBL = sqlTBL & " AND date >= '" & CDate(Request("fdate")) & "'" 
				End If
				If Request("tdate") <> "" Then
					sqlTBL = sqlTBL & " AND date  <= '" & CDate(Request("tdate")) & "'" 
				End If
				sqlTBL = sqlTBL & " ORDER BY consumer_T.lname, consumer_T.fname"
				
				rsTBL.Open sqlTBL, g_strCONN, 1, 3
				x2 = 0
				Do Until rsTBL.EOF
					totmile = Z_CZero(rsTBL("mile")) + Z_CZero(rsTBL("amile"))
					If totmile <> 0 Then
						If Request("myUri") = 0 Then
							myempid = rsTBL("emp_ID")
							mycliid = rsTBL("client")
							mydate = rsTBL("date")
							mymiles = totmile
							mymilecap = rsTBL("consumer_T.milecap")
							lngIdx = SearchArrays4(myempid)
							If lngIdx < 0 Then
								ReDim Preserve tmpempADP(x2)
								ReDim Preserve tmpmileADP(x2)
	
								tmpempADP(x2) = myempid
								tmpmileADP(x2) = mymiles
								x2 = x2 + 1
							Else
								tmpmileADP(lngIdx) = tmpmileADP(lngIdx) + mymiles
							End If
						End If
					End If
					If Z_fixnull(rsTBL("ProcMile")) = "" Then
						rsTBL("ProcMile") = Date
						rsTBL.Update
					End If
					rsTBL.MoveNext
				Loop
				rsTBL.Close
				Set rsTBL = Nothing
				'not w/n pay period
				If Request("fdate") <> "" Then
					Set rsTBL = CreateObject("ADODB.RecordSet")
					sqlTBL = "SELECT * FROM Tsheets_T, consumer_T, worker_T WHERE procMile is null AND Medicaid_Number = client AND emp_ID = Worker_T.Social_Security_Number"
					sqlTBL = sqlTBL & " AND date <= '" & CDate(Request("fdate")) & "'" 
					sqlTBL = sqlTBL & " ORDER BY consumer_T.lname, consumer_T.fname"
					
					rsTBL.Open sqlTBL, g_strCONN, 1, 3
					'x2 = 0
					Do Until rsTBL.EOF
						totmile = Z_CZero(rsTBL("mile")) + Z_CZero(rsTBL("amile"))
						If totmile <> 0 Then
							If Request("myUri") = 0 Then
								myempid = rsTBL("emp_ID")
								mycliid = rsTBL("client")
								mydate = rsTBL("date")
								mymiles = totmile
								mymilecap = rsTBL("consumer_T.milecap")
								lngIdx = SearchArrays4(myempid)
								If lngIdx < 0 Then
									ReDim Preserve tmpempADP(x2)
									ReDim Preserve tmpmileADP(x2)
		
									tmpempADP(x2) = myempid
									tmpmileADP(x2) = mymiles
									x2 = x2 + 1
								Else
									tmpmileADP(lngIdx) = tmpmileADP(lngIdx) + mymiles
								End If
							End If
						End If
						If Z_fixnull(rsTBL("ProcMile")) = "" Then
							rsTBL("ProcMile") = Date
							rsTBL.Update
						End If
						rsTBL.MoveNext
					Loop
					rsTBL.Close
					Set rsTBL = Nothing
				End If	
				''''
			End If
			If Request("myUri") = 0 Then
				'Set fso = CreateObject("Scripting.FileSystemObject")
				Set ALog = fso.OpenTextFile(AdminLog, 8, True)
				Alog.WriteLine Now & vbtab & "prfjdepi was ran by :: " & session("UserID")
				Set Alog = Nothing
				'Set fso = Nothing 
				y = 0
				Do Until y = x2 
					
						mileAmt = 0
						mileAmt = tmpmileADP(y) * 0.42 'change to db rate
						strProcB2 = strProcB2 & "FJD,50," & GetFileNum(tmpempADP(y)) & ",,,,,,,,,,,EXP," & mileAmt & ",,,," & vbCrlf
					
					y = y + 1
				Loop 
			End If
			If Request("myUri") = 0 Then
				If x2 > 0 Then
					a = 1
					tmpdate = replace(date, "/", "")
					prfjdepiCSV = "C:\work\lss-dbvortex\export\prfjdepi" & tmpdate & ".csv"
					if fso.FileExists(prfjdepiCSV) THen
						Do
							if fso.FileExists("C:\work\lss-dbvortex\export\prfjdepi" & tmpdate & "-" & a & ".csv") Then
								a = a + 1
							Else
								prfjdepiCSV = "C:\work\lss-dbvortex\export\prfjdepi" & tmpdate & "-" & a & ".csv"
								Exit Do
							End If
						Loop		 
					End If
					
					Set Prt = fso.CreateTextFile(prfjdepiCSV, True)
					Prt.WriteLine strProcH2
					Prt.WriteLine strProcB2
					'Session("dload2") = Z_DoEncrypt(prfjdepiCSV)
					
					copypath = copyfile & Z_GetFilename(prfjdepiCSV)
					Set Prt2 = fso.CreateTextFile(copypath, True)
					Prt2.WriteLine strProcH2
					Prt2.WriteLine strProcB2
					Session("dload2") = Z_DoEncrypt(Z_GetFilename(prfjdepiCSV))
				Else
					Session("dload2") = Z_DoEncrypt("NONE")
				End If
			End If
		End If
			Set Prt = fso.CreateTextFile(MileageCSV, True)
			Prt.WriteLine strProcH
			Prt.WriteLine strProcB
			If strProcH3 <> "" Then
				Prt.WriteLine vbcrlf
				Prt.WriteLine strProcH3
				Prt.WriteLine strProcB3 
			End If
			copypath = copyfile & Z_GetFilename(MileageCSV)
			Set Prt3 = fso.CreateTextFile(copypath, True)
			Prt3.WriteLine strProcH
			Prt3.WriteLine strProcB
			If strProcH3 <> "" Then
				Prt3.WriteLine vbcrlf
				Prt3.WriteLine strProcH3
				Prt3.WriteLine strProcB3 
			End If
			Session("dload") = Z_DoEncrypt(Z_GetFilename(MileageCSV))
			
			'Session("dload") = Z_DoEncrypt(MileageCSV)
		If Request("myUri") = 0 Then
			'If x2 > 0 Then
				Response.Redirect "extra2.asp?xxx=1&zzz=1&yyy=" & Request("prj")	
			'Else
				'Response.Redirect "extra2.asp?xxx=1&zzz=0"	
			'End If
		Else
			Response.Redirect "extra2.asp?xxx=1&zzz=0"	
		End If
		'response.write strProcB2
	ElseIf Request("sql") = 12 Then
		Set fso = CreateObject("Scripting.FileSystemObject")

			strProcH = "Worker Last Name, Worker Last Name, Address, City, State, Zip, Phone, RHICC1"
	

		Set rsTBL = CreateObject("ADODB.RecordSet")
		sqlTBL = "SELECT * FROM worker_T WHERE status = 'Active' AND privatepay = 1 ORDER BY lname, fname"
		
		rsTBL.Open sqlTBL, g_strCONN, 1, 3
		Do Until rsTBL.EOF
			
			tmpRCC = GetName3(z_czero(rsTBL("pm1")))
			
			strProcB = strProcB & """" &  rsTBL("lname") & """,""" & rsTBL("fname") & """,""" & rsTBL("maddress") & _
						""",""" & rsTBL("mcity") & """,""" & rsTBL("mstate") & """,""" & rsTBL("mzip") & """,""" & rsTBL("phoneno") & """,""" & tmpRCC & """" & vbCrlf
			rsTBL.MoveNext
		Loop
		rsTBL.Close
		Set rsTBL = Nothing
		
		Set Prt = fso.CreateTextFile(PPCSV, True)
		Prt.WriteLine strProcH
		Prt.WriteLine strProcB
		copypath = copyfile & Z_GetFilename(PPCSV)
		Set Prt2 = fso.CreateTextFile(copypath, True)
		Prt2.WriteLine strProcH
		Prt2.WriteLine strProcB
		Session("dload") = Z_DoEncrypt(Z_GetFilename(PPCSV))
		'Session("dload") = Z_DoEncrypt(PPCSV)

		Response.Redirect "extra2.asp"	
		'response.write strProcB2
	ElseIf Request("sql") = 13 Then
	strTitle = "Total Hours of "
				sqlProc = "SELECT * FROM [Tsheets_t] "
				If Request("seltype32") = 1 Then
					strTitle = strTitle & "PCSP Workers "
					sqlProc = sqlProc & ", [Worker_T] WHERE emp_id = Social_Security_Number"
				Else
					strTitle = strTitle & "Consumers "
					sqlProc = sqlProc & ", [Consumer_T] WHERE client = Medicaid_Number"
				End If
				Err = 0
				If Request("FrmD82") <> "" Then
					If IsDate(Request("FrmD82")) Then
						sqlProc = sqlProc & " AND date >= '" & dateAdd("d", -7, Request("FrmD82")) & "' "
					strTitle = strTitle & " from " & Request("FrmD82")
					Else
						Err = 1
					End If
				End If
				If Request("ToD82") <> "" Then
					If IsDate(Request("ToD82")) Then
						sqlProc = sqlProc & " AND date  <= '" & dateAdd("d", 7, Request("ToD82")) & "'" 
						strTitle = strTitle & " to " & Request("ToD82")
					Else
						Err = 1
					End If
				End If
				If Err <> 0 Then Response.Redirect "specrep.asp?err=60" 
				If Request("seltype32") = 1 Then
					strProcH = "File Number,PCSP Worker,Regular Hours,Extended Hours,Total Hours"
					sqlProc = sqlProc & " ORDER BY lname, fname"
				Else
					strProcH = "Consumer Name,Code,Regular Hours,Extended Hours,Total Hours"
					sqlProc = sqlProc & " ORDER BY code, lname, fname"
				End If	
					
				Set rsProc = Server.CreateObject("ADODB.RecordSet")	
				rsProc.Open sqlProc, g_strCONN, 3, 1
				x = 0
				Do Until rsProc.EOF
					If Request("seltype32") = 1 Then
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
						tmphrsMon = ValidDate(Request("FrmD82"), Request("ToD82"), rsProc("date"), rsProc("mon"), "MON")
            tmphrsTue = ValidDate(Request("FrmD82"), Request("ToD82"), rsProc("date"), rsProc("tue"), "TUE")
            tmphrsWed = ValidDate(Request("FrmD82"), Request("ToD82"), rsProc("date"), rsProc("wed"), "WED")
            tmphrsThu = ValidDate(Request("FrmD82"), Request("ToD82"), rsProc("date"), rsProc("thu"), "THU")
            tmphrsFri = ValidDate(Request("FrmD82"), Request("ToD82"), rsProc("date"), rsProc("fri"), "FRI")
            tmphrsSat = ValidDate(Request("FrmD82"), Request("ToD82"), rsProc("date"), rsProc("sat"), "SAT")
            tmphrsSun = ValidDate(Request("FrmD82"), Request("ToD82"), rsProc("date"), rsProc("sun"), "SUN")
          End If
					dblHours = tmphrsMon + tmphrsTue + tmphrsWed + tmphrsThu + tmphrsFri + tmphrsSat + tmphrsSun
					
					lngIdx = SearchArrays60(strID, tmpID3)
					If lngIdx < 0 Then
						ReDim Preserve tmpID3(x)
						ReDim Preserve tmpHrs(x)
						ReDim Preserve tmpHrsExt(x)
						
						tmpID3(x) = strID
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
				v = 0
				ahrs = 0
				chrs = 0
				mhrs = 0
				phrs = 0
				vhrs = 0
				Do Until y = x 
					myHrs = Z_CZero(tmpHrs(y) + tmpHrsExt(y))
					If myHrs > 0 Then
						If Request("seltype32") = 1 Then
							tmpName = GetName(tmpID3(y))
							strProcB = strProcB & """" & GetFileNum(tmpID3(y)) & """,""" & tmpName
						Else
							tmpName = GetName2(tmpID3(y))
							strProcB = strProcB & """" & tmpName & """,""" & GetCode(tmpID3(y))
							
							If GetCode(tmpID3(y)) = "A" Then
								a = a + 1
								ahrs = ahrs + myHrs
							ElseIf GetCode(tmpID3(y)) = "C" Then
								c = c + 1 
								chrs = chrs + myHrs
							ElseIf GetCode(tmpID3(y)) = "M" Then
								m = m + 1
								mhrs = mhrs + myHrs
							ElseIf GetCode(tmpID3(y)) = "P" Then
								p = p + 1
								phrs = phrs + myHrs
							ElseIf GetCode(tmpID3(y)) = "V" Then
								v =  + 1
								vhrs = vhrs + myHrs
							End If 	 		
						End If
						strProcB = strProcB &	""",""" & Z_CZero(tmpHrs(y)) & """,""" & Z_CZero(tmpHrsExt(y)) & """,""" & Z_CZero(tmpHrs(y) + tmpHrsExt(y)) & """" & vbCrLf
						End If
					y = y + 1
				Loop 
				ctotal = 0
				hrstotal = 0
				If Request("seltype32") <> 1 Then
					ctotal = a + c + m + p + v
					hrstotal = ahrs + chrs + mhrs + phrs + vhrs
					strProcB = strProcB &	"""TOTALS""" & vbCrlf & """CODE"",""COUNT"",""HOURS""" & vbCrLf & _
							"""A"",""" & a & """,""" & ahrs & """" & vbCrLf & _
							"""C"",""" & c & """,""" & chrs & """" & vbCrLf & _
							"""M"",""" & m & """,""" & mhrs & """" & vbCrLf & _
							"""P"",""" & p & """,""" & phrs & """" & vbCrLf & _
							"""V"",""" & v & """,""" & vhrs & """" & vbCrLf & _
							"""Total"",""" & ctotal & """,""" & hrstotal & """"
				End If
			Set fso = CreateObject("Scripting.FileSystemObject")
				Set Prt = fso.CreateTextFile(THrsCSV, True)
				Prt.WriteLine strTitle
			Prt.WriteLine strProcH
			Prt.WriteLine strProcB
			copypath = copyfile & Z_GetFilename(THrsCSV)
		Set Prt2 = fso.CreateTextFile(copypath, True)
		Prt2.WriteLine strProcH
		Prt2.WriteLine strProcB
		Session("dload") = Z_DoEncrypt(Z_GetFilename(THrsCSV))
			'Session("dload") = Z_DoEncrypt(THrsCSV)
			Response.Redirect "extra2.asp"
	ElseIf Request("sql") = 14 Then
		
	ElseIf Request("sql") = 15 Then
			PDate = Date
				markerX = 0
				If Request("ToD82") <> "" Then 
					If IsDate(Request("ToD82")) Then
						Pdate = Request("ToD82")
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
		If Request("seltype42") = 2 Then
			sqlProc = sqlProc & ", consumer_t  WHERE client = medicaid_number AND code = 'M' "
			Session("MSG") = Session("MSG") & "medicaid (simulation)."
		ElseIf Request("seltype42") = 3 Then
			sqlProc = sqlProc & ", consumer_t  WHERE client = medicaid_number AND (code = 'P' OR code = 'C' OR code = 'A') "
			Session("MSG") = Session("MSG") & "private pay (simulation)."
		ElseIf Request("seltype42") = 4 Then
			sqlProc = sqlProc & ", consumer_t  WHERE client = medicaid_number AND code = 'V' "
			Session("MSG") = Session("MSG") & "veterans (simulation)."
		End If
		sqlProc = sqlProc & "AND date <= '" & satDATE & "' AND date >= '" & sunDATE & "' AND" 
		mySunDate = sunDATE
		If Request("seltype42") = 2 Then
			sqlProc = sqlProc & " ProcMed is null AND EXT = 0 ORDER BY lname, fname ASC, date DESC"
		ElseIf Request("seltype42") = 3 Then
			sqlProc = sqlProc & " ProcPriv is null AND EXT = 0 ORDER BY lname, fname ASC, date DESC"
		ElseIf Request("seltype42") = 4 Then
			sqlProc = sqlProc & " ProcVA is null AND EXT = 0 ORDER BY lname, fname ASC, date DESC"
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
			'	strProcH = "<tr bgcolor='#040C8B'><td align='center' width='100px'><font size='1' face='trebuchet ms' color='white' color='white'>Timesheet Week</font>" & _
			'		"</td><td align='center' width='100px'><font size='1' face='trebuchet ms' color='white' color='white'>Medicaid</font></td><td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>" & _
			'		"Consumer</font></td><td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>" & _
			'		"PCSP</font></td><td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>Regular Hours</font></td><td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>" & _
			'		"Holiday Hours</font></td><td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>" & _
			'		"Extended Hours</font></td><td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>" & _
			'		"Rate</font></td><td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>" & _
			'		"Mileage</font></td><td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>" & _
			'		"Notes</font></td><td align='center'><font size='1' face='trebuchet ms' color='white' color='white'>User ID</font></td></tr>"
			strProcH = "Timesheet Week, Medicaid, Consumer Last Name, Consumer First Name, PCSP Last Name, PCSP First Name, Regular Hours, Holiday Hours, Extended Hours, Rate, Mileage,Notes, User ID"
			'End If
			Do Until rsProc.EOF
				myNotes = ""
				'Get EXT
				Set rsEXT = Server.CreateObject("ADODB.RecordSet")
				sqlEXT = "SELECT * FROM [Tsheets_t] WHERE date = '" & rsProc("date") & "' AND Client = '" & rsProc("client") & "' AND emp_id = '" & rsProc("emp_ID") & "' " & _
				 	"AND EXT = 1 AND TimeStamp = '" & rsProc("TimeStamp") & "'"
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
				If Request("seltype42") = 2 Then
					regHrs = rsProc("mon") + rsProc("tue") + rsProc("wed") + rsProc("thu") + rsProc("fri") + rsProc("sat") + rsProc("sun")
					holHrs = 0
				ElseIf Request("seltype42") = 3 Then
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
				ElseIf Request("seltype42") = 4 Then
					regHrs = rsProc("mon") + rsProc("tue") + rsProc("wed") + rsProc("thu") + rsProc("fri") + rsProc("sat") + rsProc("sun")
					holHrs = 0
				End If 
				myMile = rsProc("mile") + rsProc("amile")
				myNotes = rsProc("misc_notes") & " / " & myNotes
				If Session("lngType") = 1 Or Session("lngType") = 2 Then 
					strProcB = strProcB & """" & strDate & """,""" & rsProc("client") & """,""" & rsProc("lname") & """,""" & rsProc("fname") & """,""" &  GetNameWork(rsProc("emp_id")) & _
						""",""" & Z_FormatNumber(regHrs, 2) & """,""" & Z_FormatNumber(holHrs, 2) & """,""" & Z_FormatNumber(extHrs, 2) & _
						""",""" & Z_FormatNumber(GetPRate2(rsProc("client")), 2) & """,""" & Z_FormatNumber(myMile, 2) & """,""" & myNotes & """,""" & GetUser(rsProc("author")) & """" & vbCrLf
				Else
						strProcB = strProcB & """" & strDate & """,""" & "" & """,""" & rsProc("lname") & """,""" & rsProc("fname") & """,""" &  GetNameWork(rsProc("emp_id")) & _
						""",""" & Z_FormatNumber(regHrs, 2) & """,""" & Z_FormatNumber(holHrs, 2) & """,""" & Z_FormatNumber(extHrs, 2) & _
						""",""" & Z_FormatNumber(GetPRate2(rsProc("client")), 2) & """,""" & Z_FormatNumber(myMile, 2) & """,""" & myNotes & """,""" & GetUser(rsProc("author")) & """" & vbCrLf
				End If
				rsProc.MoveNExt
			Loop
		Else
			'NO RECORDS FOUND
			If Request("seltype42") = 2 Then
					strProcB = "No medicaid records found on " & sunDATE & " - " & satDATE & " for processing."
			ElseIf Request("seltype42") = 3 Then
					strProcB = "No Private Pay records found on " & sunDATE & " - " & satDATE & " for processing."
			ElseIf Request("seltype42") = 4 Then
					strProcB = "No Veterans records found on " & sunDATE & " - " & satDATE & " for processing."
			End If
		End If
	rsProc.CLose
	Set rsProc = Nothing	
	'NOT within 2 week period
	Set rsProc2 = Server.CreateObject("ADODB.RecordSet")
	sqlProc2 = "SELECT * FROM tsheets_t"
	If Request("seltype42") = 2 Then
		sqlProc2 = sqlProc2 & ", consumer_t  WHERE code = 'M' AND date < '" & sunDATE & "' AND client = medicaid_number AND ProcMed is null AND EXT = 0 ORDER BY lname, fname, Date ASC, client ASC"
	ElseIf Request("seltype42") = 3 Then
		sqlProc2 = sqlProc2 & ", consumer_t  WHERE (code = 'P' OR code = 'C' OR code = 'A') AND date < '" & sunDATE & "' AND client = medicaid_number AND ProcPriv is null AND EXT = 0 ORDER BY lname, fname, Date ASC, client ASC"
	ElseIf Request("seltype42") = 4 Then
		sqlProc2 = sqlProc2 & ", consumer_t  WHERE code = 'V' AND date < '" & sunDATE & "' AND client = medicaid_number AND ProcVA is null AND EXT = 0 ORDER BY lname, fname, Date ASC, client ASC"
	End If
	MarkerX = 0
	rsProc2.Open sqlProc2, g_strCONN, 1, 3	
	If Not rsProc2.EOF THen
		strProcH2 = "Processed items before the set payroll period" & vbCrLf
		strProcH2 = "Timesheet Week, Medicaid, Consumer Last Name, Consumer First Name, PCSP Last Name, PCSP First Name, Regular Hours, Holiday Hours, Extended Hours, Rate, Mileage,Notes, User ID"
		Do Until rsProc2.EOF
			myNotes = ""
			'Get EXT
			Set rsEXT = Server.CreateObject("ADODB.RecordSet")
			sqlEXT = "SELECT * FROM [Tsheets_t] WHERE date = '" & rsProc2("date") & "' AND Client = '" & rsProc2("client") & "' AND emp_id = '" & rsProc2("emp_ID") & "' " & _
			 	"AND EXT = 1 AND TimeStamp = '" & rsProc2("TimeStamp") & "'"
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
			ElseIf Request("seltype4") = 4 Then
				regHrs = rsProc2("mon") + rsProc2("tue") + rsProc2("wed") + rsProc2("thu") + rsProc2("fri") + rsProc2("sat") + rsProc2("sun")
				holHrs = 0
			End If 
			myMile = rsProc2("mile") + rsProc2("amile")
			myNotes = rsProc2("misc_notes") & " / " & myNotes
			If Session("lngType") = 1 Or Session("lngType") = 2 Then 
				strProcB2 = strProcB2 & """" & strDate & """,""" & rsProc2("client") & """,""" & rsProc2("lname") & """,""" & rsProc2("fname") & """,""" &  GetNameWork(rsProc2("emp_id")) & _
					""",""" & Z_FormatNumber(regHrs, 2) & """,""" & Z_FormatNumber(holHrs, 2) & """,""" & Z_FormatNumber(extHrs, 2) & _
					""",""" & Z_FormatNumber(GetPRate2(rsProc2("client")), 2) & """,""" & Z_FormatNumber(myMile, 2) & """,""" & myNotes & """,""" & GetUser(rsProc2("author")) & """" & vbCrLf
			Else
				strProcB2 = strProcB2 & """" & strDate & """,""" & "" & """,""" & rsProc2("lname") & """,""" & rsProc2("fname") & """,""" &  GetNameWork(rsProc2("emp_id")) & _
					""",""" & Z_FormatNumber(regHrs, 2) & """,""" & Z_FormatNumber(holHrs, 2) & """,""" & Z_FormatNumber(extHrs, 2) & _
					""",""" & Z_FormatNumber(GetPRate2(rsProc2("client")), 2) & """,""" & Z_FormatNumber(myMile, 2) & """,""" & myNotes & """,""" & GetUser(rsProc2("author")) & """" & vbCrLf
			End If
			rsProc2.MoveNext
		Loop
	End If
	rsProc2.Close
	Set rsProc2 = Nothing	
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set Prt = fso.CreateTextFile(SimProcCSV, True)
	Prt.WriteLine strProcH
	Prt.WriteLine strProcB
	Prt.WriteLine
	Prt.WriteLine strProcH2
	Prt.WriteLine strProcB2
	copypath = copyfile & Z_GetFilename(SimProcCSV)
		Set Prt2 = fso.CreateTextFile(copypath, True)
		Prt2.WriteLine strProcH
		Prt2.WriteLine strProcB
		Session("dload") = Z_DoEncrypt(Z_GetFilename(SimProcCSV))
	'Session("dload") = Z_DoEncrypt(SimProcCSV)
	Response.Redirect "extra2.asp"
	ElseIf Request("sql") = 16 Then
		'Session("MSG") = "News letter report."
		strProcH = "Last Name, First Name, Address, City, State, Zip"
		Set rsTBL = Server.CreateObject("ADODB.RecordSet")
		sqlTBL = "SELECT * FROM consumer_t, c_status_T where C_Status_t.Medicaid_number = Consumer_t.Medicaid_number AND Active = 1 and code = 'M' " & _
			"ORDER BY mZip, mAddress"
		rsTBL.Open sqlTBL, g_strCONN, 3, 1
		Do Until rsTBL.EOF
			If rsTBL("mAddress") <> "" Then
				strProcB = strProcB & """" & rsTBL("lname") & """,""" & rsTBL("fname") & """,""" & rsTBL("mAddress") & """,""" & _
					rsTBL("mCity") & """,""" & rsTBL("mState") & """,""" & rsTBL("mZip") & """" & vbCrLf
			End If
			rsTBL.MoveNext
		Loop
		rsTBL.Close
		Set rsTBL = Nothing
		Set rsTBL = Server.CreateObject("ADODB.RecordSet")
		sqlTBL = "SELECT * FROM Representative_t, conrep_T, C_status_T WHERE RID = Representative_t.[index] " & _
			"AND CID = C_status_T.Medicaid_Number AND Active = 1 " & _
			"ORDER BY zip, address"
		rsTBL.Open sqlTBL, g_strCONN, 3, 1
		Do Until rsTBL.EOF
			If rsTBL("Address") <> "" Then
				strProcB = strProcB & """" & rsTBL("lname") & """,""" & rsTBL("fname") & """,""" & rsTBL("Address") & """,""" & _
					rsTBL("City") & """,""" & rsTBL("State") & """,""" & rsTBL("Zip") & """" & vbCrLf
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
				strProcB = strProcB & """" & rsTBL("lname") & """,""" & rsTBL("fname") & """,""" & rsTBL("mAddress") & """,""" & _
					rsTBL("mCity") & """,""" & rsTBL("mState") & """,""" & rsTBL("mZip") & """" & vbCrLf
			End If
			rsTBL.MoveNext
		Loop
		rsTBL.Close
		Set rsTBL = Nothing
		
		Set fso = CreateObject("Scripting.FileSystemObject")
		Set Prt = fso.CreateTextFile(NewsCSV, True)
		Prt.WriteLine strProcH
		Prt.WriteLine strProcB
	copypath = copyfile & Z_GetFilename(NewsCSV)
		Set Prt2 = fso.CreateTextFile(copypath, True)
		Prt2.WriteLine strProcH
		Prt2.WriteLine strProcB
		Session("dload") = Z_DoEncrypt(Z_GetFilename(NewsCSV))
		'Session("dload") = Z_DoEncrypt(NewsCSV)
		Response.Redirect "extra2.asp"
	ElseIf Request("sql") = 17 Then
		strProcH = """Company"",""ClientID"",""Division"",""Customer Number"",""Program Code"",""Last Name"","" First Name"",""Program Description"","" Mailing Address"","" City"","" State"","" Zip"",""Sex"","" DOB"",""Client Identifier"",""Diag1"",""Diag2"",""Diag3"",""Diag4"",""SSN"""
		Set rsTBL = Server.CreateObject("ADODB.RecordSet")
		sqlTBL = "SELECT consumer_T.[Medicaid_Number] AS mednum, lname, fname, DOB, gender, maddress, mcity, mstate, mzip  FROM consumer_T, C_Status_t WHERE " & _
				"C_Status_t.Medicaid_number = Consumer_t.Medicaid_number AND Code = 'M' And Active = 1 " & _
				"ORDER BY lname, fname"
		rsTBL.Open sqlTBL, g_strCONN, 3, 1
		Do Until rsTBL.EOF
			fixzip = rsTBL("mzip")
			If right(rsTBL("mzip"), 1) = "-" Then 
				strlen = len(rsTBL("mzip"))
				fixzip = left(rsTBL("mzip"), strlen - 1)
			End if
			'fixzip = right(rsTBL("mzip"), 5)
			'If right(rsTBL("mzip"), 1) = "-" Then fixzip = left(rsTBL("mzip"), 5)
			strProcB = strProcB & """LSS"",""" & rsTBL("mednum") & ""","""","""",""IHC"",""" & rsTBL("lname") & """,""" & rsTBL("fname") & ""","""",""" & _
				rsTBL("maddress") & """,""" & rsTBL("mcity") & """,""" & rsTBL("mstate") & """,""" & fixzip & """,""" & rsTBL("gender") & """,""" & _
				rsTBL("DOB") & """,""" & rsTBL("mednum") & """,""3440""" & vbCrLf
			rsTBL.MoveNext
		Loop
		rsTBL.Close
		Set rsTBL = Nothing
		
		Set fso = CreateObject("Scripting.FileSystemObject")
		Set Prt = fso.CreateTextFile(MedCSV, True)
		Prt.WriteLine strProcH
		Prt.WriteLine strProcB
		
		copypath = copyfile & Z_GetFilename(medcsv)
		Set Prt2 = fso.CreateTextFile(copypath, True)
		Prt2.WriteLine strProcH
		Prt2.WriteLine strProcB
		Session("dload") = Z_DoEncrypt(Z_GetFilename(medcsv))

		Response.Redirect "extra2.asp"
	ElseIf Request("sql") = 18 Then 'cons with case
		Set fso = CreateObject("Scripting.FileSystemObject")
		Set rsTBL = Server.CreateObject("ADODB.RecordSet")
	sqlTBL = "SELECT Consumer_t.[lname] as clname, Consumer_t.[fname] as cfname, maddress, mcity, mstate, mzip, DOB, PhoneNo, Consumer_t.Medicaid_Number, CMID, PMID FROM Consumer_t, C_Status_t, CMCon_T, Case_Manager_t " & _
					"WHERE CID = Consumer_t.Medicaid_number AND C_Status_t.Medicaid_number = Consumer_t.Medicaid_number AND Active = 1 AND CMID = Case_Manager_t.[index] ORDER BY Consumer_t.lname, Consumer_t.fname"
		rsTBL.Open sqlTBL, g_strCONN, 1, 3
		strProcH = "Last Name, First Name, Mailing Address, City, State, Zip, Phone, DOB, RCC Last Name, RCC First Name, Agency, Case Manager Last Name, Case Manager First Name, Agency Address, City, State, Zip"
		Do Until rsTBL.EOF	
			strProcB = strProcB & """" & rsTBL("clname") & """,""" & rsTBL("cfname") & """,""" & rsTBL("mAddress") & """,""" & rsTBL("mCity") & """,""" & _
				rsTBL("mState") & """,""" & rsTBL("mZIP") & """,""" & rsTBL("PhoneNo") & """,""" & rsTBL("DOB") & """,""" & GetCM(rsTBL("PMID")) & ""","""& _
				GetCMAgency(GetCMAgencyID(rsTBL("CMID"))) & """,""" &  GetCMName(rsTBL("CMID")) & """,""" & GetCMAdr(GetCMAgencyID(rsTBL("CMID"))) & _
				"""" & vbcrlf
			rsTBL.MoveNext
		Loop
		Set Prt = fso.CreateTextFile(ActiveCon, True)
		Prt.WriteLine strProcH
		Prt.WriteLine strProcB
		rsTBL.Close
		Set rsTBL = Nothing
		copypath = copyfile & Z_GetFilename(ActiveConCase)
		Set Prt2 = fso.CreateTextFile(copypath, True)
		Prt2.WriteLine strProcH
		Prt2.WriteLine strProcB
		Session("dload") = Z_DoEncrypt(Z_GetFilename(ActiveConCase))
		'Session("dload") = Z_DoEncrypt(ActiveCon)
		Response.Redirect "extra2.asp"	
	ElseIf Request("sql") = 19 Then 'cuurent care
		Set fso = CreateObject("Scripting.FileSystemObject")
		strProcH = "Last Name,First Name,Current Care Plan,RIHCC Last Name, RIHCC First Name"
		Set rsTBL = CreateObject("ADODB.RecordSet")
		sqlTBL = "SELECT * " & _
					"FROM consumer_T, C_Status_T " & _
					"WHERE consumer_T.medicaid_number = C_Status_T.medicaid_number " & _
					"AND Active = 1 " 
		If Request("selrh") > 0 Then sqlTBL = sqlTBL & "AND PMID = " & Request("selrh") & " "
		sqlTBL = sqlTBL & "ORDER BY CarePlan,Lname, Fname"
		rsTBL.Open sqlTBL, g_strCONN, 1, 3
		Do Until rsTBL.EOF
			strProcB = strProcB & """" & rsTBL("lname") & """,""" & rsTBL("fname") & """,""" & rsTBL("CarePlan") & """,""" & GetCM(rsTBL("PMID")) & """" & vbCrLf
			rsTBL.MoveNext
		Loop
		rsTBL.Close
		Set rsTBL = Nothing
		Set Prt = fso.CreateTextFile(ConPMTown, True)
		Prt.WriteLine strProcH
		Prt.WriteLine strProcB
		copypath = copyfile & Z_GetFilename(ConPMTown)
		Set Prt2 = fso.CreateTextFile(copypath, True)
		Prt2.WriteLine strProcH
		Prt2.WriteLine strProcB
		Session("dload") = Z_DoEncrypt(Z_GetFilename(ConPMTown))
		'Session("dload") = Z_DoEncrypt(ConPMTown)
		Response.Redirect "extra2.asp"	
	ElseIf Request("sql") = 20 Then 'start and inactive dte
		Set fso = CreateObject("Scripting.FileSystemObject")
		strProcH = "Last Name,First Name,RIHCC Last Name, RIHCC First Name, Start Date, Inactive Date, Inactive Reason"
		Set rsTBL = CreateObject("ADODB.RecordSet")
		sqlTBL = "SELECT * " & _
					"FROM consumer_T, C_Status_T " & _
					"WHERE consumer_T.medicaid_number = C_Status_T.medicaid_number " & _
					"ORDER BY Inactive_Date, Lname, Fname"
		rsTBL.Open sqlTBL, g_strCONN, 1, 3
		Do Until rsTBL.EOF
			inacres = ""
			if rsTBL("Enter_Nursing_Home") Then inacres = "Enter Nursing Home or Other Setting"
			if rsTBL("Unable_Self_Direct") Then inacres = "Unable to Self-Direct"
			if rsTBL("Unable_Suitable_Worker") Then inacres = "Unable to Self-Direct"
			if rsTBL("death") Then inacres = "Death"
			if trim(rsTBL("A_Other")) <> "" Then inacres = trim(rsTBL("A_Other"))
			strProcB = strProcB & """" & rsTBL("lname") & """,""" & rsTBL("fname") & """,""" & GetCM(rsTBL("PMID")) & """,""" & rsTBL("start_date") & _
				""",""" & rsTBL("inactive_date") & """,""" & inacres & """" & vbCrLf
			rsTBL.MoveNext
		Loop
		rsTBL.Close
		Set rsTBL = Nothing
		Set Prt = fso.CreateTextFile(strtinactivedte, True)
		Prt.WriteLine strProcH
		Prt.WriteLine strProcB
		copypath = copyfile & Z_GetFilename(strtinactivedte)
		Set Prt2 = fso.CreateTextFile(copypath, True)
		Prt2.WriteLine strProcH
		Prt2.WriteLine strProcB
		Session("dload") = Z_DoEncrypt(Z_GetFilename(strtinactivedte))
		'Session("dload") = Z_DoEncrypt(ConPMTown)
		Response.Redirect "extra2.asp"	
	ElseIf Request("sql") = 21 Then 'con with pcsp
		Set fso = CreateObject("Scripting.FileSystemObject")
		strProcH = "Last Name,First Name,RIHCC Last Name, RIHCC First Name, Max Hours, PCSP Worker Last Name, PCSP Worker First Name"
		Set rsTBL = CreateObject("ADODB.RecordSet")
		sqlTBL = "SELECT * FROM Consumer_T , c_Status_T WHERE Consumer_T.Medicaid_Number = c_Status_T.Medicaid_Number AND Active = 1 ORDER BY lname, fname"
		rsTBL.Open sqlTBL, g_strCONN, 1, 3
		Do Until rsTBL.EOF
			strProcB = strProcB & """" & rsTBL("lname") & """,""" & rsTBL("fname") & """,""" & GetCM(rsTBL("PMID")) & """,""" & rsTBL("maxhrs") & _
				""",""" & "" & """" & vbCrLf
			'GET WORKERS
			Set rsWork = Server.CreateObject("ADODB.RecordSet")
			sqlwork = "SELECT * FROM Conwork_T WHERE CID = '" & rsTBL("medicaid_number") & "'"
			rsWork.Open sqlwork, g_strCONN, 3, 1
			Do Until rsWork.EOF
				strProcB = strProcB & """" & "" & """,""" & "" & """,""" & "" & """,""" & "" & """,""" & "" & """,""" & GetWorkName(rsWork("WID")) & """" & vbCrLf
				rsWork.MoveNext
			Loop
			rsWork.Close
			Set rsWork = Nothing	
			rsTBL.MoveNext
		Loop
		rsTBL.Close
		Set rsTBL = Nothing
		Set Prt = fso.CreateTextFile(conpcsp, True)
		Prt.WriteLine strProcH
		Prt.WriteLine strProcB
		copypath = copyfile & Z_GetFilename(conpcsp)
		Set Prt2 = fso.CreateTextFile(copypath, True)
		Prt2.WriteLine strProcH
		Prt2.WriteLine strProcB
		Session("dload") = Z_DoEncrypt(Z_GetFilename(conpcsp))
		'Session("dload") = Z_DoEncrypt(ConPMTown)
		Response.Redirect "extra2.asp"	
	ElseIf Request("sql") = 22 Then 'license exp
		Set fso = CreateObject("Scripting.FileSystemObject")
		strProcH = "Last Name,First Name,Mailing Address, City, State, Zip, DOB, License Num.,Expiration Date,RIHCC Last Name, RIHCC First Name"
		Set rsTBL = CreateObject("ADODB.RecordSet")
		sqlTBL = "SELECT * FROM worker_t WHERE status = 'Active' AND Driver = 1 ORDER BY LicenseExpDate, worker_t.lname, worker_t.fname"
		rsTBL.Open sqlTBL, g_strCONN, 1, 3
		Do Until rsTBL.EOF
			strProcB = strProcB & """" & rsTBL("lname") & """,""" & rsTBL("fname") & """,""" & rsTBL("mAddress") & """,""" & rsTBL("mcity") & _
				""",""" & rsTBL("mstate") & """,""" & rsTBL("mzip") & """,""" & rsTBL("DOB") & """,""" &  rsTBL("LicenseNo") & """,""" & _
				rsTBL("LicenseExpDate") & """,""" & GetCM(rsTBL("PM1")) & """" & vbCrLf
			rsTBL.MoveNext
		Loop
		rsTBL.Close
		Set rsTBL = Nothing
		Set Prt = fso.CreateTextFile(lcsExp, True)
		Prt.WriteLine strProcH
		Prt.WriteLine strProcB
		copypath = copyfile & Z_GetFilename(lcsExp)
		Set Prt2 = fso.CreateTextFile(copypath, True)
		Prt2.WriteLine strProcH
		Prt2.WriteLine strProcB
		Session("dload") = Z_DoEncrypt(Z_GetFilename(lcsExp))
		'Session("dload") = Z_DoEncrypt(ConPMTown)
		Response.Redirect "extra2.asp"	
	ElseIf Request("sql") = 23 Then 'insur exp
		Set fso = CreateObject("Scripting.FileSystemObject")
		strProcH = "Last Name,First Name,Mailing Address, City, State, Zip, DOB,Expiration Date,RIHCC Last Name, RIHCC First Name"
		Set rsTBL = CreateObject("ADODB.RecordSet")
		sqlTBL = "SELECT * FROM worker_t WHERE status = 'Active' AND Driver = 1 ORDER BY insuranceexpdate, worker_t.lname, worker_t.fname"
		rsTBL.Open sqlTBL, g_strCONN, 1, 3
		Do Until rsTBL.EOF
			strProcB = strProcB & """" & rsTBL("lname") & """,""" & rsTBL("fname") & """,""" & rsTBL("mAddress") & """,""" & rsTBL("mcity") & _
				""",""" & rsTBL("mstate") & """,""" & rsTBL("mzip") & """,""" & rsTBL("DOB") & """,""" & _
				rsTBL("Insuranceexpdate") & """,""" & GetCM(rsTBL("PM1")) & """" & vbCrLf
			rsTBL.MoveNext
		Loop
		rsTBL.Close
		Set rsTBL = Nothing
		Set Prt = fso.CreateTextFile(InsExp, True)
		Prt.WriteLine strProcH
		Prt.WriteLine strProcB
		copypath = copyfile & Z_GetFilename(InsExp)
		Set Prt2 = fso.CreateTextFile(copypath, True)
		Prt2.WriteLine strProcH
		Prt2.WriteLine strProcB
		Session("dload") = Z_DoEncrypt(Z_GetFilename(InsExp))
		'Session("dload") = Z_DoEncrypt(ConPMTown)
		Response.Redirect "extra2.asp"	
	ElseIf Request("sql") = 24 Then 'TBtest
		Set fso = CreateObject("Scripting.FileSystemObject")
		strProcH = "Last Name,First Name,Date of Hire, Orientation, Essentials Training, PP Training, TB Test 1, TB Test 2,LNA Active, LNA Inactive, Physical"
		Set rsTBL = CreateObject("ADODB.RecordSet")
		sqlTBL = "SELECT lname, fname, Date_Hired, tb, tb2, phy, orient, pptrain, lnaactive, lnainactive, W_Files_T.[essentials] as essent, essentialsdate FROM Worker_t, W_Files_T WHERE SSN = Social_Security_Number AND status = 'Active' " & _
					"ORDER BY lname, fname"
		rsTBL.Open sqlTBL, g_strCONN, 1, 3
		Do Until rsTBL.EOF
			tb1 = ""
			if rsTBL("tb") Then tb1 = "X"
			tb2 = ""
			if rsTBL("tb2") Then tb2 = "X"
			phy = ""
			if rsTBL("phy") Then phy = "X"
			orient = ""
			if rsTBL("orient") Then orient = "X"
			pptrain = ""
			if rsTBL("pptrain") Then pptrain = "X"
			lnaactive = ""
			if rsTBL("lnaactive") Then lnaactive = "X"
			lnainactive = ""
			if rstbl("lnainactive") Then lnainactive = "X"
			essentials = ""
			if rstbl("essent") Then essentials = "X"
			strProcB = strProcB & """" & rsTBL("lname") & """,""" & rsTBL("fname") & """,""" & rsTBL("Date_Hired") & """,""" & orient & _
				""",""" & essentials & """,""" & """,""" & pptrain & """,""" & tb1 & """,""" & tb2 & """,""" & _
				lnaactive & """,""" & lnainactive & """,""" & phy & """" & vbCrLf
			rsTBL.MoveNext
		Loop
		rsTBL.Close
		Set rsTBL = Nothing
		Set Prt = fso.CreateTextFile(tbtest, True)
		Prt.WriteLine strProcH
		Prt.WriteLine strProcB
		copypath = copyfile & Z_GetFilename(tbtest)
		Set Prt2 = fso.CreateTextFile(copypath, True)
		Prt2.WriteLine strProcH
		Prt2.WriteLine strProcB
		Session("dload") = Z_DoEncrypt(Z_GetFilename(tbtest))
		'Session("dload") = Z_DoEncrypt(ConPMTown)
		Response.Redirect "extra2.asp"		
	ElseIf Request("sql") = 25 Then 'private pay
		Set fso = CreateObject("Scripting.FileSystemObject")
		strProcH = "Last Name,First Name,ID,Address,City,State,Zip,Contract Hours"
		Set rsTBL = CreateObject("ADODB.RecordSet")
		sqlTBL = "SELECT * FROM Consumer_T, c_status_T WHERE (consumer_t.medicaid_number LIKE '%Private Pay%' OR (code = 'P' OR code = 'C' OR code = 'A')) AND consumer_t.medicaid_number = c_status_t.medicaid_number " & _
					"AND Active = 1 ORDER BY lname, fname"
		rsTBL.Open sqlTBL, g_strCONN, 1, 3
		Do Until rsTBL.EOF
			strProcB = strProcB & """" & rsTBL("lname") & """,""" & rsTBL("fname") & """,""" & rsTBL("medicaid_number") & """,""" & rsTBL("Address") & _
				""",""" & rsTBL("city") & """,""" & rsTBL("state") & """,""" & rsTBL("zip") & """,""" & rsTBL("contract") & """" & vbCrLf
				tmpcontract = tmpcontract + Z_Czero(rsTBL("contract"))
			rsTBL.MoveNext
		Loop
		strProcB = strProcB & """" & "" & """,""" & "" & """,""" & "" & """,""" & "" & """,""" & "" & """,""" & "" & """,""" & "" & "TOTAL:" & _
						""",""" & tmpcontract & """"
		rsTBL.Close
		Set rsTBL = Nothing
		Set Prt = fso.CreateTextFile(ppay, True)
		Prt.WriteLine strProcH
		Prt.WriteLine strProcB
		copypath = copyfile & Z_GetFilename(ppay)
		Set Prt2 = fso.CreateTextFile(copypath, True)
		Prt2.WriteLine strProcH
		Prt2.WriteLine strProcB
		Session("dload") = Z_DoEncrypt(Z_GetFilename(ppay))
		'Session("dload") = Z_DoEncrypt(ConPMTown)
		Response.Redirect "extra2.asp"	
	ElseIf Request("sql") = 26 Then 'site visit con
		Set fso = CreateObject("Scripting.FileSystemObject")
		strProcH = "Date,Last Name,First Name,Address,City,State,Zip,Phone No.,Project Manager Last Name, Project Manager First Name,Comment"
		Set rsTBL = CreateObject("ADODB.RecordSet")
		sqlTBL = "SELECT Consumer_t.Medicaid_number as cmednum, C_Site_Visit_Dates_t.[index] as cid, site_V_date, pmid FROM Consumer_t, C_Status_t, C_Site_Visit_Dates_t , Proj_Man_T WHERE " & _
					"PMID = Proj_Man_T.ID AND Consumer_t.Medicaid_number = C_Status_t.Medicaid_number AND Consumer_t.Medicaid_number =" & _
					" C_Site_Visit_Dates_t.Medicaid_number AND Active = 1 AND NOT Site_V_Date IS NULL " & _
					"ORDER BY Proj_Man_T.Lname, Proj_Man_T.Fname ASC, Consumer_t.Medicaid_number, Site_V_Date DESC"
		rsTBL.Open sqlTBL, g_strCONN, 1, 3
				
				tmpIDx = ""
				x = 0
				
				Do Until rsTBL.EOF
					If rsTBL("cmednum") <> tmpIDx then 
						ReDim Preserve tmp(x)
						tmp(x) = rsTBL("cid") & "|" & rsTBL("Site_V_Date") & "|" & rsTBL("PMID")
						x = x + 1
					End If
					tmpIDx = rsTBL("cmednum")
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
					" C_Site_Visit_Dates_t.Medicaid_number AND Active = 1 AND NOT Site_V_Date IS NULL " & _
					"AND C_Site_Visit_Dates_t.[index] = " & tmp2(0)
					
					rsTBL2.Open sqlTBL2, g_strCONN, 1, 3
					
						Set rsPM = Server.CreateObject("ADODB.RecordSet")
						sqlPM = "SELECT * FROM Proj_Man_T WHERE ID = " & rsTBL2("PMID")
						rsPM.Open sqlPM, g_strCONN, 1, 3
						If Not rsPM.EOF Then
							PMname = rsPM("lname") & """,""" & rsPM("fname")
						Else
							PMname = rsPM("ID")
						End If
						rsPM.Close
						Set rsPM = Nothing
						newComment = Replace(rsTBL2("Comments"), "|",  " ")
						strProcB = strProcB & """" & rsTBL2("Site_V_Date") & _
							""",""" & rsTBL2("lname") & """,""" & _
							rsTBL2("fname") & """,""" & _
							rsTBL2("mAddress") & """,""" & rsTBL2("mCity") & """,""" & rsTBL2("mState") & """,""" & rsTBL2("mzip") & """,""" & rsTBL2("PhoneNo") & _
							""",""" & PMname & """,""" & newComment & """" & vbCrLf
					rsTBL2.Close
					zzz =zzz + 1
				Loop
				Set rsTBL2 = Nothing
		Set Prt = fso.CreateTextFile(svcon, True)
		Prt.WriteLine strProcH
		Prt.WriteLine strProcB
		copypath = copyfile & Z_GetFilename(svcon)
		Set Prt2 = fso.CreateTextFile(copypath, True)
		Prt2.WriteLine strProcH
		Prt2.WriteLine strProcB
		Session("dload") = Z_DoEncrypt(Z_GetFilename(svcon))
		'Session("dload") = Z_DoEncrypt(ConPMTown)
		Response.Redirect "extra2.asp"	
	ElseIf Request("sql") = 27 Then 'site visit wor
		Set fso = CreateObject("Scripting.FileSystemObject")
		strProcH = "Date,Last Name,First Name,Address,City,State,Zip,Phone No.,RIHCC,RIHCC 2,Comment"
		Set rsTBL = CreateObject("ADODB.RecordSet")
		sqlTBL = "SELECT worker_t.social_security_number as wssn, w_log_t.[index] as wid, sitev, pm1, pm2 " & _ 
					"FROM Worker_t, w_log_t, consumer_t, conwork_t, proj_man_t  " & _
					"WHERE Worker_t.social_security_number = w_log_t.ssn " & _
					"AND PM1 = Proj_Man_T.ID " & _
					"AND consumer_t.medicaid_number = CID " & _
					"AND worker_t.[index] = WID " & _
					"AND status = 'Active' " & _
					"AND NOT sitev IS NULL " & _
					"ORDER BY Proj_Man_T.Lname, Proj_Man_T.Fname, worker_T.social_security_number ASC, sitev DESC"
		rsTBL.Open sqlTBL, g_strCONN, 1, 3
				
				tmpIDx = ""
				x = 0
				
				Do Until rsTBL.EOF
					If rsTBL("wssn") <> tmpIDx then 
						ReDim Preserve tmp(x)
						tmp(x) = rsTBL("wid") & "|" & rsTBL("sitev") & "|" & rsTBL("PM1") & "|" & rsTBL("PM2") 
						x = x + 1
					End If
					tmpIDx = rsTBL("wssn")
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
						"status = 'Active' AND NOT sitev IS NULL AND w_log_t.[index] = " & tmp2(0)
					
					rsTBL2.Open sqlTBL2, g_strCONN, 1, 3
						strProcB = strProcB & """" & rsTBL2("sitev") & """,""" & rsTBL2("lname") & """,""" & _
								rsTBL2("fname") & """,""" & _
								rsTBL2("mAddress") & """,""" & rsTBL2("mCity") & """,""" & rsTBL2("mState") & """,""" & rsTBL2("mzip") & """,""" & _
								rsTBL2("PhoneNo") & """,""" & GetName3(tmp2(2)) & """,""" & GetName3(tmp2(3)) & """,""" & rsTBL2("scom") & """" & vbCrLf
								'add PM
					rsTBL2.Close
					zzz =zzz + 1
				Loop
				Set rsTBL2 = Nothing
		Set Prt = fso.CreateTextFile(svwor, True)
		Prt.WriteLine strProcH
		Prt.WriteLine strProcB
		copypath = copyfile & Z_GetFilename(svwor)
		Set Prt2 = fso.CreateTextFile(copypath, True)
		Prt2.WriteLine strProcH
		Prt2.WriteLine strProcB
		Session("dload") = Z_DoEncrypt(Z_GetFilename(svwor))
		'Session("dload") = Z_DoEncrypt(ConPMTown)
		Response.Redirect "extra2.asp"	
	ElseIf Request("sql") = 30 Then
		'Dim myDay(6), sundatex(1), rsHrs(6)', SearchArraysacode(1)
		myDay(0) = "sun"
		myDay(1) = "mon"
		myDay(2) = "tue"
		myDay(3) = "wed"
		myDay(4) = "thu"
		myDay(5) = "fri"
		myDay(6) = "sat"
		sundatex(0) = Request("closedate2")
		sundatex(1) = DateAdd("d", 7, CDate(Request("closedate2")))
		Set fso = CreateObject("Scripting.FileSystemObject")
		strProcH = "Worker Last Name,Worker First Name,Consumer,Date,Hours,Activity"
		Set rsTBL = Server.CreateObject("ADODB.RecordSet")
		Set rsCli = Server.CreateObject("ADODB.RecordSet")
		sqlTBL = "SELECT distinct emp_id, lname, fname FROM Tsheets_T, worker_T  WHERE emp_id = Social_Security_Number " & _
				"AND date >= '" & CDate(Request("closedate2")) & "' AND date  <= '" & CDate(Request("todate2")) & "' ORDER BY lname, fname"
		Session("Msg") = Session("Msg") & " from " & Request("closedate2")
		Session("Msg") = Session("Msg") & " to " & Request("todate2")
		Session("Msg") = Session("Msg") & ". "
		rsTBL.Open sqlTBL, g_strCONN, 3, 1
		Do Until rsTBL.EOF
      		sqlcli = "SELECT DISTINCT client FROM TSheets_T WHERE emp_ID = '" & rsTBL("emp_ID") & "' " & _
      				"AND date >= '" & CDate(Request("closedate2")) & "' AND date  <= '" & CDate(Request("todate2")) & "'" 
 		   	rsCli.Open sqlcli, g_strCONN, 3, 1
      		Do Until rsCli.EOF
	     		ctrs = 0
      			Do Until ctrs = 2
	      			For lngI = 0 To 6
	      				sqlHrs = "SELECT emp_ID, " & myDay(lngI) & " AS [val], misc_notes, date FROM tsheets_T " & _
	      						"WHERE client = '" & rsCli("client") & "' " & _
	      						"AND emp_ID= '" & rsTBL("emp_ID") & "' AND date = '" & CDate(sundatex(ctrs)) & "' " & _
	      						"AND " & myDay(lngI) & " <> 0 ORDER BY timestamp"
	      				Set rsHrs(lngI) = Server.CreateObject("ADODB.RecordSet")
	      				rsHrs(lngI).Open sqlHrs, g_strCONN, 3, 1
	      				'Response.Write "<code>" & sqlTbl & vbCrLf & sqlCli & vbCrLF & sqlHrs & "<code>" & vbCrLF
	      			Next
	      			'x = 0
	      			For lngI = 0 To 6
		      			dayhrs = 0
		      			'y = 0
		      			x = 0
		      			Do Until rsHrs(lngI).EOF
		      				accode = Split(Trim(rsHrs(lngI)("misc_notes")), ",")
		      				If UBound(accode) > 1 Then
                				Exit Do
                			Else
                  				y = 0
                  				Do Until y = UBound(accode) + 1
	                  				If accode(y) <> "" Then
		                  				lngIdx = SearchArraysacode(accode(y))
		                    			If lngIdx < 0 Then
		                      				ReDim Preserve myacode(x)
		                        			myacode(x) = accode(y)
		                        			x = x + 1
		                      			End If
		                  			End if
	                  				y = y + 1
                				Loop
                  				dayhrs = dayhrs + rsHrs(lngi)("val")
                			End If
                			rsHrs(lngi).MoveNext
						Loop
						actcode = ""
						If x = 1 And dayhrs > 1.25 Then
							myDate = Z_GetDate(sundatex(ctrs), myDay(lngI))
							For ctr2 = 0 to Ubound(myacode) 
								actcode = actcode & ACdesc(myacode(ctr2))
							Next
							strProcB = strProcB & """" & GetNameWork(rsTBL("emp_id")) & """,""" & GetName2(rsCli("client")) & _
									""",""" & myDate & """,""" & dayhrs & """,""" & actcode  & """" & vbCrLf
						End If
						ReDim myacode(0)
					Next
						
					For lngI = 0 To 6
						rsHrs(lngI).Close
						set rsHrs(lngi) = Nothing
					Next
					ctrs = ctrs + 1
				Loop
      			rsCli.MoveNext
      		Loop
      		rsCli.Close
			rsTBL.MoveNext
		Loop
		rsTBL.Close
		Set rsCli = Nothing
		Set rsTBL = Nothing	
		Set Prt = fso.CreateTextFile(insuf, True)
		Prt.WriteLine strProcH
		Prt.WriteLine strProcB
		copypath = copyfile & Z_GetFilename(insuf)
		Set Prt2 = fso.CreateTextFile(copypath, True)
		Prt2.WriteLine strProcH
		Prt2.WriteLine strProcB
		Session("dload") = Z_DoEncrypt(Z_GetFilename(insuf))
		'response.write Z_GetFilename(insuf)
		'Session("dload") = Z_DoEncrypt(ConPMTown)
		Response.Redirect "extra2.asp"	
	ElseIf Request("sql") = 31 Then
		Set fso = CreateObject("Scripting.FileSystemObject")
		strProcH = "Last Name,First Name,Mailing Address,City,State,Zip,UtliPro Badge ID,Date Of Hire,Termination Date,Salary,RIHCC"
		Set rsTBL = Server.CreateObject("ADODB.RecordSet")
		sqlTBL = "SELECT worker_t.lname as wlname, worker_t.fname as wfname, maddress, mcity, mstate, mzip, Term_date, date_hired, salary, pm1, pm2, ubadge FROM worker_t,Proj_Man_T WHERE status = 'Active' AND pm1 = proj_man_T.id ORDER BY Proj_Man_T.Lname, Proj_Man_T.Fname, month(date_hired), day(date_hired), worker_T.lname, worker_T.fname"
		If Request("Sort") = 1 Then 
			sqlTBL = "SELECT worker_t.lname as wlname, worker_t.fname as wfname, maddress, mcity, mstate, mzip, Term_date, date_hired, salary, pm1, pm2, ubadge FROM worker_t,Proj_Man_T WHERE status = 'Active' AND pm1 = proj_man_T.id ORDER BY month(date_hired), day(date_hired), Proj_Man_T.Lname, Proj_Man_T.Fname, worker_T.lname, worker_T.fname"
		End If
		If Request("Sort") = 1 Then sqlTBL = sqlTBL & " DESC"
		rsTBL.Open sqlTBL, g_strCONN, 1, 3
		Do Until rsTBL.EOF
			strProcB = strProcB & """" & rsTBL("wlname") & """,""" & rsTBL("wfname") & """,""" & rsTBL("mAddress") & _
						""",""" & rsTBL("mCity") & """,""" & rsTBL("mState") & """,""" & rsTBL("mZip") & """,""" & rsTBL("ubadge") & _
						""",""" &	rsTBL("Date_Hired") & """,""" & rsTBL("Term_date") & """,""" & Z_FormatNumber(rsTBL("Salary"), 2) & _
						""",""" & GetName3(rsTBL("PM1")) & """,""" & GetName3(rsTBL("PM2")) & """" & vbCrLf
						'<td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("CID") & "</font></td></tr>"
			rsTBL.MoveNext
		Loop
		rsTBL.Close
		Set rsTBL = Nothing
		Set Prt = fso.CreateTextFile(dteofhire, True)
		Prt.WriteLine strProcH
		Prt.WriteLine strProcB
		copypath = copyfile & Z_GetFilename(dteofhire)
		Set Prt2 = fso.CreateTextFile(copypath, True)
		Prt2.WriteLine strProcH
		Prt2.WriteLine strProcB
		Session("dload") = Z_DoEncrypt(Z_GetFilename(dteofhire))
		'Session("dload") = Z_DoEncrypt(ConPMTown)
		Response.Redirect "extra2.asp"			
	ElseIf Request("sql") = 32 Then
		Set fso = CreateObject("Scripting.FileSystemObject")
		strProcH = "Last Name,First Name,Mailing Address,City,State,Zip,Phone,Mobile,Email,RIHCC1,UltiPro"
		Set rsTBL = Server.CreateObject("ADODB.RecordSet")
		sqlTBL = "SELECT maddress, mcity, mstate, mzip, lname, fname, phoneno, cellno, email, pm1, COALESCE(ubadge, '') AS ubadge FROM Worker_T WHERE status = 'Active' AND Driver = 1 ORDER BY lname, fname"
		rsTBL.Open sqlTBL, g_strCONN, 1, 3
		Do Until rsTBL.EOF
			strProcB = strProcB & """" & rsTBL("lname") & """,""" & rsTBL("fname") & """,""" & rsTBL("mAddress") & _
						""",""" & rsTBL("mCity") & """,""" & rsTBL("mState") & """,""" & rsTBL("mZip") & """,""" & rsTBL("PhoneNo") & _
						""",""" &	rsTBL("CellNo") & """,""" & rsTBL("eMail") & """,""" & GetName3(rsTBL("PM1")) & """,""" & rsTBL("ubadge") & """" & vbCrLf
			rsTBL.MoveNext
		Loop
		rsTBL.Close
		Set rsTBL = Nothing
		Set Prt = fso.CreateTextFile(wdriver, True)
		Prt.WriteLine strProcH
		Prt.WriteLine strProcB
		copypath = copyfile & Z_GetFilename(wdriver)
		Set Prt2 = fso.CreateTextFile(copypath, True)
		Prt2.WriteLine strProcH
		Prt2.WriteLine strProcB
		Session("dload") = Z_DoEncrypt(Z_GetFilename(wdriver))
		'Session("dload") = Z_DoEncrypt(ConPMTown)
		Response.Redirect "extra2.asp"			
	ElseIf Request("sql") = 33 Then
		Set fso = CreateObject("Scripting.FileSystemObject")
		strProcH = "Last Name,First Name,Phone,Mobile,Email,Method of Communication,Language"
		Set rsTBL = Server.CreateObject("ADODB.RecordSet")
		sqlTBL = "SELECT [lname], [fname], [phoneno], [cellno], [email], [language], [prefcom] " & _
				"FROM Worker_T AS w " & _
				"INNER JOIN language_T AS l ON w.[langid]=l.[index] " & _
				"WHERE status = 'Active' ORDER BY lname, fname"
		rsTBL.Open sqlTBL, g_strCONN, 1, 3
		Do Until rsTBL.EOF
			prefcom = ""
			If rsTBL("prefcom") = 1 Then prefcom = "Mail"
			If rsTBL("prefcom") = 2 Then prefcom = "Email"
			If rsTBL("prefcom") = 3 Then prefcom = "Phone"
			If rsTBL("prefcom") = 4 Then prefcom = "Text"
			strProcB = strProcB & """" & rsTBL("lname") & """,""" & rsTBL("fname") & """,""" & rsTBL("PhoneNo") & _
						""",""" &	rsTBL("CellNo") & """,""" & rsTBL("eMail") & """,""" & prefcom & """,""" & _
						rsTBL("language") & """" & vbCrLf
			rsTBL.MoveNext
		Loop
		rsTBL.Close
		Set rsTBL = Nothing
		Set Prt = fso.CreateTextFile(wcom, True)
		Prt.WriteLine strProcH
		Prt.WriteLine strProcB
		copypath = copyfile & Z_GetFilename(wcom)
		Set Prt2 = fso.CreateTextFile(copypath, True)
		Prt2.WriteLine strProcH
		Prt2.WriteLine strProcB
		Session("dload") = Z_DoEncrypt(Z_GetFilename(wcom))
		'Session("dload") = Z_DoEncrypt(ConPMTown)
		Response.Redirect "extra2.asp"			
	End If
	
%>