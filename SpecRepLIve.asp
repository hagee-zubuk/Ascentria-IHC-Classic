<%Language=VBScript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_security.asp" -->
<%
	DIM tmp(), xtmp()
	
	Function GetName(zzz)
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
		Set rsName = Server.CreateObject("ADODB.RecordSet")
		sqlName = "SELECT * FROM consumer_T WHERE medicaid_number = '" & zzz & "' "
		rsName.Open sqlName, g_strCONN, 3, 1
		If Not rsName.EOF Then
			GetName2 = rsName("Lname") & ", " & rsName("fname")
		End If
		rsName.Close
		Set rsName = Nothing
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
	Session("PrintPrev") = ""
	Session("PrintPrevRep") = ""
	Session("PrintPrevPRoc") = ""
		If Request("sunDATE") <> "" and Request("satDATE") <> "" and Request("chkr") = 1 then 
			sunDATE = Request("sunDATE")
			satDATE = Request("satDATE")
			If Request("SelRep") = 27  Then 
				if Request("seltype") = 1 then selPay = "SELECTED"
				if Request("seltype") = 2 then selMed = "SELECTED"
				sel27 = "SELECTED"
			ElseIf Request("SelRep") = 39 Then
				sel39 = "SELECTED"
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
		'GET CONSUMERS
		Set rsCon = Server.CreateObject("ADODB.RecordSet")
		sqlCon = "SELECT * FROM Consumer_T, C_Status_T WHERE consumer_T.Medicaid_Number = C_Status_T.Medicaid_Number AND Active = true ORDER BY Lname, Fname"
		rsCon.Open sqlCon, g_strCONN, 3, 1
		Do Until rsCon.EOF
			strCON = strCON & "<option value='" & rsCON("Consumer_t.Medicaid_Number") & "'>" & GetName2(rsCON("Consumer_t.Medicaid_Number")) & "</option>" & vbCrLf
			rsCon.MoveNExt
		Loop
		rsCon.Close
		Set rsCon = Nothing
		
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
				
				tmpID = ""
				x = 0
				
				Do Until rsTBL.EOF
					If rsTBL("Consumer_t.Medicaid_number") <> tmpID then 
						ReDim Preserve tmp(x)
						tmp(x) = rsTBL("C_Site_Visit_Dates_t.index") & "|" & rsTBL("Site_V_Date") & "|" & rsTBL("PMID")
						x = x + 1
					End If
					tmpID = rsTBL("Consumer_t.Medicaid_number")
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
				
				tmpID = ""
				x = 0
				
				Do Until rsTBL.EOF
					If rsTBL("Consumer_t.Medicaid_number") <> tmpID then 
						ReDim Preserve tmp(x)
						tmp(x) = rsTBL("C_Site_Visit_Dates_t.index") & "|" & rsTBL("phoneCon_last") & "|" & rsTBL("PMID")
						x = x + 1
					End If
					tmpID = rsTBL("Consumer_t.Medicaid_number")
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
					"</font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Project Manager</b></font></td></tr>"
					'"</font></td><td align='center'><font size='1' face='trebuchet ms'>Consumer ID" & _
					'"</font></td></tr>"
					
					'<a style='text-decoration: none;' href='JavaScript: SVSort();'><font size='1' face='trebuchet ms' color='blue'>Date Of Hire</font>
				Set rsTBL = Server.CreateObject("ADODB.RecordSet")
				sqlTBL = "SELECT * FROM worker_t, consumer_t, conwork_t, Proj_Man_T " & _
					"WHERE PMID = Proj_Man_T.ID AND consumer_t.medicaid_number = CID AND CStr(worker_t.index) = WID " & _
					"AND status = 'Active' ORDER BY Proj_Man_T.Lname, Proj_Man_T.Fname, month(date_hired), day(date_hired), worker_T.lname, worker_T.fname"
				If Request("Sort") = 1 Then 
					sqlTBL = "SELECT * FROM worker_t, consumer_t, conwork_t, Proj_Man_T " & _
						"WHERE PMID = Proj_Man_T.ID AND consumer_t.medicaid_number = CID AND CStr(worker_t.index) = WID " & _
						"AND status = 'Active' ORDER BY month(date_hired), day(date_hired), Proj_Man_T.Lname, Proj_Man_T.Fname, worker_T.lname, worker_T.fname"
				End If
				If Request("Sort") = 1 Then sqlTBL = sqlTBL & " DESC"
				rsTBL.Open sqlTBL, g_strCONN, 1, 3
				Do Until rsTBL.EOF
					strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("worker_t.lname") & _
						", " & rsTBL("worker_t.fname") & "</font></td><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("worker_t.Address") & _
						", " & rsTBL("worker_t.City") & ", " & rsTBL("worker_t.State") & ", " & rsTBL("worker_t.Zip") & "</font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms'>&nbsp;" & rsTBL("worker_t.PhoneNo") & _
						"</font></td><td align='center'><font size='1' face='trebuchet ms'>&nbsp;" & _
						rsTBL("Date_Hired") & "</td><td align='center'><font size='1' face='trebuchet ms'>&nbsp;" & Z_FormatNumber(rsTBL("Salary"), 2) & _
						"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("Proj_man_T.Lname") & ", " & _
						rsTBL("Proj_man_T.Fname") & "</font></td></tr>"
						'<td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("CID") & "</font></td></tr>"
					rsTBL.MoveNext
				Loop
				rsTBL.Close
				Set rsTBL = Nothing
			ElseIf Request("SelRep") = 6 Then '4
				Session("MSG") = "Consumer by Town report."
				strHEAD = "<tr bgcolor='#040C8B'><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Project Manager</b></font>" & _
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
					"</font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Address</b></font></td><td align='center'>" & _
					"<font size='1' face='trebuchet ms' color='white'><b>Phone No.</b></font></td><td align='center'><font size='1' " & _
					"face='trebuchet ms' color='white'><b>DOB</b></font></td></tr>"
				Set rsTBL = Server.CreateObject("ADODB.RecordSet")
				sqlTBL = "SELECT lname, fname, address, city, state, zip, DOB, PhoneNo  FROM Consumer_t, C_Status_t " & _
					"WHERE C_Status_t.Medicaid_number = Consumer_t.Medicaid_number AND Active = True ORDER BY lname, fname"
				rsTBL.Open sqlTBL, g_strCONN, 1, 3
				Do Until rsTBL.EOF
					strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("lname") & ", " & _
						rsTBL("fname") & "</font></td><td align='center'><font size='1' face='trebuchet ms'>" & _
						rsTBL("Address") & ", " & rsTBL("City") & ", " & rsTBL("State") & ", " & rsTBL("Zip") & "</font></td><td align='center'>" & _
						"<font size='1' face='trebuchet ms'>&nbsp;" & rsTBL("PhoneNo") & "<td align='center'><font size='1' " & _
						"face='trebuchet ms'>&nbsp;" & rsTBL("DOB") & "</font></td></tr>"
					rsTBL.MoveNext
				Loop
				rsTBL.Close
				Set rsTBL = Nothing
			ElseIf Request("SelRep") = 3 Then '7
				Session("MSG") = "All Active PCSP Worker report."
				strHEAD = "<tr bgcolor='#040C8B'></font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Name</b> " &_
					"</font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Address</b></font></td><td align='center'>" & _
					"<font size='1' face='trebuchet ms' color='white'><b>Phone No.</b></font></td><td align='center'><font size='1' " & _
					"face='trebuchet ms' color='white'><b>DOB</b></font></td></tr>"
				Set rsTBL = Server.CreateObject("ADODB.RecordSet")
				sqlTBL = "SELECT lname, fname, address, city, state, zip, DOB, PhoneNo  FROM Worker_t  WHERE status = 'Active' " & _
					"ORDER BY lname, fname"
				rsTBL.Open sqlTBL, g_strCONN, 1, 3
				Do Until rsTBL.EOF
					strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("lname") & ", " & _
						rsTBL("fname") & "</font></td><td align='center'><font size='1' face='trebuchet ms'>" & _
						rsTBL("Address") & ", " & rsTBL("City") & ", " & rsTBL("State") & ", " & rsTBL("Zip") & "</font></td><td align='center'>" & _
						"<font size='1' face='trebuchet ms'>&nbsp;" & rsTBL("PhoneNo") & "<td align='center'><font size='1' " & _
						"face='trebuchet ms'>&nbsp;" & rsTBL("DOB") & "</font></td></tr>"
					rsTBL.MoveNext
				Loop
				rsTBL.Close
				Set rsTBL = Nothing
			ElseIf Request("SelRep") = 12 Then '10
				Session("MSG") = "Consumer Referrals report."
				strHEAD = "<tr bgcolor='#040C8B'><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Referral Date</b>" & _
					"</font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Name</b></font></td></tr>"
				Set rsTBL = Server.CreateObject("ADODB.RecordSet")
				sqlTBL = "SELECT lname, fname, Referral_Date, C_Status_t.Active FROM Consumer_t, C_Status_t " & _
					"WHERE Consumer_t.Medicaid_number = C_Status_t.Medicaid_number AND Active = True ORDER BY Referral_Date DESC, lname, fname"
				rsTBL.Open sqlTBL, g_strCONN, 1, 3
				Do Until rsTBL.EOF
					strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>&nbsp;" & _
						rsTBL("Referral_Date") & "</td><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("lname") & _
						", " & rsTBL("fname") & "</font></td></tr>"
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
							"face='trebuchet ms'>" & rsCon("MaxHrs") & "</td>" 
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
						"</font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Name</b></font></td></tr>"
				Set rsWor = Server.CreateObject("ADODB.RecordSet")
			sqlWor = "SELECT * FROM Worker_t WHERE [status] = 'Active' ORDER BY Lname, fname"
				rsWor.Open sqlWor, g_strCONN, 1, 3
				Do Until rsWor.EOF
					Set rsLink = Server.CreateObject("ADODB.RecordSet")
					sqlLink = "SELECT * FROM ConWork_t WHERE WID = '" & rsWor("index") & "' "
					rsLink.Open sqlLink, g_strCONN, 1, 3
					If rsLink.EOF Then
						strBODY = strBODY & "<tr><td align='left'><font size='1' face='trebuchet ms'>&nbsp;" & _
							rsWor("Social_Security_Number") & "</td><td align='center'><font size='1' face='trebuchet ms'>" & _
							rsWor("lname") & ", " & rsWor("fname") & "</font></td></tr>"
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
				sqlIA = "SELECT COUNT(Medicaid_number) AS TIA FROM C_Status_t WHERE Active = False"
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
					Session("MSG") = "All Active Consumers Sorted by Project Manager And Town report."
					strHEAD = "<tr bgcolor='#040C8B'><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Consumer Name</b>" & _
						"</font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Address</b></font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Phone No.</b></font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Project Manager</b></font></td></tr>"
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
						"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Insurance Expiration Date</b></font></td></tr>"
				Set rsTBL = Server.CreateObject("ADODB.RecordSet")
				sqlTBL = "SELECT fname, lname, address, city, state, zip, status,insuranceexpdate FROM worker_t " & _
					"WHERE status = 'Active' AND Driver = True ORDER BY insuranceexpdate, lname, fname"
				rsTBL.Open sqlTBL, g_strCONN, 1, 3
				Do Until rsTBL.EOF	
					strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("lname") & ", " & _
						rsTBL("fname") & "</font></td><td align='center'><font size='1' face='trebuchet ms'>" & _
						rsTBL("Address") & ", " & rsTBL("City") & ", " & rsTBL("State") & ", " & rsTBL("Zip") & "</font></td><td align='center'>" & _
						"<font size='1' face='trebuchet ms'>&nbsp;" & rsTBL("Insuranceexpdate") & "</td></tr>"
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
				tmpID = ""
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
				Session("MSG") = "All Active PCSP Worker by Project Manager And Town report."
				strHEAD = "<tr bgcolor='#040C8B'><td align='center'><font size='1' face='trebuchet ms' color='white'><b>PCSP Worker Name</b>" & _
						"</font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Address</b></font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Phone No.</b></font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Project Manager</b></font></td></tr>"
				Set rsTBL = Server.CreateObject("ADODB.RecordSet")
				sqlTBL = "SELECT * FROM worker_t, consumer_t, conwork_t, Proj_Man_T " & _
					"WHERE PMID = Proj_Man_T.ID AND consumer_t.medicaid_number = CID AND CStr(worker_t.index) = WID " & _
					"AND status = 'Active' ORDER BY Proj_Man_T.Lname, Proj_Man_T.Fname, worker_t.lname, worker_T.fname, worker_t.City"

				rsTBL.Open sqlTBL, g_strCONN, 1, 3
				Do Until rsTBL.EOF	
					Set rsPM = Server.CreateObject("ADODB.RecordSet")
					sqlPM = "SELECT * FROM Proj_Man_T WHERE ID = " &  rsTBL("PMID")
					rsPM.Open sqlPM, g_strCONN, 1, 3
					If Not rsPM.EOF Then
						PMname = rsPM("lname") & ", " & rsPM("fname")
					Else
						PMname = rsPM("ID")
					End If
					rsPM.Close
					Set rsPM = Nothing
					strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("worker_t.lname") & ", " & _
						rsTBL("worker_t.fname") & "</font></td><td align='center'><font size='1' face='trebuchet ms'>" & _
						rsTBL("worker_t.Address") & ", " & rsTBL("worker_t.City") & ", " & rsTBL("worker_t.State") & ", " & _
						rsTBL("worker_t.Zip") & "</font></td><td align='center'>" & _
						"<font size='1' face='trebuchet ms'>&nbsp;" & rsTBL("worker_t.PhoneNo") & "&nbsp;</td><td align='center'><font size='1' " & _
						"face='trebuchet ms'>&nbsp;" & PMname & "</font></td></tr>"
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
					"face='trebuchet ms' color='white'><b>Phone No.</b></font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Project Manager</b></font>" & _
					"</td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Comment</b></font>" & _
					"</td></tr>"
				
				Set rsTBL = Server.CreateObject("ADODB.RecordSet")
		
				'sqlTBL = "SELECT * FROM Worker_t, w_log_t " & _
				'		"WHERE Worker_t.social_security_number = w_log_t.ssn AND " & _
				'		"status = 'Active' AND NOT IsNull(phonec) ORDER BY Worker_t.social_security_number, phonec DESC"
						
				sqlTBL = "SELECT * " & _ 
					"FROM Worker_t, w_log_t, consumer_t, conwork_t, proj_man_t  " & _
					"WHERE Worker_t.social_security_number = w_log_t.ssn " & _
					"AND PMID = Proj_Man_T.ID " & _
					"AND consumer_t.medicaid_number = CID " & _
					"AND CStr(worker_t.index) = WID " & _
					"AND status = 'Active' " & _
					"AND NOT IsNull(phonec) " & _
					"ORDER BY Proj_Man_T.Lname, Proj_Man_T.Fname, worker_T.social_security_number ASC, phonec DESC"
				
				rsTBL.Open sqlTBL, g_strCONN, 1, 3
				
				tmpID = ""
				x = 0
				
				Do Until rsTBL.EOF
					If rsTBL("worker_T.social_security_number") <> tmpID then 
						ReDim Preserve tmp(x)
						tmp(x) = rsTBL("w_log_t.index") & "|" & rsTBL("phoneC") & "|" & rsTBL("PMID") & "|" & rsTBL("medicaid_number")
						'tmp(x) = rsTBL("w_log_t.index") & "|" & rsTBL("phoneC")
						x = x + 1
					End If
					tmpID = rsTBL("worker_t.social_security_number")
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
						Set rsPM = Server.CreateObject("ADODB.RecordSet")
						sqlPM = "SELECT * FROM Proj_Man_T WHERE ID = " & tmp2(2)
						'sqlPM = "SELECT * FROM Proj_Man_T WHERE ID = " & rsTBL2("PMID")
						'response.write sqlPM
						rsPM.Open sqlPM, g_strCONN, 1, 3
						If Not rsPM.EOF Then
							PMname = rsPM("lname") & ", " & rsPM("fname")
						Else
							PMname = rsPM("ID")
						End If
						rsPM.Close
						Set rsPM = Nothing
						strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL2("phonec") & _
								"</td><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL2("lname") & ", " & _
								rsTBL2("fname") & "</font></td><td align='center'><font size='1' face='trebuchet ms'>" & _
								rsTBL2("Address") & ", " & rsTBL2("City") & ", " & rsTBL2("State") & ", " & rsTBL2("zip") & "</font></td><td align='center'>" & _
								"<font size='1' face='trebuchet ms'>&nbsp;" & rsTBL2("PhoneNo") & "</font></td><td align='center'>" & _
								"<font size='1' face='trebuchet ms'>" & PMname & "</font></td><td align='center' width='300px'>" & _
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
					"<font size='1' face='trebuchet ms' color='white'><b>Project Manager</b></font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Comment</b></font>" & _
					"</td></tr>"
				
				Set rsTBL = Server.CreateObject("ADODB.RecordSet")
		
				'sqlTBL = "SELECT * FROM Worker_t, w_log_t " & _
				'		"WHERE Worker_t.social_security_number = w_log_t.ssn AND " & _
				'		"status = 'Active' AND NOT IsNull(sitev) ORDER BY Worker_t.social_security_number, sitev DESC"
				
				sqlTBL = "SELECT * " & _ 
					"FROM Worker_t, w_log_t, consumer_t, conwork_t, proj_man_t  " & _
					"WHERE Worker_t.social_security_number = w_log_t.ssn " & _
					"AND PMID = Proj_Man_T.ID " & _
					"AND consumer_t.medicaid_number = CID " & _
					"AND CStr(worker_t.index) = WID " & _
					"AND status = 'Active' " & _
					"AND NOT IsNull(sitev) " & _
					"ORDER BY Proj_Man_T.Lname, Proj_Man_T.Fname, worker_T.social_security_number ASC, sitev DESC"
							
				rsTBL.Open sqlTBL, g_strCONN, 1, 3
				
				tmpID = ""
				x = 0
				
				Do Until rsTBL.EOF
					If rsTBL("worker_t.social_security_number") <> tmpID then 
						ReDim Preserve tmp(x)
						tmp(x) = rsTBL("w_log_t.index") & "|" & rsTBL("sitev") & "|" & rsTBL("PMID") 
						x = x + 1
					End If
					tmpID = rsTBL("worker_t.social_security_number")
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
						Set rsPM = Server.CreateObject("ADODB.RecordSet")
						'sqlPM = "SELECT * FROM Proj_Man_T WHERE ID = " & rsTBL2("PMID")
						sqlPM = "SELECT * FROM Proj_Man_T WHERE ID = " & tmp2(2)
						rsPM.Open sqlPM, g_strCONN, 1, 3
						If Not rsPM.EOF Then
							PMname = rsPM("lname") & ", " & rsPM("fname")
						Else
							PMname = rsPM("ID")
						End If
						rsPM.Close
						Set rsPM = Nothing
						strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL2("sitev") & _
								"</td><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL2("lname") & ", " & _
								rsTBL2("fname") & "</font></td><td align='center'><font size='1' face='trebuchet ms'>" & _
								rsTBL2("Address") & ", " & rsTBL2("City") & ", " & rsTBL2("State") & ", " & rsTBL2("zip") & "</font></td><td align='center'>" & _
								"<font size='1' face='trebuchet ms'>" & rsTBL2("PhoneNo") & "</font></td><td align='center'>" & _
								"<font size='1' face='trebuchet ms'>" & PMname & "</font></td><td align='center' width='300px'>" & _
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
					'If (Month(Request("FrmD8")) - 1) <> 0 Then 
					'		sqlProc = sqlProc & " AND Month(date) >= " & Month(Request("FrmD8")) - 1 & " " 
							sqlTBL = sqlTBL & " AND date >= " & CDate(Request("FrmD8")) - 7 & " "
					'	Else
					'		tmpYear = Year(Request("FrmD8")) - 1
					'		sqlProc = sqlProc & " AND date >= #12/1/" & tmpYear & "#"
					'	End If
					Session("Msg") = Session("Msg") & " from " & Request("FrmD8")
				Else
					Err = 1
				End If
			End If
			If Request("ToD8") <> "" Then
				If IsDate(Request("ToD8")) Then
					'If Month(Request("ToD8")) <> 1 Then
					'	sqlProc = sqlProc & " AND Month(date) - 1 <= " & Month(Request("ToD8")) & " " 
					'Else
						sqlProc = sqlProc & " AND date  <= #" & CDate(Request("ToD8")) + 7 & "#" 
					'End If
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
					Set rsName = Server.CreateObject("ADODB.RecordSet")
					sqlName = "SELECT lname, fname FROM Worker_t WHERE Social_Security_Number = '" & rsProc("emp_id") & "' "
					rsName.Open sqlName, g_strCONN, 1, 3
					If Not rsName.EOF Then
						tmpName = rsName("lname") & ", " & rsName("fname")
					Else
						tmpName = "N/A"
					End If
					rsName.Close
					Set rsName = Nothing
					Set rsName2 = Server.CreateObject("ADODB.RecordSet")
					sqlName2 = "SELECT lname, fname FROM consumer_T WHERE medicaid_number = '" & rsProc("client") & "' "
					rsName2.Open sqlName2, g_strCONN, 1, 3
					If Not rsName2.EOF Then
						tmpName2 = rsName2("lname") & ", " & rsName2("fname")
					Else
						tmpName2 = "N/A"
					End If
					rsName2.Close
					Set rsName2 = Nothing
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
					If rsProc("emp_id") <> tmpEID Then
					
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
							strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & rsProc("client") & "</font>" & vbCrLf & _
								"</td><td align='center'><font size='1' face='trebuchet ms'><b>*</b>" & tmpName2 & "</font></td><td align='center'>" & vbCrLf & _
								"<font size='1' face='trebuchet ms'>" & tmpTSWk1 & " - " & tmpTSWk2 & "</font></td>" & vbCrLf & _
								"<td align='center'><font size='1' face='trebuchet ms'>" & Z_FormatNumber(FtotHrs,2) & "</font></td><td align='center'>" & _
								"<textarea rows='1' readonly style='font-size: 10;' >" & rsProc("misc_notes") & "</textarea></td></tr>"
						Else
							tmpMSG = ""
							If mark = 0 Then tmpMSG = rsProc("misc_notes")
							strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & rsProc("client") & "</font></td>" & vbCrLf & _
								"<td align='center'><font size='1' face='trebuchet ms'>" & tmpName2 & "</font></td><td align='center'><font size='1' face='trebuchet ms'>" & vbCrLf & _
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
					"<font size='1' face='trebuchet ms' color='white'><b>Project Manager</b></font></td></tr>"
					
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
					"<font size='1' face='trebuchet ms' color='white'><b>Project Manager</b></font></td></tr>"
					
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
					Response.Redirect = "SpecRep.asp"
				End If
			End If 
			If Request("Todate") <> "" Then 
				If IsDate(Request("Todate")) Then
					Pdate2 = Request("Todate")
				Else
					Session("MSG") = "Enter valid date."
					Response.Redirect = "SpecRep.asp"
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
						"<font size='1' face='trebuchet ms' color='white'><b>Notes</b></font></td></tr>"
						maxfont = "black"
						If rsProc("MAX") = True Then maxfont = "Red"
						If Thours <> 0 Then
							'mark = 0
							If rsProc("EXT") = True Then
								mark = 1
								strBODY = strBODY & "<tr bgcolor='#F8F8FF'><td align='center'><font size='1' face='trebuchet ms' color='" & maxfont & "'>&nbsp;" & Right(rsProc("emp_id"), 4) & "&nbsp;</font></td><td align='center'>*<font size='1' face='trebuchet ms' color='" & maxfont & "'>" & _
									tmpName & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & maxfont & "'>&nbsp;" & rsProc("date") & " - " & Cdate(rsProc("date")) + 6 & _
									"&nbsp;</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & maxfont & "'>" & Z_FormatNumber(THours,2) & _
									"</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & maxfont & "'>&nbsp;" & rsProc("misc_notes") & "</font></td></tr>"
							Else
								tmpMSG = ""
								If mark = 0 Then tmpMSG = rsProc("misc_notes")
								strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms' color='" & maxfont & "'>&nbsp;" & Right(rsProc("emp_id"), 4) & "&nbsp;</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & maxfont & "'>" & _
									tmpName & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & maxfont & "'>&nbsp;" & rsProc("date") & " - " & Cdate(rsProc("date")) + 6 & _
									"&nbsp;</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & maxfont & "'>" & Z_FormatNumber(THours,2) & "</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & maxfont & "'>&nbsp;" & tmpMSG & "</font></td></tr>"
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
							strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>&nbsp;" & rsProc("client") & "&nbsp;</font></td><td align='center'><font size='1' face='trebuchet ms'>" & _
								tmpName & "</font></td><td align='center'><font size='1' face='trebuchet ms'>&nbsp;" & rsProc("date") & " - " & Cdate(rsProc("date")) + 6 & _
								"&nbsp;</font></td><td align='center'><font size='1' face='trebuchet ms'>" & Z_FormatNumber(THours,2) & "</font></td></tr>"
					  End if
						
					End If
					
					rsProc.MoveNext
				Loop
				Session("MSG") = "Records from " & Pdate & " to " & Pdate2 & " for "
				If Request("seltype") = 1 Then
						Session("MSG") = Session("MSG") & "payroll report. <br>* extended hours<br>red font - over max hours"
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
						"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Ammendment Expiration Date</b></font></td></tr>"
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
						rsTBL("End_Date") &  "</font></td></tr>"
					rsTBL.MoveNext
				Loop
				rsTBL.Close
				Set rsTBL = Nothing	
			ElseIf Request("SelRep") = 30 Then 
				Session("MSG") = "All Active PCSP Worker Extended Hours report"
				Set rsTBL = Server.CreateObject("ADODB.RecordSet")
				sqlTBL = "SELECT * FROM tsheets_T, worker_T WHERE EXT = true AND emp_id = social_security_number ORDER BY lname, fname ASC, date DESC"
				rsTBL.Open sqlTBL, g_strCONN, 3, 1
				strHEAD = "<tr bgcolor='#040C8B'><td align='center'><font size='1' face='trebuchet ms' color='white'><b>SSN</b></font></td><td align='center'>" & _
						"<font size='1' face='trebuchet ms' color='white'><b>Name</b></font></td><td align='center'><font size='1' " & _
						"face='trebuchet ms' color='white'><b>Timesheet Week</b></font></td><td align='center'>" & _
						"<font size='1' face='trebuchet ms' color='white'><b>Hours</b></font></td><td align='center'>" & _
						"<font size='1' face='trebuchet ms' color='white'><b>Notes</b></font></td></tr>"
				Do Until rsTBL.EOF
					THours = rsTBL("mon") + rsTBL("tue") + rsTBL("wed") + rsTBL("thu") + rsTBL("fri") + rsTBL("sat") + rsTBL("sun")
					tmpName = rsTBL("lname") & ", " & rsTBL("fname")
					If Thours <> 0 Then
						strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>&nbsp;" & Right(rsTBL("emp_id"), 4) & "&nbsp;</font></td><td align='center'><font size='1' face='trebuchet ms'>" & _
									tmpName & "</font></td><td align='center'><font size='1' face='trebuchet ms'>&nbsp;" & rsTBL("date") & " - " & Cdate(rsTBL("date")) + 6 & _
									"&nbsp;</font></td><td align='center'><font size='1' face='trebuchet ms'>" & Z_FormatNumber(THours,2) & _
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
						"<font size='1' face='trebuchet ms' color='white'><b>Notes</b></font></td></tr>"
				End If
				Set rsTBL = Server.CreateObject("ADODB.RecordSet")
				sqlTBL = "SELECT * FROM [Tsheets_t]"
				If Request("seltype2") = 1 Then
					sqlTBL = sqlTBL & ", worker_t  WHERE emp_id = social_security_number "
				ElseIf Request("seltype2") = 2 Then
					sqlTBL = sqlTBL & ", consumer_t  WHERE client = medicaid_number"
				End If
				Err = 0
				If Request("FrmD8") <> "" Then
					If IsDate(Request("FrmD8")) Then
						'If (Month(Request("FrmD8")) - 1) <> 0 Then 
						'	sqlTBL = sqlTBL & " AND Month(date) >= " & Month(Request("FrmD8")) - 1 & " " 
							sqlTBL = sqlTBL & " AND date >= " & CDate(Request("FrmD8")) - 7 & " "
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
					'	If Month(Request("ToD8")) <> 1 Then
					'		sqlTBL = sqlTBL & " AND Month(date) - 1 <= " & Month(Request("ToD8")) & " " 
					'	Else
							sqlTBL = sqlTBL & " AND date  <= #" & CDate(Request("ToD8")) + 7 & "#" 
					'	End If
						Session("Msg") = Session("Msg") & " to " & Request("ToD8")
					Else
						Err = 1
					End If
				End If
				If Request("seltype2") = 1 Then
					sqlTBL = sqlTBL & " ORDER BY lname ASC, fname ASC, date DESC"
				ElseIf Request("seltype2") = 2 Then
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
							If Request("seltype2") = 2 Then
								strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms' >&nbsp;" & rsTBL("client") & "&nbsp;</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & maxfont & "'>" & _
											GetName2(rsTBL("client")) & "</font></td><td align='center'><font size='1' face='trebuchet ms' >&nbsp;" & tmpTSWk1 & " - " & tmpTSWk2 & _
											"&nbsp;</font></td><td align='center'><font size='1' face='trebuchet ms' >" & Z_FormatNumber(FtotHrs,2) & _
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
						"<font size='1' face='trebuchet ms' color='white'><b>Name</b></font></td><td align='center'><font size='1' " & _
						"face='trebuchet ms' color='white'><b>Timesheet Week</b></font></td><td align='center'>" & _
						"<font size='1' face='trebuchet ms' color='white'><b>Hours</b></font></td><td align='center'>" & _
						"<font size='1' face='trebuchet ms' color='white'><b>Notes</b></font></td></tr>"
				Set rsTBL = Server.CreateObject("ADODB.RecordSet")
				sqlTBL = "SELECT * FROM tsheets_T, worker_T WHERE emp_id = Social_Security_Number AND ext = true"
				Err = 0
				If Request("FrmD8") <> "" Then
					If IsDate(Request("FrmD8")) Then
						'If (Month(Request("FrmD8")) - 1) <> 0 Then 
					'		sqlProc = sqlProc & " AND Month(date) >= " & Month(Request("FrmD8")) - 1 & " " 
							sqlTBL = sqlTBL & " AND date >= " & CDate(Request("FrmD8")) - 7 & " "
					'	Else
					'		tmpYear = Year(Request("FrmD8")) - 1
					'		sqlProc = sqlProc & " AND date >= #12/1/" & tmpYear & "#"
					'	End If
						Session("Msg") = Session("Msg") & " from " & Request("FrmD8")
					Else
						Err = 1
					End If
				End If
				If Request("ToD8") <> "" Then
					If IsDate(Request("ToD8")) Then
						'If Month(Request("ToD8")) <> 1 Then
					'	sqlProc = sqlProc & " AND Month(date) - 1 <= " & Month(Request("ToD8")) & " " 
					'Else
						sqlProc = sqlProc & " AND date  <= #" & CDate(Request("ToD8")) + 7 & "#" 
					'End If
						Session("Msg") = Session("Msg") & " to " & Request("ToD8")
					Else
						Err = 1
					End If
				End If
				sqlTBL  = sqlTBL  & " ORDER BY lname, fname, date, timestamp"
				Session("Msg") = Session("Msg") & ". " 
				If Err <> 0 Then Response.Redirect "specrep.asp?err=32" 
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
							strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms' >&nbsp;" & Right(rsTBL("emp_id"), 4) & "&nbsp;</font></td><td align='center'><font size='1' face='trebuchet ms' color='" & maxfont & "'>" & _
										GetName(rsTBL("emp_id")) & "</font></td><td align='center'><font size='1' face='trebuchet ms' >&nbsp;" & tmpTSWK1 & " - " & tmpTSWK2 & _
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
						"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Driver's License Expiration Date</b></font></td></tr>"
				Set rsTBL = Server.CreateObject("ADODB.RecordSet")
				sqlTBL = "SELECT fname, lname, address, city, state, zip, status,LicenseExpDate FROM worker_t " & _
					"WHERE status = 'Active' AND Driver = True ORDER BY LicenseExpDate, lname, fname"
				rsTBL.Open sqlTBL, g_strCONN, 1, 3
				Do Until rsTBL.EOF	
					strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("lname") & ", " & _
						rsTBL("fname") & "</font></td><td align='center'><font size='1' face='trebuchet ms'>" & _
						rsTBL("Address") & ", " & rsTBL("City") & ", " & rsTBL("State") & ", " & rsTBL("Zip") & "</font></td><td align='center'>" & _
						"<font size='1' face='trebuchet ms'>&nbsp;" & rsTBL("LicenseExpDate") & "</td></tr>"
					rsTBL.MoveNext
				Loop
				rsTBL.Close 
			ElseIf Request("SelRep") = 33 Then
				Session("MSG") = "PCSP Worker by Driver's License Expiration Date report."
				strHEAD = "<tr bgcolor='#A4CADB'><td align='center'><font size='1' face='trebuchet ms'><b>PCSP Worker Name</b>" & _
						"</font></td><td align='center'><font size='1' face='trebuchet ms'><b>Address</b></font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms'><b>Driver's License Expiration Date</b></font></td></tr>"
				Set rsTBL = Server.CreateObject("ADODB.RecordSet")
				sqlTBL = "SELECT fname, lname, address, city, state, zip, status,LicenseExpDate FROM worker_t " & _
					"WHERE status = 'Active' AND Driver = True ORDER BY LicenseExpDate, lname, fname"
				rsTBL.Open sqlTBL, g_strCONN, 1, 3
				Do Until rsTBL.EOF	
					strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("lname") & ", " & _
						rsTBL("fname") & "</font></td><td align='center'><font size='1' face='trebuchet ms'>" & _
						rsTBL("Address") & ", " & rsTBL("City") & ", " & rsTBL("State") & ", " & rsTBL("Zip") & "</font></td><td align='center'>" & _
						"<font size='1' face='trebuchet ms'>&nbsp;" & rsTBL("LicenseExpDate") & "</td></tr>"
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
						rsTBL("mAddress") & ", " & rsTBL("mCity") & ", " & rsTBL("mState") & ", " & rsTBL("mZip") & "</font>" & _
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
					"FROM Representative_T, consumer_T, Conrep_T " & _
					"WHERE Cint(RID) = Representative_T.index " & _
					"AND CID = Consumer_t.medicaid_Number " & _ 
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
						"</font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Address</b></font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Consumer Name</b></font></td></tr>"
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
						"</td><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("Consumer_T.lname") & ", " & _
						rsTBL("Consumer_T.fname") & "</font></td></tr>"
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
						"</font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Current Care Plan</b></font></td></tr>"
				Set rsTBL = Server.CreateObject("ADODB.RecordSet")
				sqlTBL = "SELECT * " & _
					"FROM consumer_T, C_Status_T " & _
					"WHERE consumer_T.medicaid_number = C_Status_T.medicaid_number " & _
					"AND Active = True " & _ 
					"ORDER BY Lname, Fname"
				rsTBL.Open sqlTBL, g_strCONN, 1, 3
				Do Until rsTBL.EOF	
					strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("lname") & ", " & _
						rsTBL("fname") & "</font></td><td align='center'><font size='1' face='trebuchet ms'>" & _
						rsTBL("CarePlan") & "&nbsp;</font></td></tr>"
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
						"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Project Manager</b></font></td></tr>"
				Set rsTBL = Server.CreateObject("ADODB.RecordSet")
				sqlTBL = "SELECT * " & _
					"FROM worker_T, consumer_T, ConWork_T, Proj_man_T " & _
					"WHERE Status = 'Active' " & _
					"AND PMID = Proj_man_T.ID " & _
					"AND worker_T.index = Cint(WID) " & _
					"AND CID = Consumer_t.medicaid_Number " & _
					"ORDER BY Proj_man_t.Lname, Proj_man_t.Fname, worker_T.Lname, worker_T.Fname "
				rsTBL.Open sqlTBL, g_strCONN, 1, 3
				Do Until rsTBL.EOF	
					Set rsChkMe = Server.CreateObject("ADODB.RecordSet")
					sqlChkMe = "SELECT * FROM Tsheets_T WHERE emp_id = '" & rsTBL("worker_T.Social_security_number") & "' AND [Date] = #" & _
						Pdate & "# "'AND [Date] >= #" & Pdate2 & "# "
					rsChkMe.Open sqlChkMe, g_strCONN, 3, 1
					If rsChkMe.EOF Then
						strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("worker_T.lname") & ", " & _
							rsTBL("worker_T.fname") & "</font></td><td align='center'><font size='1' face='trebuchet ms'>" & _
							Pdate & " - " & Cdate(Pdate) + 6 & "&nbsp;</font></td><td align='center'><font size='1' face='trebuchet ms'>" & _
							rsTBL("Proj_man_t.lname") & ", " & rsTBL("Proj_man_t.fname") & "</font></td></tr>"
					End If
					rsChkMe.Close
					Set rsChkMe = Nothing
					Set rsChkMe2 = Server.CreateObject("ADODB.RecordSet")
					sqlChkMe2 = "SELECT * FROM Tsheets_T WHERE emp_id = '" & rsTBL("worker_T.Social_security_number") & "' AND [Date] = #" & _
						Cdate(Pdate2) - 6 & "# "'AND [Date] >= #" & Pdate2 & "# "
					rsChkMe2.Open sqlChkMe2, g_strCONN, 3, 1
					If rsChkMe2.EOF Then
						strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("worker_T.lname") & ", " & _
							rsTBL("worker_T.fname") & "</font></td><td align='center'><font size='1' face='trebuchet ms'>" & _
							Cdate(Pdate2) - 6 & " - " & Pdate2 & "&nbsp;</font></td><td align='center'><font size='1' face='trebuchet ms'>" & _
							rsTBL("Proj_man_t.lname") & ", " & rsTBL("Proj_man_t.fname") & "</font></td></tr>"
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
						"</font></td><td align='center'><font size='1' face='trebuchet ms' color='white'><b>Start Date</b></font></td>" & _
						"<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Inactive Date</b></font></td></tr>"
				Set rsTBL = Server.CreateObject("ADODB.RecordSet")
				sqlTBL = "SELECT * " & _
					"FROM consumer_T, C_Status_T " & _
					"WHERE consumer_T.medicaid_number = C_Status_T.medicaid_number " & _
					"ORDER BY Lname, Fname"
				rsTBL.Open sqlTBL, g_strCONN, 1, 3
				Do Until rsTBL.EOF	
					strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL("lname") & ", " & _
						rsTBL("fname") & "</font></td><td align='center'><font size='1' face='trebuchet ms'>" & _
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
						Else
							strHEAD = strHEAD & "<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Site Visit Date</b></font></td>"
							strHEAD = strHEAD & "<td align='center'><font size='1' face='trebuchet ms' color='white'><b>Phone Call Date</b></font></td>"
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
						Else
							sqlTBL = sqlTBL & " AND (Site_V_Date >= #" & Request("FrmD8") & "#"
							sqlTBL = sqlTBL & " OR phoneCon_last >= #" & Request("FrmD8") & "#)"
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
						Else
							sqlTBL = sqlTBL & " AND (Site_V_Date <= #" & Request("ToD8") & "#"
							sqlTBL = sqlTBL & " OR phoneCon_last <= #" & Request("ToD8") & "#)"
							Session("Msg") = Session("Msg") & " to " & Request("ToD8")
						End If
					End If
				End If
				If Err <> 0 Then  Response.Redirect "specrep.asp?err=41" 
				If Request("SelLog") = 1 Then
					sqlTBL = sqlTBL & " ORDER BY Site_V_Date DESC, Lname ASC, Fname ASC"
				ElseIf Request("SelLog") = 2 Then
					sqlTBL = sqlTBL & " ORDER BY phoneCon_last DESC, Lname ASC, Fname ASC"
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
						End If
					Else
						ReDim Preserve tmp(x)
						tmp(x) = rsTBL("C_Site_Visit_Dates_t.index") & "|" & rsTBL("Site_V_date") & "|" & rsTBL("phoneCon_last")
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
							if tmpj(1) = "" Then tmpDateLog = tmpj(2)
							if tmpj1(1) = "" Then tmpDateLog1 = tmpj1(2)
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
						sqlTBL2 = "SELECT * FROM Consumer_T, C_Site_Visit_Dates_T WHERE 	 Consumer_t.Medicaid_number =" & _
							" C_Site_Visit_Dates_t.Medicaid_number AND C_Site_Visit_Dates_t.index = " & tmp2(0)
							
						rsTBL2.Open sqlTBL2, g_strCONN, 1, 3
						If Not rsTBL2.EOF Then
							strBODY = strBODY & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & rsTBL2("lname") & ", " & rsTBL2("fname") & "</font></td>" & _
								"<td align='center'><font size='1' face='trebuchet ms'>" & rsTBL2("Site_V_Date") & "&nbsp;</font></td>" & _
								"<td align='center'><font size='1' face='trebuchet ms'>" & rsTBL2("phoneCon_last") & "&nbsp;</font></td>" 
							If rsTBL2("Site_V_Date") <> "" Then 
								newComment = Replace(rsTBL2("Comments"), "|",  " ")
								strBODY = strBODY & "<td align='center'><font size='1' face='trebuchet ms'>" & newComment & "&nbsp;</font></td></tr>"
							Else
								 strBODY = strBODY & "<td align='center'><font size='1' face='trebuchet ms'>" & rsTBL2("PCom") & "&nbsp;</font></td></tr>"
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
			End If
			
			Session("PrintPrev") = strHEAD & "|" & strBODY & "|" & Session("MSG")
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
		function ExpCSV()
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
		}
		function hidetxt()
		{
			if (document.frmRep.SelRep.value == 18 || document.frmRep.SelRep.value == 24 || document.frmRep.SelRep.value == 29 || document.frmRep.SelRep.value == 32 || document.frmRep.SelRep.value == 31 || document.frmRep.SelRep.value == 41)
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
				}
			else
				{
				 document.frmRep.seltype2.style.visibility = 'hidden';
				 document.frmRep.FrmD8.style.visibility = 'hidden';
				 document.frmRep.ToD8.style.visibility = 'hidden';
				 document.frmRep.txtFrm.style.visibility = 'hidden';
				 document.frmRep.txtTo.style.visibility = 'hidden';
				}
		}
		function hided8()
		{
			if (document.frmRep.SelRep.value == 27 || document.frmRep.SelRep.value == 39)
				{document.frmRep.closedate.style.visibility = 'visible';
				 document.frmRep.Todate.style.visibility = 'visible';
				 if (document.frmRep.SelRep.value == 27)
				 {document.frmRep.seltype.style.visibility = 'visible';}
				 else
				 {document.frmRep.seltype.style.visibility = 'hidden';}
				 document.frmRep.txtCal.style.visibility = 'visible';
				 document.frmRep.txtTCal.style.visibility = 'visible';
				 document.frmRep.cal1.style.visibility = 'visible';}
			else
				{document.frmRep.Todate.style.visibility = 'hidden';
				 document.frmRep.closedate.style.visibility = 'hidden';
				 document.frmRep.seltype.style.visibility = 'hidden';
				 document.frmRep.txtCal.style.visibility = 'hidden';
				 document.frmRep.txtTCal.style.visibility = 'hidden';
				 document.frmRep.cal1.style.visibility = 'hidden';}
			}
			function weeknum()
			{
				document.frmRep.action = "weeknum.asp?tmpdate=" + document.frmRep.closedate.value;
				document.frmRep.submit();
			}
			function PrintPrev()
			{
				document.frmRep.action = "Print.asp";
				document.frmRep.submit();
			}
			function hideCon()
			{
				if (document.frmRep.SelRep.value == 41)
					{
					document.frmRep.txtCon.style.visibility = 'visible';
					document.frmRep.SelCon.style.visibility = 'visible';
					document.frmRep.SelLog.style.visibility = 'visible';
					}
				else
					{
					document.frmRep.txtCon.style.visibility = 'hidden';
					document.frmRep.SelCon.style.visibility = 'hidden';
					document.frmRep.SelLog.style.visibility = 'hidden';
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
	<body bgcolor='white'  LEFTMARGIN='0' TOPMARGIN='0' onload='hidetxt(); hided8();hideCon(); '>
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
					<td align='center'><font size='1' face='trebuchet MS'>Type:</font></td>
					<td colspan='3'>
						<select name='SelRep'  onchange='hidetxt(); hided8(); hideCon();' >
							<option value='2' <%=Sel2%>>Active Consumers</option>
							<option value='3' <%=Sel3%>>Active PCSP Workers</option>
							<option value='36' <%=Sel36%>>Case Manager List</option>
							<option value='5' <%=Sel5%>>Census Information</option>
							<option value='14' <%=Sel14%>>Consumers by Project Manager And Town</option>
							<option value='6' <%=Sel6%>>Consumers by Town</option>
							<option value='38' <%=Sel38%>>Consumers Current Care Plan</option>
							<option value='19' <%=Sel19%>>Consumer Date Of Birth</option>
							<option value='16' <%=Sel16%>>Consumer Health</option>
							<option value='41' <%=Sel41%>>Consumer Logs</option>
							<option value='29' <%=Sel29%>>Consumer Start and Ammendment Expiration Date</option>
							<option value='18' <%=Sel18%>>Consumer Start and End Date</option>
							<option value='40' <%=Sel40%>>Consumer Start and Inactive Date</option>
							<option value='17' <%=Sel17%>>Consumer Start Date</option>
							<option value='26' <%=Sel26%>>Consumer with No Log</option>
							<option value='1' <%=Sel1%>>Consumers with PCSP Worker and Hours</option>
							<option value='37' <%=Sel37%>>Inactive Consumers List</option>
							<option value='4' <%=Sel4%>>Inactive Consumers with PCSP Worker and Hours</option>
							<option value='33' <%=Sel33%>>PCSP Worker by Drivers License Expiration Date</option>
							<option value='15' <%=Sel15%>>PCSP Worker by Insurance Expiration Date</option>
							<option value='20' <%=Sel20%>>PCSP Workers by Project Manager And Town</option>
							<option value='7' <%=Sel7%>>PCSP Workers by Town</option>
							<option value='8' <%=Sel8%>>PCSP Workers Date of Hire</option>
							<option value='30' <%=Sel30%>>PCSP Workers Extended Hours</option>
							<option value='9' <%=Sel9%>>PCSP Workers Interested in More Consumers</option>
							<option value='28' <%=Sel28%>>PCSP Workers Mailing Address</option>
							<option value='42' <%=Sel42%>>PCSP Workers Separation Code</option>
							<option value='21' <%=Sel21%>>PCSP Workers (Inactive) Termination Date</option>
							<option value='24' <%=Sel24%>>PCSP Workers Total Hours</option>
							<option value='34' <%=Sel34%>>PCSP Workers with Active Consumers</option>
							<option value='10' <%=Sel10%>>PCSP Workers with No Consumer</option>
							<option value='25' <%=Sel25%>>PCSP Workers with No Log</option>
							<option value='39' <%=Sel39%>>PCSP Workers with Unsubmitted Timesheets</option>
							<option value='11' <%=Sel11%>>Phone Contact for Consumers</option>
							<option value='22' <%=Sel22%>>Phone Contact for PCSP Worker</option>
							<option value='12' <%=Sel12%>>Referrals</option>
							<option value='35' <%=Sel35%>>Representative List</option>
							<option value='13' <%=Sel13%>>Site Visit for Consumers</option>
							<option value='23' <%=Sel23%>>Site Visit for PCSP Worker</option>
							<option value='27' <%=Sel27%>>* History -  Timesheet / Medicaid </option>
							<option value='31' <%=Sel31%>>* Date Range -  Payroll / Medicaid</option>
							<option value='32' <%=Sel32%>>* Date Range -  Extended Hours</option>
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
						</select>
					</td>
				</tr>
				<tr>
					<td>&nbsp;</td>
					<td align='right'>
						<input name='txtCal' style='width: 40px; border: none;' readonly value='From:'>
						<input tabindex="1" name='closedate' style='width:80px;' value='<%=sunDATE%>'
						type="text" onchange='weeknum();' readonly><input tabindex="2" type="button" value="..." name="cal1" style="width: 15px;"
						onclick="showCalendarControl(document.frmRep.closedate);" class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'"> &nbsp;
						<input name='txtTCal' style='width: 40px; border: none;' readonly value='To:'>
						<input tabindex="1" name='Todate' style='width:80px;' readonly value='<%=satDATE%>'
						type="text">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
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
						</select>
					</td>
				</tr>
				<tr><td>&nbsp;</td></tr>
				<tr><td colspan='4' align='center'><font color='red' face='trebuchet MS' size='1'>&nbsp;<%=Session("MSG")%>&nbsp;</font></td></tr>
			</table>
			<br>
			<% If strBODY <> "" Then%>
				<center>
				<% If Request("SelRep") = 11 Or Request("SelRep") = 13 Or Request("SelRep") = 22 Or Request("SelRep") = 23 Then%>
					<input  style='width: 140px;' align='center' type='button' class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" value='Print Preview' onclick='JavaScript: LandWarn();'>
				<% Else %>
					<input  style='width: 140px;' align='center' type='button' class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" value='Print Preview' onclick='PrintPrev();'>
				<% End If%>
				<% If Request("SelRep") = 19 Or Request("SelRep") = 28 Or Request("SelRep") = 2 Or Request("SelRep") = 3 Or Request("SelRep") = 35 Then%>
					<input  style='width: 140px;' align='center' type='button' class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" value='Export to CSV' onclick='JavaScript: ExpCSV();'>
				<% End If %>
		<% End If%>
			<br><br>
			<table cellSpacing='0' cellPadding='0' align='center' border='1'>
				<%=strHEAD%>
				<%=strBODY%>
			</table>
			</form>
		<!-- #include file="_boxdown.asp" -->
	</body>
</html>
<%
Session("MSG") = ""
%>