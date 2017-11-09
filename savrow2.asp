<%@Language=VBScript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<%
DIM	lngI, tblEMP, strSQL, strTableScript, tmpID, ctrI, tmpctr
'SAVE ROW FOR WEEK 1
tmpTimeStamp = Now
	ctrI = Request("count")
	If Request("count") <> 0 Then
		'EXISTING HOURS ROW
		For i = 0 to ctrI 
			Set tblEMP = Server.CreateObject("ADODB.Recordset")
			strSQL = "SELECT * FROM [tsheets_t]"
			tblEMP.Open strSQL, g_strCONN, 1, 3
			tblEMP.Movefirst
			tmpctr = Request("chk" & i)
			If tmpctr <> "" Then
				strTmp = "ID='" & tmpctr & "' "
				tblEMP.Find(strTmp)
				If Not tblEMP.EOF Then
					'''''''ERROR CHECK''''''''
					'GET TOTAL EXTENDED HOURS
					tmpXHRS = 0
					tmpXHRS = Z_CZero(Request("hsunx1" & i)) + Z_CZero(Request("hmonx1" & i)) +Z_CZero(Request("htuex1" & i)) +Z_CZero(Request("hwedx1" & i)) + _
						Z_CZero(Request("hthux1" & i)) + Z_CZero(Request("hfrix1" & i)) + Z_CZero(Request("hsatx1" & i))
					'GET TOTAL HOURS
					tmpTotHRS = 0
					tmpTotHRS = Z_CZero(Request("hsunx1" & i)) + Z_CZero(Request("hmonx1" & i)) +Z_CZero(Request("htuex1" & i)) +Z_CZero(Request("hwedx1" & i)) + _
						Z_CZero(Request("hthux1" & i)) + Z_CZero(Request("hfrix1" & i)) + Z_CZero(Request("hsatx1" & i)) + Z_CZero(Request("hsun1" & i)) + _
						Z_CZero(Request("hmon1" & i)) + Z_CZero(Request("htue1" & i)) + Z_CZero(Request("hwed1" & i)) + _
						Z_CZero(Request("hthu1" & i)) + Z_CZero(Request("hfri1" & i)) + Z_CZero(Request("hsat1" & i))
					If tmpXHRS <> 0 And Request("Mnotes1" & i) = "" Then
						Session("MSG") = Session("MSG") & "<br>Notes are required if there are extended hours in week 1 for Consumer " & Request("hdept1" & i) & "." 
					ElseIf tmpTotHRS = 0 Then
						Session("MSG") = Session("MSG") & "<br>Total hours for " & Request("hdept1" & i) & " will become zero(0) in week 1. Please just delete the row."
					Else
						MaxHrs = "MaxHrs"
						If Instr(Request("Mnotes1" & i), "80,") > 0 Then 
							MaxHrs = "vahmhrs"
						ElseIf Instr(Request("Mnotes1" & i), "82,") > 0 Then 
							MaxHrs = "vahahrs"
						End If
						'GET MAX HOURS AND MILEAGE CAP
						Set tblCon = Server.CreateObject("ADODB.RecordSet")
						sqlCon = "SELECT " & MaxHrs & ", milecap FROM Consumer_t WHERE medicaid_number = '" & Request("conid1" & i) & "' "
						tblCon.Open sqlCon, g_strCONN, 1, 3
						If Not tblCon.EOF Then
							tmpMax = Z_CZero(tblCon(MaxHrs))
							tmpMile = Z_CZero(tblCon("milecap"))
						End If
						tblCon.Close
						Set tblCon = Nothing
						'GET EXISTING HOURS AND MILEAGE
						Set rsChk = Server.CreateObject("ADODB.RecordSet")
						sqlChk = "SELECT * FROM tSHEETS_T WHERE date = '" & Request("1day") & "' AND client = '" & _
							Request("conid1" & i) & "' AND Ext = 0"
						If Instr(Request("Mnotes1" & i), "80,") > 0 Then 
							sqlChk = sqlChk & " AND misc_notes = '80,'"
						ElseIf Instr(Request("Mnotes1" & i), "82,") > 0 Then 
							sqlChk = sqlChk & " AND misc_notes = '82,'"
						End If 
						rsChk.Open sqlChk, g_strCONN, 3, 1
						If Not rsChk.EOF Then
							HrsCon = 0
							MileCon = 0
							Do Until rsChk.EOF
								HrsCon = HrsCon + Z_CZero(rsChk("mon")) + Z_CZero(rsChk("tue")) + Z_CZero(rsChk("wed")) + Z_CZero(rsChk("thu")) + Z_CZero(rsChk("fri")) + _
									Z_CZero(rsChk("sat")) + Z_CZero(rsChk("sun"))
								MileCon = MileCon +  Z_CZero(rsChk("mile"))
								rsChk.MoveNext
							Loop
						End If	
						rsChk.Close
						Set rsChk = Nothing
						'GET HOURS AND MILEAGE
						tmpHRS = 0
						tmpHRS = HrsCon + Z_CZero(Request("hsun1" & i)) + _
							Z_CZero(Request("hmon1" & i)) + Z_CZero(Request("htue1" & i)) + Z_CZero(Request("hwed1" & i)) + _
							Z_CZero(Request("hthu1" & i)) + Z_CZero(Request("hfri1" & i)) + Z_CZero(Request("hsat1" & i))
						tmpMileTot = MileCon + Z_CZero(Request("txtmile" & i))
						If tmpHRS > tmpMax Then
							Session("MSG") = Session("MSG") & "<br>Total hours for " &  Request("hdept1" & i) & " is over the allowed hours in week 1. This row will be red-flagged on History report and Process Items report."
							tblEMP("MAX") = True
							'SET MAX = TRUE ON DB FOR ALL WORKERS
							Set rsMax = Server.CreateObject("ADODB.RecordSet")
							sqlMax = "SELECT * FROM tSHEETS_T WHERE date = '" & Request("1day") & "' AND client = '" & _
								Request("conid1" & i) & "' AND Ext = 0 AND MAX = 0"
							If Instr(Request("Mnotes1" & i), "80,") > 0 Then 
								sqlMax = sqlMax & " AND misc_notes = '80,'"
							ElseIf Instr(Request("Mnotes1" & i), "82,") > 0 Then 
								sqlMax = sqlMax & " AND misc_notes = '82,'"
							End If 
							rsMax.Open sqlMax, g_strCONN, 1, 3
							If Not rsMax.EOF Then
								Do Until rsMAX.EOF
									rsMAX("MAX") = True
									rsMAX.Update
									rsMAX.MoveNext
								Loop
							End If
							rsMAX.Close
							Set rsMax = Nothing
						Else
							tblEMP("MAX") = False
							'SET MAX = FALSE ON DB FOR ALL WORKERS
							Set rsMax = Server.CreateObject("ADODB.RecordSet")
							sqlMax = "SELECT * FROM tSHEETS_T WHERE date = '" & Request("1day") & "' AND client = '" & _
								Request("conid1" & i) & "' AND Ext = 0 AND MAX = 1"
							If Instr(Request("Mnotes1" & i), "80,") > 0 Then 
								sqlMax = sqlMax & " AND misc_notes = '80,'"
							ElseIf Instr(Request("Mnotes1" & i), "82,") > 0 Then 
								sqlMax = sqlMax & " AND misc_notes = '82,'"
							End If 
							rsMax.Open sqlMax, g_strCONN, 1, 3
							If Not rsMax.EOF Then
								Do Until rsMAX.EOF
									rsMAX("MAX") = False
									rsMAX.Update
									rsMAX.MoveNext
								Loop
							End If
							rsMAX.Close
							Set rsMax = Nothing
						End If
						''''''
						If tmpMileTot > tmpMile Then
							Session("MSG") = Session("MSG") & "<br>Total Mileage for " &  Request("hdept1" & i) & " is over the allowed mileage in week 1. This row will be flagged on History report and Process Items report."
							tblEMP("milecap") = True
							'SET milecap = TRUE ON DB FOR ALL WORKERS
							Set rsMax = Server.CreateObject("ADODB.RecordSet")
							sqlMax = "SELECT * FROM tSHEETS_T WHERE date = '" & Request("1day") & "' AND client = '" & _
								Request("conid1" & i) & "' AND Ext = 0 AND emp_id <> '" & Session("idemp") & "'"
							rsMax.Open sqlMax, g_strCONN, 1, 3
							If Not rsMax.EOF Then
								Do Until rsMAX.EOF
									rsMAX("milecap") = True
									rsMAX.Update
									rsMAX.MoveNext
								Loop
							End If
							rsMAX.Close
							Set rsMax = Nothing
						Else
							tblEMP("milecap") = False
							'SET milecap = FALSE ON DB FOR ALL WORKERS
							Set rsMax = Server.CreateObject("ADODB.RecordSet")
							sqlMax = "SELECT * FROM tSHEETS_T WHERE date = '" & Request("1day") & "' AND client = '" & _
								Request("conid1" & i) & "' AND Ext = 0 AND emp_id <> '" & Session("idemp") & "'"
							rsMax.Open sqlMax, g_strCONN, 1, 3
							If Not rsMax.EOF Then
								Do Until rsMAX.EOF
									rsMAX("milecap") = False
									rsMAX.Update
									rsMAX.MoveNext
								Loop
							End If
							rsMAX.Close
							Set rsMax = Nothing
						End If
						'SAVE HOURS
						If Not IsNumeric(Request("hmon1" & i)) Then
							tblEMP("mon") = 0
						Else
							tblEMP("mon") = Request("hmon1" & i)
						End If
						If Not IsNumeric(Request("htue1" & i)) Then
							tblEMP("tue") = 0
						Else
							tblEMP("tue") = Request("htue1" & i)
						End If
						If Not IsNumeric(Request("hwed1" & i)) Then
							tblEMP("wed") = 0
						Else
							tblEMP("wed") = Request("hwed1" & i)
						End If
						If Not IsNumeric(Request("hthu1" & i)) Then
							tblEMP("thu") = 0
						Else
							tblEMP("thu") = Request("hthu1" & i)
						End If
						If Not IsNumeric(Request("hfri1" & i)) Then
							tblEMP("fri") = 0
						Else
							tblEMP("fri") = Request("hfri1" & i)
						End If
						If Not IsNumeric(Request("hsat1" & i)) Then
							tblEMP("sat") = 0
						Else
							tblEMP("sat") = Request("hsat1" & i)
						End If
						If Not IsNumeric(Request("hsun1" & i)) Then
							tblEMP("sun") = 0
						Else
							tblEMP("sun") = Request("hsun1" & i)
						End If	
						'If Not IsNumeric(Request("txtmile" & i)) Then
						'	tblEMP("mile") = 0
						'Else
							tblEMP("mile") = Z_Czero(Request("txtmile" & i))
						'End If
						If Not IsNumeric(Request("txtamile" & i)) Then
							tblEMP("amile") = 0
						Else
							tblEMP("amile") = Request("txtamile" & i)
						End If
						tblEMP("misc_notes") = Request("Mnotes1" & i)
						tblEMP("actcode") = Request("com1" & i)
						tblEMP("author") = session("UserID")
						tmpTS = tblEMP("timestamp")
						tblEMP.Update
						'SAVE EXTENDED HOURS
						Set rsEXT = Server.CreateObject("ADODB.RecordSet")
						sqlEXT = "SELECT * FROM [tsheets_t] WHERE " & _
							"[emp_id] = '" & tblEMP("emp_ID") & "' AND " & _
							"[date] = '" & tblEMP("date") &  "' AND EXT = 1 AND client = '" & tblEMP("Client") & _
								"' AND timestamp = '" & tmpTS & "' AND ID = " & tmpctr + 1
						rsEXT.Open sqlEXT, g_strCONN, 1, 3
						If Not rsEXT.EOF Then
							If tmpXHRS <> 0 Then
								rsEXT("misc_notes") = Request("Mnotes1" & i)
								rsEXT("actcode") = Request("com1" & i)
								'rsEXT("timestamp") = tmpTimeStamp
								rsEXT("EXT") = True
								If Not IsNumeric(Request("hmonx1" & i)) Then
									rsEXT("mon") = 0
								Else
									rsEXT("mon") = Request("hmonx1" & i)
								End If
								If Not IsNumeric(Request("htuex1" & i)) Then
									rsEXT("tue") = 0
								Else
									rsEXT("tue") = Request("htuex1" & i)
								End If
								If Not IsNumeric(Request("hwedx1" & i)) Then
									rsEXT("wed") = 0
								Else
									rsEXT("wed") = Request("hwedx1" & i)
								End If
								If Not IsNumeric(Request("hthux1" & i)) Then
									rsEXT("thu") = 0
								Else
									rsEXT("thu") = Request("hthux1" & i)
								End If
								If Not IsNumeric(Request("hfrix1" & i)) Then
									rsEXT("fri") = 0
								Else
									rsEXT("fri") = Request("hfrix1" & i)
								End If
								If Not IsNumeric(Request("hsatx1" & i)) Then
									rsEXT("sat") = 0
								Else
									rsEXT("sat") = Request("hsatx1" & i)
								End If
								If Not IsNumeric(Request("hsunx1" & i)) Then
									rsEXT("sun") = 0
								Else
									rsEXT("sun") = Request("hsunx1" & i)
								End If		
							Else
								rsEXT("mon") = 0
								rsEXT("tue") = 0
								rsEXT("wed") = 0
								rsEXT("thu") = 0
								rsEXT("fri") = 0
								rsEXT("sat") = 0
								rsEXT("sun") = 0
								rsEXT("misc_notes") = Request("Mnotes1" & i)
								rsEXT("actcode") = Request("com1" & i)
								'rsEXT("timestamp") = tmpTimeStamp
							End If	
							rsEXT.Update	
							rsEXT.Close
							Set rsEXT = Nothing
						End If
					End If
				End If
			End If	
			tblEMP.Close
			Set tblEMP = Nothing
		Next
	End If
	'NEW HOURS ROW
	'GET NEW EXTENDED HOURS
	tmpXHRSN = 0
	tmpXHRSN = Z_CZero(Request("hmonX1")) + Z_CZero(Request("htueX1")) + Z_CZero(Request("hwedX1")) + _
			Z_CZero(Request("hthuX1")) + Z_CZero(Request("hfriX1")) + Z_CZero(Request("hsatX1")) + _
			Z_CZero(Request("hsunX1"))
	'GET NEW HOURS
	tmpHRSN = 0
	tmpHRSN = Z_CZero(Request("hmon1")) + Z_CZero(Request("htue1")) + Z_CZero(Request("hwed1")) + _
			Z_CZero(Request("hthu1")) + Z_CZero(Request("hfri1")) + Z_CZero(Request("hsat1")) + _
			Z_CZero(Request("hsun1"))
	If Request("hdept1") <> ""  Then
		'''''''ERROR CHECK''''''''
		If tmpXHRSN = 0 And tmpHRSN = 0 Then
			xHRS = Z_DoEncrypt(Request("hmon1") & "|" & Request("htue1") & "|" & Request("hwed1") & "|" & _
				Request("hthu1") & "|" & Request("hfri1") & "|" & Request("hsat1") & "|" & _
				Request("hsun1") & "|" & Request("Mnotes1") & "|" & Request("hmonX1") & "|" & Request("htueX1") & "|" & Request("hwedX1") & "|" & _
				Request("hthuX1") & "|" & Request("hfriX1") & "|" & Request("hsatX1") & "|" & _
				Request("hsunX1") & "|" & Request("com1") & "|" & Request("fon1"))
			Session("MSG") = Session("MSG") & "<br>Total hours cannot be equal to 0 in new consumer in week 1."
		Else
			If (tmpHRSN <> 0 Or tmpXHRSN <> 0) And (Request("Mnotes1") = "" Or Request("com1") = "" Or Request("fon1") = "") Then
				xHRS = Z_DoEncrypt(Request("hmon1") & "|" & Request("htue1") & "|" & Request("hwed1") & "|" & _
				Request("hthu1") & "|" & Request("hfri1") & "|" & Request("hsat1") & "|" & _
				Request("hsun1") & "|" & Request("Mnotes1") & "|" & Request("hmonX1") & "|" & Request("htueX1") & "|" & Request("hwedX1") & "|" & _
				Request("hthuX1") & "|" & Request("hfriX1") & "|" & Request("hsatX1") & "|" & _
				Request("hsunX1") & "|" & Request("com1") & "|" & Request("fon1"))
				Set rsConNam = Server.CreateObject("ADODB.RecordSet")
				sqlConNam = "SELECT * FROM Consumer_t WHERE Medicaid_Number = '" &  Request("hdept1") & "' "
				rsConNam.Open sqlConNam, g_strCONN, 1, 3
				ConNam = Request("hdept1")
				If Not rsConNam.EOF Then
					ConNam = rsConNam("LName") & ", " & rsConNam("FName")
				End If 	
				rsConNam.Close
				Set rsConNam = Nothing
				Session("MSG") = Session("MSG") & "<br>Activity code(s)/Notes/Phone Num are required in week 1 for consumer " & ConNam & "."
			Else
				MaxHrs = "MaxHrs"
				If Instr(Request("Mnotes1"), "80,") > 0 Then 
					MaxHrs = "vahmhrs"
				ElseIf Instr(Request("Mnotes1"), "82,") > 0 Then 
					MaxHrs = "vahahrs"
				End If
				'GET MAX HOURS AND MILEAGE CAP
				Set tblCon = Server.CreateObject("ADODB.RecordSet")
				sqlCon = "SELECT " & MaxHrs & ", milecap FROM Consumer_t WHERE medicaid_number = '" & Request("hdept1") & "' "
				tblCon.Open sqlCon, g_strCONN, 1, 3
				If Not tblCon.EOF Then
					tmpMax = Z_CZero(tblCon(MaxHrs))
					tmpMile = Z_CZero(tblCon("milecap"))
				End If
				tblCon.Close
				Set tblCon = Nothing
				'GET EXISTING HOURS
				sqlChk = "SELECT * FROM tSHEETS_T WHERE date = '" & Request("1day") & "' AND client = '" & _
				Request("hdept1") & "' AND Ext = 0"
				If Instr(Request("Mnotes1"), "80,") > 0 Then 
					sqlChk = sqlChk & " AND misc_notes = '80,'"
				ElseIf Instr(Request("Mnotes1"), "82,") > 0 Then 
					sqlChk = sqlChk & " AND misc_notes = '82,'"
				End If 
				Set rsChk = Server.CreateObject("ADODB.RecordSet")
				rsChk.Open sqlChk, g_strCONN, 3, 1
				If Not rsChk.EOF Then
					HrsCon = 0
					MileCon = 0
					Do Until rsChk.EOF
						HrsCon = HrsCon + Z_CZero(rsChk("mon")) + Z_CZero(rsChk("tue")) + Z_CZero(rsChk("wed")) + Z_CZero(rsChk("thu")) + Z_CZero(rsChk("fri")) + _
							Z_CZero(rsChk("sat")) + Z_CZero(rsChk("sun"))
						MileCon = MileCon +  Z_CZero(rsChk("mile"))
						rsChk.MoveNExt
					Loop
				End If	
				rsChk.Close
				Set rsChk = Nothing
				'GET HOURS AND MILEAGE
				tmpHRS = 0
				tmpHRS = tmpHRSN + HrsCon
				tmpMileTot = MileCon + Z_CZero(Request("txtmile"))
				If tmpHRS > tmpMax Then
					Set rsConNam = Server.CreateObject("ADODB.RecordSet")
					sqlConNam = "SELECT * FROM Consumer_t WHERE Medicaid_Number = '" &  Request("hdept1") & "' "
					rsConNam.Open sqlConNam, g_strCONN, 1, 3
						ConNam = Request("hdept1")
					If Not rsConNam.EOF Then
						ConNam = rsConNam("LName") & ", " & rsConNam("FName")
					End If 	
					rsConNam.Close
					Set rsConNam = Nothing
					Session("MSG") = Session("MSG") & "<br>Total hours for " & ConNam & " is over the allowed hours in week 1. This row will be red-flagged on History report and Process Items report."
					maxHRS1 = "true"
					'SET MAX = TRUE ON DB FOR ALL WORKERS
					Set rsMax = Server.CreateObject("ADODB.RecordSet")
					sqlMax = "SELECT * FROM tSHEETS_T WHERE date = '" & Request("1day") & "' AND client = '" & _
						Request("hdept1") & "' AND Ext = 0 AND MAX = 0"
					If Instr(Request("Mnotes1"), "80,") > 0 Then 
						sqlMax = sqlMax & " AND misc_notes = '80,'"
					ElseIf Instr(Request("Mnotes1"), "82,") > 0 Then 
						sqlMax = sqlMax & " AND misc_notes = '82,'"
					End If
					rsMax.Open sqlMax, g_strCONN, 1, 3
					If Not rsMax.EOF Then
						Do Until rsMAX.EOF
							rsMAX("MAX") = True
							rsMAX.Update
							rsMAX.MoveNext
						Loop
					End If
					rsMAX.Close
					Set rsMax = Nothing
				Else
					maxHRS1 = "false"
					'SET MAX = FALSE ON DB FOR ALL WORKERS
					Set rsMax = Server.CreateObject("ADODB.RecordSet")
					sqlMax = "SELECT * FROM tSHEETS_T WHERE date = '" & Request("1day") & "' AND client = '" & _
						Request("hdept1") & "' AND Ext = 0 AND MAX = 1"
					If Instr(Request("Mnotes1"), "80,") > 0 Then 
						sqlMax = sqlMax & " AND misc_notes = '80,'"
					ElseIf Instr(Request("Mnotes1"), "82,") > 0 Then 
						sqlMax = sqlMax & " AND misc_notes = '82,'"
					End If 
					rsMax.Open sqlMax, g_strCONN, 1, 3
					If Not rsMax.EOF Then
						Do Until rsMAX.EOF
							rsMAX("MAX") = False
							rsMAX.Update
							rsMAX.MoveNext
						Loop
					End If
					rsMAX.Close
					Set rsMax = Nothing
				End If
				''''
				If tmpMileTot > tmpMile Then
					Set rsConNam = Server.CreateObject("ADODB.RecordSet")
					sqlConNam = "SELECT * FROM Consumer_t WHERE Medicaid_Number = '" &  Request("hdept1") & "' "
					rsConNam.Open sqlConNam, g_strCONN, 1, 3
						ConNam = Request("hdept1")
					If Not rsConNam.EOF Then
						ConNam = rsConNam("LName") & ", " & rsConNam("FName")
					End If 	
					rsConNam.Close
					Set rsConNam = Nothing
					Session("MSG") = Session("MSG") & "<br>Total Mileage for " &  ConNam & " is over the allowed mileage in week 1. This row will be flagged on History report and Process Items report."
					maxMile = "True"
					'SET milecap = TRUE ON DB FOR ALL WORKERS
					Set rsMax = Server.CreateObject("ADODB.RecordSet")
					sqlMax = "SELECT * FROM tSHEETS_T WHERE date = '" & Request("1day") & "' AND client = '" & _
						Request("hdept1") & "' AND Ext = 0 AND emp_id <> '" & Session("idemp") & "'"
					rsMax.Open sqlMax, g_strCONN, 1, 3
					If Not rsMax.EOF Then
						Do Until rsMAX.EOF
							rsMAX("milecap") = True
							rsMAX.Update
							rsMAX.MoveNext
						Loop
					End If
					rsMAX.Close
					Set rsMax = Nothing
				Else
					maxMile = "False"
					'SET milecap = FALSE ON DB FOR ALL WORKERS
					Set rsMax = Server.CreateObject("ADODB.RecordSet")
					sqlMax = "SELECT * FROM tSHEETS_T WHERE date = '" & Request("1day") & "' AND client = '" & _
						Request("hdept1") & "' AND Ext = 0 AND emp_id <> '" & Session("idemp") & "'"
					rsMax.Open sqlMax, g_strCONN, 1, 3
					If Not rsMax.EOF Then
						Do Until rsMAX.EOF
							rsMAX("milecap") = False
							rsMAX.Update
							rsMAX.MoveNext
						Loop
					End If
					rsMAX.Close
					Set rsMax = Nothing
				End If
				'SAVE HOURS
				Set tblEMP = Server.CreateObject("ADODB.RecordSet")
				strSQL = "SELECT * FROM tsheets_T"
				tblEMP.Open strSQL, g_strCONN, 1, 3
				tblEMP.AddNew
				tblEMP("MAX") = maxHRS1
				tblEMP("milecap") = maxMile
				tblEMP("client") = Request("hdept1")
				If Not IsNumeric(Request("hmon1")) Then
					tblEMP("mon") = 0
				Else
					tblEMP("mon") = Request("hmon1")
				End If
				If Not IsNumeric(Request("htue1")) Then
					tblEMP("tue") = 0
				Else
					tblEMP("tue") = Request("htue1")
				End If
				If Not IsNumeric(Request("hwed1")) Then
					tblEMP("wed") = 0
				Else
					tblEMP("wed") = Request("hwed1")
				End If
				If Not IsNumeric(Request("hthu1")) Then
					tblEMP("thu") = 0
				Else
					tblEMP("thu") = Request("hthu1")
				End If
				If Not IsNumeric(Request("hfri1")) Then
					tblEMP("fri") = 0
				Else
					tblEMP("fri") = Request("hfri1")
				End If
				If Not IsNumeric(Request("hsat1")) Then
					tblEMP("sat") = 0
				Else
					tblEMP("sat") = Request("hsat1")
				End If
				If Not IsNumeric(Request("hsun1")) Then
					tblEMP("sun") = 0
				Else
					tblEMP("sun") = Request("hsun1")
				End If	
				'If Not IsNumeric(Request("txtmile")) Then
				'	tblEMP("mile") = 0
				'Else
					tblEMP("mile") = Z_Czero(Request("txtmile"))
				'End If	
				If Not IsNumeric(Request("txtamile")) Then
					tblEMP("amile") = 0
				Else
					tblEMP("amile") = Request("txtamile")
				End If	
				tblEMP("date") = Request("1day")
				tblEMP("emp_id") = Session("idemp") 
				tblEMP("misc_notes") = Request("Mnotes1")
				tblEMP("actcode") = Request("com1")
				tblEMP("callerID") = Request("fon1")
				tblEMP("author") = session("UserID")
				tblEMP("timestamp") = tmpTimeStamp
				tblEMP.Update
				tblEMP.Close
				Set tblEMP = Nothing
				'SAVE EXTENDED HOURS
				Set tblEXT = Server.CreateObject("ADODB.Recordset")
				strSQL = "SELECT * FROM [tsheets_t]"
				tblEXT.Open strSQL, g_strCONN, 1, 3
				tblEXT.addnew
				tblEXT("EXT") = True
				tblEXT("client") = Request("hdept1")
				If Not IsNumeric(Request("hmonX1")) Then
					tblEXT("mon") = 0
				Else
					tblEXT("mon") = Cdbl(Request("hmonX1"))
				End If
				If Not IsNumeric(Request("htueX1")) Then
					tblEXT("tue") = 0
				Else
					tblEXT("tue") = Request("htueX1")
				End If
				If Not IsNumeric(Request("hwedX1")) Then
					tblEXT("wed") = 0
				Else
					tblEXT("wed") = Request("hwedX1")
				End If
				If Not IsNumeric(Request("hthuX1")) Then
					tblEXT("thu") = 0
				Else
					tblEXT("thu") = Request("hthuX1")
				End If
				If Not IsNumeric(Request("hfriX1")) Then
					tblEXT("fri") = 0
				Else
					tblEXT("fri") = Request("hfriX1")
				End If
				If Not IsNumeric(Request("hsatX1")) Then
					tblEXT("sat") = 0
				Else
					tblEXT("sat") = Request("hsatX1")
				End If
				If Not IsNumeric(Request("hsunX1")) Then
					tblEXT("sun") = 0
				Else
					tblEXT("sun") = Request("hsunX1")
				End If	
				tblEXT("date") = Request("1day")
				tblEXT("emp_id") = Session("idemp") 
				tblEXT("author") = session("UserID")
				tblEXT("misc_notes") = Request("Mnotes1")
				tblEXT("actcode") = Request("com1")
				tblEXT("callerID") = Request("fon1")
				tblEXT("timestamp") = tmpTimeStamp
				tblEXT.UPDATE
				tblEXT.Close
				Set tblEXT = Nothing
			End If
		End If	
	Else
		If tmpXHRSN <> 0 Or tmpHRSN <> 0 Then
			xHRS = Z_DoEncrypt(Request("hmon1") & "|" & Request("htue1") & "|" & Request("hwed1") & "|" & _
				Request("hthu1") & "|" & Request("hfri1") & "|" & Request("hsat1") & "|" & _
				Request("hsun1") & "|" & Request("Mnotes1") & "|" & Request("hmonX1") & "|" & Request("htueX1") & "|" & Request("hwedX1") & "|" & _
				Request("hthuX1") & "|" & Request("hfriX1") & "|" & Request("hsatX1") & "|" & _
				Request("hsunX1"))
			Session("MSG") = Session("MSG") & "<br>Please select a new consumer in week 1." 
		End If
	End If
	'session("MSG") = "PTO1: " & Request("txtPTO1")
	'If Z_CZero(Request("txtPTO1")) <> 0 Then
		'SAVE PTO
		Set rsPTO = Server.CreateObject("ADODB.RecordSet")
		sqlPTO = "SELECT * FROM W_PTO_T WHERE WorkerID = '" & Session("idemp")  & "' AND date = '" & Request("1day") & "'"
		rsPTO.Open sqlPTO, g_strCONN, 1, 3
		If rsPTO.EOF Then
			If Z_CZero(Request("txtPTO1")) <> 0 Then
				rsPTO.AddNew
				rsPTO("WorkerID") = Session("idemp")
				rsPTO("date") = Request("1day")
				rsPTO("PTO") = Z_CZero(Request("txtPTO1"))
				rsPTO.Update
			End If
		Else
			If IsNull(rsPTO("procitem")) Then
				'rsPTO("WorkerID") = Session("idemp")
				'rsPTO("date") = Request("1day")
				rsPTO("PTO") = Z_CZero(Request("txtPTO1"))
				rsPTO.Update
			Else
				rsPTO("PTO") = Z_CZero(Request("txtPTO1"))
				rsPTO.Update
			End If
		End If
		rsPTO.Close
		Set rsPTO = Nothing
	'End If
'End If	
'SAVE ROW FOR WEEK 2
'If Request("paychk2") = "" AND Request("medchk2") = "" Then
	ctrI = Request("count2")
	If Request("count2") <> 0 Then
		'EXISTING HOURS ROW
		For i = 0 to ctrI 
			Set tblEMP = Server.CreateObject("ADODB.Recordset")
			strSQL = "SELECT * FROM [tsheets_t]"
			tblEMP.Open strSQL, g_strCONN, 1, 3
			tblEMP.Movefirst
			tmpctr = Request("chkS" & i)
			If tmpctr <> "" Then
				strTmp = "ID='" & tmpctr & "' "
				tblEMP.Find(strTmp)
				If Not tblEMP.EOF Then
					'''''''ERROR CHECK''''''''
					'GET TOTAL EXTENDED HOURS
					tmpXHRS = 0
					tmpXHRS = Z_CZero(Request("hsunx2" & i)) + Z_CZero(Request("hmonx2" & i)) +Z_CZero(Request("htuex2" & i)) +Z_CZero(Request("hwedx2" & i)) + _
						Z_CZero(Request("hthux2" & i)) + Z_CZero(Request("hfrix2" & i)) + Z_CZero(Request("hsatx2" & i))
					'GET TOTAL HOURS
					tmpTotHRS = 0
					tmpTotHRS = Z_CZero(Request("hsunx2" & i)) + Z_CZero(Request("hmonx2" & i)) +Z_CZero(Request("htuex2" & i)) +Z_CZero(Request("hwedx2" & i)) + _
						Z_CZero(Request("hthux2" & i)) + Z_CZero(Request("hfrix2" & i)) + Z_CZero(Request("hsatx2" & i)) + Z_CZero(Request("hsun2" & i)) + _
						Z_CZero(Request("hmon2" & i)) + Z_CZero(Request("htue2" & i)) + Z_CZero(Request("hwed2" & i)) + _
						Z_CZero(Request("hthu2" & i)) + Z_CZero(Request("hfri2" & i)) + Z_CZero(Request("hsat2" & i))
					If tmpXHRS <> 0 And Request("Mnotes2" & i) = "" Then
						Session("MSG") = Session("MSG") & "<br>Notes are required if there are extEnded hours in week 2 for Consumer " & Request("hdept2" & i) & "." 
					ElseIf tmpTotHRS = 0 Then
						Session("MSG") = Session("MSG") & "<br>Total hours for " & Request("hdept2" & i) & " will become zero(0) in week 2. Please just delete the row."
					Else
						MaxHrs = "MaxHrs"
						If Instr(Request("Mnotes2" & i), "80,") > 0 Then 
							MaxHrs = "vahmhrs"
						ElseIf Instr(Request("Mnotes2" & i), "82,") > 0 Then 
							MaxHrs = "vahahrs"
						End If
						'GET MAX HOURS AND MILEAGE CAP
						Set tblCon = Server.CreateObject("ADODB.RecordSet")
						sqlCon = "SELECT " & MaxHrs & ", milecap FROM Consumer_t WHERE medicaid_number = '" & Request("conid2" & i) & "' "
						tblCon.Open sqlCon, g_strCONN, 1, 3
						If Not tblCon.EOF Then
							tmpMax = Z_CZero(tblCon(MaxHrs))
							tmpMile = Z_CZero(tblCon("milecap"))
						End If
						tblCon.Close
						Set tblCon = Nothing
						'GET EXISTING HOURS AND MILEAGE
						Set rsChk = Server.CreateObject("ADODB.RecordSet")
						sqlChk = "SELECT * FROM tSHEETS_T WHERE date = '" & Request("2day") & "' AND client = '" & _
							Request("conid2" & i) & "' AND Ext = 0"
						If Instr(Request("Mnotes2" & i), "80,") > 0 Then 
							sqlChk = sqlChk & " AND misc_notes = '80,'"
						ElseIf Instr(Request("Mnotes2" & i), "82,") > 0 Then 
							sqlChk = sqlChk & " AND misc_notes = '82,'"
						End If 
						rsChk.Open sqlChk, g_strCONN, 3, 1
						If Not rsChk.EOF Then
							HrsCon = 0
							MileCon = 0
							Do Until rsChk.EOF
								HrsCon = HrsCon + Z_CZero(rsChk("mon")) + Z_CZero(rsChk("tue")) + Z_CZero(rsChk("wed")) + Z_CZero(rsChk("thu")) + Z_CZero(rsChk("fri")) + _
									Z_CZero(rsChk("sat")) + Z_CZero(rsChk("sun"))
								MileCon = MileCon +  Z_CZero(rsChk("mile"))
								rsChk.MoveNext
							Loop
						End If	
						rsChk.Close
						Set rsChk = Nothing
						'GET HOURS AND MILEAGE
						tmpHRS = 0
						tmpHRS = HrsCon + Z_CZero(Request("hsun2" & i)) + _
							Z_CZero(Request("hmon2" & i)) + Z_CZero(Request("htue2" & i)) + Z_CZero(Request("hwed2" & i)) + _
							Z_CZero(Request("hthu2" & i)) + Z_CZero(Request("hfri2" & i)) + Z_CZero(Request("hsat2" & i))
						tmpMileTot = MileCon + Z_CZero(Request("txt2mile" & i))
						If tmpHRS > tmpMax Then
							Session("MSG") = Session("MSG") & "<br>Total hours for " &  Request("hdept2" & i) & " is over the allowed hours in week 2. This row will be red-flagged on History report and Process Items report."
							tblEMP("MAX") = True
							'SET MAX = TRUE ON DB FOR ALL WORKERS
							Set rsMax = Server.CreateObject("ADODB.RecordSet")
							sqlMax = "SELECT * FROM tSHEETS_T WHERE date = '" & Request("2day") & "' AND client = '" & _
								Request("conid2" & i) & "' AND Ext = 0 AND MAX = 0"
							If Instr(Request("Mnotes2" & i), "80,") > 0 Then 
								sqlMax = sqlMax & " AND misc_notes = '80,'"
							ElseIf Instr(Request("Mnotes2" & i), "82,") > 0 Then 
								sqlMax = sqlMax & " AND misc_notes = '82,'"
							End If 
							rsMax.Open sqlMax, g_strCONN, 1, 3
							If Not rsMax.EOF Then
								Do Until rsMAX.EOF
									rsMAX("MAX") = True
									rsMAX.Update
									rsMAX.MoveNext
								Loop
							End If
							rsMAX.Close
							Set rsMax = Nothing
						Else
							tblEMP("MAX") = False
							'SET MAX = FALSE ON DB FOR ALL WORKERS
							Set rsMax = Server.CreateObject("ADODB.RecordSet")
							sqlMax = "SELECT * FROM tSHEETS_T WHERE date = '" & Request("2day") & "' AND client = '" & _
								Request("conid2" & i) & "' AND Ext = 0 AND MAX = 1"
							If Instr(Request("Mnotes2" & i), "80,") > 0 Then 
								sqlMax = sqlMax & " AND misc_notes = '80,'"
							ElseIf Instr(Request("Mnotes2" & i), "82,") > 0 Then 
								sqlMax = sqlMax & " AND misc_notes = '82,'"
							End If 
							rsMax.Open sqlMax, g_strCONN, 1, 3
							If Not rsMax.EOF Then
								Do Until rsMAX.EOF
									rsMAX("MAX") = False
									rsMAX.Update
									rsMAX.MoveNext
								Loop
							End If
							rsMAX.Close
							Set rsMax = Nothing
						End If
						''''
						If tmpMileTot > tmpMile Then
							Session("MSG") = Session("MSG") & "<br>Total Mileage for " &  Request("hdept2" & i) & " is over the allowed mileage in week 2. This row will be flagged on History report and Process Items report."
							tblEMP("milecap") = True
							'SET milecap = TRUE ON DB FOR ALL WORKERS
							Set rsMax = Server.CreateObject("ADODB.RecordSet")
							sqlMax = "SELECT * FROM tSHEETS_T WHERE date = '" & Request("2day") & "' AND client = '" & _
								Request("conid2" & i) & "' AND Ext = 0 AND emp_id <> '" & Session("idemp") & "'"
							rsMax.Open sqlMax, g_strCONN, 1, 3
							If Not rsMax.EOF Then
								Do Until rsMAX.EOF
									rsMAX("milecap") = True
									rsMAX.Update
									rsMAX.MoveNext
								Loop
							End If
							rsMAX.Close
							Set rsMax = Nothing
						Else
							tblEMP("milecap") = False
							'SET milecap = FALSE ON DB FOR ALL WORKERS
							Set rsMax = Server.CreateObject("ADODB.RecordSet")
							sqlMax = "SELECT * FROM tSHEETS_T WHERE date = '" & Request("2day") & "' AND client = '" & _
								Request("conid2" & i) & "' AND Ext = 0 AND emp_id <> '" & Session("idemp") & "'"
							rsMax.Open sqlMax, g_strCONN, 1, 3
							If Not rsMax.EOF Then
								Do Until rsMAX.EOF
									rsMAX("milecap") = False
									rsMAX.Update
									rsMAX.MoveNext
								Loop
							End If
							rsMAX.Close
							Set rsMax = Nothing
						End If
						'SAVE HOURS
						If Not IsNumeric(Request("hmon2" & i)) Then
							tblEMP("mon") = 0
						Else
							tblEMP("mon") = Request("hmon2" & i)
						End If
						If Not IsNumeric(Request("htue2" & i)) Then
							tblEMP("tue") = 0
						Else
							tblEMP("tue") = Request("htue2" & i)
						End If
						If Not IsNumeric(Request("hwed2" & i)) Then
							tblEMP("wed") = 0
						Else
							tblEMP("wed") = Request("hwed2" & i)
						End If
						If Not IsNumeric(Request("hthu2" & i)) Then
							tblEMP("thu") = 0
						Else
							tblEMP("thu") = Request("hthu2" & i)
						End If
						If Not IsNumeric(Request("hfri2" & i)) Then
							tblEMP("fri") = 0
						Else
							tblEMP("fri") = Request("hfri2" & i)
						End If
						If Not IsNumeric(Request("hsat2" & i)) Then
							tblEMP("sat") = 0
						Else
							tblEMP("sat") = Request("hsat2" & i)
						End If
						If Not IsNumeric(Request("hsun2" & i)) Then
							tblEMP("sun") = 0
						Else
							tblEMP("sun") = Request("hsun2" & i)
						End If	
						'If Not IsNumeric(Request("txt2mile" & i)) Then
						'	tblEMP("mile") = 0
						'Else
							tblEMP("mile") = Z_Czero(Request("txt2mile" & i))
						'End If
						If Not IsNumeric(Request("txt2amile" & i)) Then
							tblEMP("amile") = 0
						Else
							tblEMP("amile") = Request("txt2amile" & i)
						End If
						tblEMP("misc_notes") = Request("Mnotes2" & i)
						tblEMP("actcode") = Request("com2" & i)
						tblEMP("author") = session("UserID")
						tmpTS = tblEMP("timestamp")
						tblEMP.Update
						'SAVE EXTENDED HOURS
						Set rsEXT = Server.CreateObject("ADODB.RecordSet")
						sqlEXT = "SELECT * FROM [tsheets_t] WHERE " & _
							"[emp_id] = '" & tblEMP("emp_ID") & "' AND " & _
							"[date] = '" & tblEMP("date") &  "' AND EXT = 1 AND client = '" & tblEMP("Client") & _
							"' AND timestamp = '" & tmpTS & "' AND ID = " & tmpctr + 1
						rsEXT.Open sqlEXT, g_strCONN, 1, 3
						If Not rsEXT.EOF Then
							If tmpXHRS <> 0 Then
								rsEXT("misc_notes") = Request("Mnotes2" & i)
								rsEXT("actcode") = Request("com2" & i)
								'rsEXT("timestamp") = tmpTimeStamp
								rsEXT("EXT") = True
								If Not IsNumeric(Request("hmonx2" & i)) Then
									rsEXT("mon") = 0
								Else
									rsEXT("mon") = Request("hmonx2" & i)
								End If
								If Not IsNumeric(Request("htuex2" & i)) Then
									rsEXT("tue") = 0
								Else
									rsEXT("tue") = Request("htuex2" & i)
								End If
								If Not IsNumeric(Request("hwedx2" & i)) Then
									rsEXT("wed") = 0
								Else
									rsEXT("wed") = Request("hwedx2" & i)
								End If
								If Not IsNumeric(Request("hthux2" & i)) Then
									rsEXT("thu") = 0
								Else
									rsEXT("thu") = Request("hthux2" & i)
								End If
								If Not IsNumeric(Request("hfrix2" & i)) Then
									rsEXT("fri") = 0
								Else
									rsEXT("fri") = Request("hfrix2" & i)
								End If
								If Not IsNumeric(Request("hsatx2" & i)) Then
									rsEXT("sat") = 0
								Else
									rsEXT("sat") = Request("hsatx2" & i)
								End If
								If Not IsNumeric(Request("hsunx2" & i)) Then
									rsEXT("sun") = 0
								Else
									rsEXT("sun") = Request("hsunx2" & i)
								End If		
							Else
								rsEXT("mon") = 0
								rsEXT("tue") = 0
								rsEXT("wed") = 0
								rsEXT("thu") = 0
								rsEXT("fri") = 0
								rsEXT("sat") = 0
								rsEXT("sun") = 0
								rsEXT("misc_notes") = Request("Mnotes2" & i)
								rsEXT("actcode") = Request("com2" & i)
								'rsEXT("timestamp") = tmpTimeStamp
							End If	
							rsEXT.Update	
							rsEXT.Close
							Set rsEXT = Nothing
						End If
					End If
				End If
			End If	
			tblEMP.Close
			Set tblEMP = Nothing
		Next
	End If
	'NEW HOURS ROW
	'GET NEW EXTENDED HOURS
	tmpXHRSN = 0
	tmpXHRSN = Z_CZero(Request("hmonX2")) + Z_CZero(Request("htueX2")) + Z_CZero(Request("hwedX2")) + _
			Z_CZero(Request("hthuX2")) + Z_CZero(Request("hfriX2")) + Z_CZero(Request("hsatX2")) + _
			Z_CZero(Request("hsunX2"))
	'GET NEW HOURS
	tmpHRSN = 0
	tmpHRSN = Z_CZero(Request("hmon2")) + Z_CZero(Request("htue2")) + Z_CZero(Request("hwed2")) + _
			Z_CZero(Request("hthu2")) + Z_CZero(Request("hfri2")) + Z_CZero(Request("hsat2")) + _
			Z_CZero(Request("hsun2"))
	If Request("hdept2") <> ""  Then
		'''''''ERROR CHECK''''''''
		If tmpXHRSN = 0 And tmpHRSN = 0 Then
			xHRS2 = Z_DoEncrypt(Request("hmon2") & "|" & Request("htue2") & "|" & Request("hwed2") & "|" & _
				Request("hthu2") & "|" & Request("hfri2") & "|" & Request("hsat2") & "|" & _
				Request("hsun2") & "|" & Request("Mnotes2") & "|" & Request("hmonX2") & "|" & Request("htueX2") & "|" & Request("hwedX2") & "|" & _
				Request("hthuX2") & "|" & Request("hfriX2") & "|" & Request("hsatX2") & "|" & _
				Request("hsunX2") & "|" & Request("com2") & "|" & Request("fon2"))
			Session("MSG") = Session("MSG") & "<br>Total hours cannot be equal to 0 in new consumer in week 2."
		Else
			If (tmpHRSN <> 0 Or tmpXHRSN <> 0) And (Request("Mnotes2") = "" Or Request("com2") = "" Or Request("fon2") = "") Then
				xHRS2 = Z_DoEncrypt(Request("hmon2") & "|" & Request("htue2") & "|" & Request("hwed2") & "|" & _
				Request("hthu2") & "|" & Request("hfri2") & "|" & Request("hsat2") & "|" & _
				Request("hsun2") & "|" & Request("Mnotes2") & "|" & Request("hmonX2") & "|" & Request("htueX2") & "|" & Request("hwedX2") & "|" & _
				Request("hthuX2") & "|" & Request("hfriX2") & "|" & Request("hsatX2") & "|" & _
				Request("hsunX2") & "|" & Request("com2") & "|" & Request("fon2"))
				Set rsConNam = Server.CreateObject("ADODB.RecordSet")
				sqlConNam = "SELECT * FROM Consumer_t WHERE Medicaid_Number = '" &  Request("hdept2") & "' "
				rsConNam.Open sqlConNam, g_strCONN, 1, 3
				ConNam = Request("hdept2")
				If Not rsConNam.EOF Then
					ConNam = rsConNam("LName") & ", " & rsConNam("FName")
				End If 	
				rsConNam.Close
				Set rsConNam = Nothing
				Session("MSG") = Session("MSG") & "<br>Activity code(s)/Notes/Phone Num are required in week 2 for consumer " & ConNam & "."
			Else
				MaxHrs = "MaxHrs"
				If Instr(Request("Mnotes2"), "80,") > 0 Then 
					MaxHrs = "vahmhrs"
				ElseIf Instr(Request("Mnotes2"), "82,") > 0 Then 
					MaxHrs = "vahahrs"
				End If
				'GET MAX HOURS AND MILEAGE CAP
				Set tblCon = Server.CreateObject("ADODB.RecordSet")
				sqlCon = "SELECT " & MaxHrs & ", milecap FROM Consumer_t WHERE medicaid_number = '" & Request("hdept2") & "' "
				tblCon.Open sqlCon, g_strCONN, 1, 3
				If Not tblCon.EOF Then
					tmpMax = Z_CZero(tblCon(MaxHrs))
					tmpMile = Z_CZero(tblCon("milecap"))
				End If
				tblCon.Close
				Set tblCon = Nothing
				response.write "max hours: " & tmpMax & "<br>"
				'GET EXISTING HOURS
				sqlChk = "SELECT * FROM tSHEETS_T WHERE date = '" & Request("2day") & "' AND client = '" & _
					Request("hdept2") & "' AND Ext = 0"
				If Instr(Request("Mnotes2"), "80,") > 0 Then 
					sqlChk = sqlChk & " AND misc_notes = '80,'"
				ElseIf Instr(Request("Mnotes2"), "82,") > 0 Then 
					sqlChk = sqlChk & " AND misc_notes = '82,'"
				End If 
				response.write sqlChk & "<br>"
				Set rsChk = Server.CreateObject("ADODB.RecordSet")
				rsChk.Open sqlChk, g_strCONN, 3, 1
				If Not rsChk.EOF Then
					HrsCon = 0
					MileCon = 0
					Do Until rsChk.EOF
						HrsCon = HrsCon + Z_CZero(rsChk("mon")) + Z_CZero(rsChk("tue")) + Z_CZero(rsChk("wed")) + Z_CZero(rsChk("thu")) + Z_CZero(rsChk("fri")) + _
							Z_CZero(rsChk("sat")) + Z_CZero(rsChk("sun"))
						MileCon = MileCon +  Z_CZero(rsChk("mile"))
						rsChk.MoveNExt
					Loop
				End If	
				rsChk.Close
				Set rsChk = Nothing
				'GET HOURS AND MILEAGE
				tmpHRS = 0
				tmpHRS = tmpHRSN + HrsCon
				response.write "HRS: " & tmphrs & "<br>"
				tmpMileTot = MileCon + Z_CZero(Request("txt2mile"))
				If tmpHRS > tmpMax Then
					Set rsConNam = Server.CreateObject("ADODB.RecordSet")
					sqlConNam = "SELECT * FROM Consumer_t WHERE Medicaid_Number = '" &  Request("hdept2") & "' "
					rsConNam.Open sqlConNam, g_strCONN, 1, 3
						ConNam = Request("hdept2")
					If Not rsConNam.EOF Then
						ConNam = rsConNam("LName") & ", " & rsConNam("FName")
					End If 	
					rsConNam.Close
					Set rsConNam = Nothing
					Session("MSG") = Session("MSG") & "<br>Total hours for " & ConNam & " is over the allowed hours in week 2. This row will be red-flagged on History report and Process Items report."
					maxHRS1 = "true"
					'SET MAX = TRUE ON DB FOR ALL WORKERS
					Set rsMax = Server.CreateObject("ADODB.RecordSet")
					sqlMax = "SELECT * FROM tSHEETS_T WHERE date = '" & Request("2day") & "' AND client = '" & _
						Request("hdept2") & "' AND Ext = 0 AND MAX = 0"
					If Instr(Request("Mnotes2"), "80,") > 0 Then 
						sqlMax = sqlMax & " AND misc_notes = '80,'"
					ElseIf Instr(Request("Mnotes2"), "82,") > 0 Then 
						sqlMax = sqlMax & " AND misc_notes = '82,'"
					End If
					rsMax.Open sqlMax, g_strCONN, 1, 3
					If Not rsMax.EOF Then
						Do Until rsMAX.EOF
							rsMAX("MAX") = True
							rsMAX.Update
							rsMAX.MoveNext
						Loop
					End If
					rsMAX.Close
					Set rsMax = Nothing
				Else
					maxHRS1 = "false"
					Set rsMax = Server.CreateObject("ADODB.RecordSet")
					sqlMax = "SELECT * FROM tSHEETS_T WHERE date = '" & Request("2day") & "' AND client = '" & _
						Request("hdept2") & "' AND Ext = 0 AND MAX = 1"
					If Instr(Request("Mnotes2"), "80,") > 0 Then 
						sqlMax = sqlMax & " AND misc_notes = '80,'"
					ElseIf Instr(Request("Mnotes2"), "82,") > 0 Then 
						sqlMax = sqlMax & " AND misc_notes = '82,'"
					End If 
					rsMax.Open sqlMax, g_strCONN, 1, 3
					If Not rsMax.EOF Then
						Do Until rsMAX.EOF
							rsMAX("MAX") = False
							rsMAX.Update
							rsMAX.MoveNext
						Loop
					End If
					rsMAX.Close
					Set rsMax = Nothing
				End If
				''''
				If tmpMileTot > tmpMile Then
					Set rsConNam = Server.CreateObject("ADODB.RecordSet")
					sqlConNam = "SELECT * FROM Consumer_t WHERE Medicaid_Number = '" &  Request("hdept2") & "' "
					rsConNam.Open sqlConNam, g_strCONN, 1, 3
						ConNam = Request("hdept2")
					If Not rsConNam.EOF Then
						ConNam = rsConNam("LName") & ", " & rsConNam("FName")
					End If 	
					rsConNam.Close
					Set rsConNam = Nothing
					Session("MSG") = Session("MSG") & "<br>Total Mileage for " &  ConNam & " is over the allowed mileage in week 2. This row will be flagged on History report and Process Items report."
					maxMile = "True"
					'SET milecap = TRUE ON DB FOR ALL WORKERS
					Set rsMax = Server.CreateObject("ADODB.RecordSet")
					sqlMax = "SELECT * FROM tSHEETS_T WHERE date = '" & Request("2day") & "' AND client = '" & _
						Request("hdept2") & "' AND Ext = 0 AND emp_id <> '" & Session("idemp") & "'"
					rsMax.Open sqlMax, g_strCONN, 1, 3
					If Not rsMax.EOF Then
						Do Until rsMAX.EOF
							rsMAX("milecap") = True
							rsMAX.Update
							rsMAX.MoveNext
						Loop
					End If
					rsMAX.Close
					Set rsMax = Nothing
				Else
					maxMile = "False"
					'SET milecap = FALSE ON DB FOR ALL WORKERS
					Set rsMax = Server.CreateObject("ADODB.RecordSet")
					sqlMax = "SELECT * FROM tSHEETS_T WHERE date = '" & Request("2day") & "' AND client = '" & _
						Request("hdept2") & "' AND Ext = 0 AND emp_id <> '" & Session("idemp") & "'"
					rsMax.Open sqlMax, g_strCONN, 1, 3
					If Not rsMax.EOF Then
						Do Until rsMAX.EOF
							rsMAX("milecap") = False
							rsMAX.Update
							rsMAX.MoveNext
						Loop
					End If
					rsMAX.Close
					Set rsMax = Nothing
				End If
				'SAVE HOURS
				Set tblEMP = Server.CreateObject("ADODB.RecordSet")
				strSQL = "SELECT * FROM tsheets_T"
				tblEMP.Open strSQL, g_strCONN, 1, 3
				tblEMP.AddNew
				tblEMP("MAX") = maxHRS1
				tblEMP("milecap") = maxMile
				tblEMP("client") = Request("hdept2")
				If Not IsNumeric(Request("hmon2")) Then
					tblEMP("mon") = 0
				Else
					tblEMP("mon") = Request("hmon2")
				End If
				If Not IsNumeric(Request("htue2")) Then
					tblEMP("tue") = 0
				Else
					tblEMP("tue") = Request("htue2")
				End If
				If Not IsNumeric(Request("hwed2")) Then
					tblEMP("wed") = 0
				Else
					tblEMP("wed") = Request("hwed2")
				End If
				If Not IsNumeric(Request("hthu2")) Then
					tblEMP("thu") = 0
				Else
					tblEMP("thu") = Request("hthu2")
				End If
				If Not IsNumeric(Request("hfri2")) Then
					tblEMP("fri") = 0
				Else
					tblEMP("fri") = Request("hfri2")
				End If
				If Not IsNumeric(Request("hsat2")) Then
					tblEMP("sat") = 0
				Else
					tblEMP("sat") = Request("hsat2")
				End If
				If Not IsNumeric(Request("hsun2")) Then
					tblEMP("sun") = 0
				Else
					tblEMP("sun") = Request("hsun2")
				End If	
				'If Not IsNumeric(Request("txt2mile")) Then
				'	tblEMP("mile") = 0
				'Else
					tblEMP("mile") = Z_CZero(Request("txt2mile"))
				'End If
				If Not IsNumeric(Request("txt2amile")) Then
					tblEMP("amile") = 0
				Else
					tblEMP("amile") = Request("txt2amile")
				End If		
				tblEMP("date") = Request("2day")
				tblEMP("emp_id") = Session("idemp") 
				tblEMP("misc_notes") = Request("Mnotes2")
					tblEMP("actcode") = Request("com2")
				tblEMP("callerID") = Request("fon2")
				tblEMP("author") = session("UserID")
				tblEMP("timestamp") = tmpTimeStamp
				tblEMP.Update
				tblEMP.Close
				Set tblEMP = Nothing
				'SAVE EXTENDED HOURS
				Set tblEXT = Server.CreateObject("ADODB.Recordset")
				strSQL = "SELECT * FROM [tsheets_t]"
				tblEXT.Open strSQL, g_strCONN, 1, 3
				tblEXT.addnew
				tblEXT("EXT") = True
				tblEXT("client") = Request("hdept2")
				If Not IsNumeric(Request("hmonX2")) Then
					tblEXT("mon") = 0
				Else
					tblEXT("mon") = Cdbl(Request("hmonX2"))
				End If
				If Not IsNumeric(Request("htueX2")) Then
					tblEXT("tue") = 0
				Else
					tblEXT("tue") = Request("htueX2")
				End If
				If Not IsNumeric(Request("hwedX2")) Then
					tblEXT("wed") = 0
				Else
					tblEXT("wed") = Request("hwedX2")
				End If
				If Not IsNumeric(Request("hthuX2")) Then
					tblEXT("thu") = 0
				Else
					tblEXT("thu") = Request("hthuX2")
				End If
				If Not IsNumeric(Request("hfriX2")) Then
					tblEXT("fri") = 0
				Else
					tblEXT("fri") = Request("hfriX2")
				End If
				If Not IsNumeric(Request("hsatX2")) Then
					tblEXT("sat") = 0
				Else
					tblEXT("sat") = Request("hsatX2")
				End If
				If Not IsNumeric(Request("hsunX2")) Then
					tblEXT("sun") = 0
				Else
					tblEXT("sun") = Request("hsunX2")
				End If	
				tblEXT("date") = Request("2day")
				tblEXT("emp_id") = Session("idemp") 
				tblEXT("author") = session("UserID")
				tblEXT("misc_notes") = Request("Mnotes2")
				tblEXT("actcode") = Request("com2")
				tblEXT("callerID") = Request("fon2")
				tblEXT("timestamp") = tmpTimeStamp
				tblEXT.UPDATE
				tblEXT.Close
				Set tblEXT = Nothing
			End If
		End If	
	Else
		If tmpXHRSN <> 0 Or tmpHRSN <> 0 Then
			xHRS2 = Z_DoEncrypt(Request("hmon2") & "|" & Request("htue2") & "|" & Request("hwed2") & "|" & _
				Request("hthu2") & "|" & Request("hfri2") & "|" & Request("hsat2") & "|" & _
				Request("hsun2") & "|" & Request("Mnotes2") & "|" & Request("hmonX2") & "|" & Request("htueX2") & "|" & Request("hwedX2") & "|" & _
				Request("hthuX2") & "|" & Request("hfriX2") & "|" & Request("hsatX2") & "|" & _
				Request("hsunX2"))
			Session("MSG") = Session("MSG") & "<br>Please select a new consumer in week 2." 
		End If
	End If
	'If Z_CZero(Request("txtPTO2")) <> 0 Then
	'SAVE PTO
	Set rsPTO = Server.CreateObject("ADODB.RecordSet")
	sqlPTO = "SELECT * FROM W_PTO_T WHERE WorkerID = '" & Session("idemp")  & "' AND date = '" & Request("2day") & "'"
	rsPTO.Open sqlPTO, g_strCONN, 1, 3
	If rsPTO.EOF Then
		If Z_CZero(Request("txtPTO2")) <> 0 Then
			rsPTO.AddNew
			rsPTO("WorkerID") = Session("idemp")
			rsPTO("date") = Request("2day")
			rsPTO("PTO") = Z_CZero(Request("txtPTO2"))
			rsPTO.Update
		End If
	Else
		If IsNull(rsPTO("procitem")) Then
			'rsPTO("WorkerID") = Session("idemp")
			'rsPTO("date") = Request("2day")
			rsPTO("PTO") = Z_CZero(Request("txtPTO2"))
			rsPTO.Update
		Else
			rsPTO("PTO") = Z_CZero(Request("txtPTO2"))
			rsPTO.Update
		End If
	End If
	rsPTO.Close
	Set rsPTO = Nothing
'End If
'End If	
If Session("MSG") <> "" Then
	Response.Redirect "view.asp?Con=" & Request("hdept1") & "&xHRS=" & xHRS & "&Con2=" & Request("hdept2") & "&xHRS2=" & xHRS2
Else
	Response.Redirect "view.asp"
End If
%>