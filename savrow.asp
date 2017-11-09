<%@Language=VBScript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<%
DIM	lngI, tblEMP, strSQL, strTableScript, tmpID, ctrI, tmpctr
'''WEEK1
If Request("paychk") = "" AND Request("medchk") = "" Then
	ctrI = Request("count")
	If Request("count") <> 0 Then
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
					''''Get Max Hours
					Set tblCon = Server.CreateObject("ADODB.RecordSet")
					sqlCon = "SELECT * FROM Consumer_t WHERE medicaid_number = '" & Request("conid" & i) & "' "
					tblCon.Open sqlCon, g_strCONN, 1, 3
					If Not tblCon.EOF Then
						tmpMax = Z_CZero(tblCon("MaxHrs"))
					End If
					tblCon.Close
					Set tblCon = Nothing
					''''
					'''get existing hours
					Set rsChk = Server.CreateObject("ADODB.RecordSet")
					sqlChk = "SELECT * FROM tSHEETS_T WHERE date = #" & Request("1day") & "# AND client = '" & _
					Request("conid" & i) & "' AND emp_ID <> '" & Session("idemp") & "' AND Ext = false"
					rsChk.Open sqlChk, g_strCONN, 3, 1
					If Not rsChk.EOF Then
						HrsCon = 0
						Do Until rsChk.EOF
							HrsCon = HrsCon + Z_CZero(rsChk("mon")) + Z_CZero(rsChk("tue")) + Z_CZero(rsChk("wed")) + Z_CZero(rsChk("thu")) + Z_CZero(rsChk("fri")) + _
								Z_CZero(rsChk("sat")) + Z_CZero(rsChk("sun"))
							rsChk.MoveNExt
						Loop
					End If	
					rsChk.Close
					Set rsChk = Nothing
					'''
					tmpHrs = HrsCon + Z_CZero(Request("hmon1" & i)) + Z_CZero(Request("htue1" & i)) + Z_CZero(Request("hwed1" & i)) + _
							Z_CZero(Request("hthu1" & i)) + Z_CZero(Request("hfri1" & i)) + Z_CZero(Request("hsat1" & i)) + _
							Z_CZero(Request("hsun1" & i))
					tmpHrschk = Z_CZero(Request("hmon1" & i)) + Z_CZero(Request("htue1" & i)) + Z_CZero(Request("hwed1" & i)) + _
							Z_CZero(Request("hthu1" & i)) + Z_CZero(Request("hfri1" & i)) + Z_CZero(Request("hsat1" & i)) + _
							Z_CZero(Request("hsun1" & i)) + + Z_CZero(Request("hmonX1" & i)) + Z_CZero(Request("htueX1" & i)) + Z_CZero(Request("hwedX1" & i)) + _
							Z_CZero(Request("hthuX1" & i)) + Z_CZero(Request("hfriX1" & i)) + Z_CZero(Request("hsatX1" & i)) + _
							Z_CZero(Request("hsunX1" & i))
					If tmpHrs > tmpMax Then
						Session("MSG") = Session("MSG") & "<br>Total hours for " &  Request("hdept1" & i) & " is over the allowed hours in week 1.<br>" & _
							"If there are no errors found on this row, all inputted data will be saved."
						if Request("hmon1" & i)= "" then
							tblEMP("mon") = 0
						else
							tblEMP("mon") = Request("hmon1" & i)
						end if
						if Request("htue1" & i)= "" then
							tblEMP("tue") = 0
						else
							tblEMP("tue") = Request("htue1" & i)
						end if
						if Request("hwed1" & i) = "" then
							tblEMP("wed") = 0
						else
							tblEMP("wed") = Request("hwed1" & i)
						end if
						if Request("hthu1" & i) = "" then
							tblEMP("thu") = 0
						else
							tblEMP("thu") = Request("hthu1" & i)
						end if
						if Request("hfri1" & i) = "" then
							tblEMP("fri") = 0
						else
							tblEMP("fri") = Request("hfri1" & i)
						end if
						if Request("hsat1" & i) = "" then
							tblEMP("sat") = 0
						else
							tblEMP("sat") = Request("hsat1" & i)
						end if
						if Request("hsun1" & i) = "" then
							tblEMP("sun") = 0
						else
							tblEMP("sun") = Request("hsun1" & i)
						end if	
						tmpHRS2 = 0
						tmpHRS2 = Z_Cdbl(Request("hsunx1" & i)) + Z_Cdbl(Request("hmonx1" & i)) +Z_Cdbl(Request("htuex1" & i)) +Z_Cdbl(Request("hwedx1" & i)) + _
							Z_Cdbl(Request("hthux1" & i)) + Z_Cdbl(Request("hfrix1" & i)) + Z_Cdbl(Request("hsatx1" & i))
						If Request("chkEXT1" & i) = "" And tmpHRS2 = 0 Then 
							tblEMP("misc_notes") = Request("Mnotes1" & i)
						End If
						tblEMP("author") = Request("UserID")
						tblEMP.Update
						
						Set rsEXT = Server.CreateObject("ADODB.RecordSet")
						sqlEXT = "SELECT * FROM [tsheets_t] WHERE " & _
							"[emp_id] = '" & tblEMP("emp_ID") & "' AND " & _
							"[date] = #" & tblEMP("date") &  "# AND EXT = True AND client = '" & tblEMP("Client") & "' "
						rsEXT.Open sqlEXT, g_strCONN, 1, 3
						If Not rsEXT.EOF Then
							If Request("chkEXT1" & i) <> "" Then
									tmpHRS2 = 0
									tmpHRS2 = Z_Cdbl(Request("hsunx1" & i)) + Z_Cdbl(Request("hmonx1" & i)) +Z_Cdbl(Request("htuex1" & i)) +Z_Cdbl(Request("hwedx1" & i)) + _
										Z_Cdbl(Request("hthux1" & i)) + Z_Cdbl(Request("hfrix1" & i)) + Z_Cdbl(Request("hsatx1" & i))
									If NOT(Request("chkEXT1" & i) <> "" and tmpHRS2 <> 0) Then rsEXT("misc_notes") = Request("Mnotes1" & i)
									If tmpHRS2 <> 0 Then	 
										If Request("Mnotes1" & i) <> "" Then
											rsEXT("misc_notes") = Request("Mnotes1" & i)
											rsEXT("EXT") = True
											if Request("hmonx1" & i)= "" then
												rsEXT("mon") = 0
											else
												rsEXT("mon") = Request("hmonx1" & i)
											end if
											if Request("htuex1" & i)= "" then
												rsEXT("tue") = 0
											else
												rsEXT("tue") = Request("htuex1" & i)
											end if
											if Request("hwedx1" & i) = "" then
												rsEXT("wed") = 0
											else
												rsEXT("wed") = Request("hwedx1" & i)
											end if
											if Request("hthux1" & i) = "" then
												rsEXT("thu") = 0
											else
												rsEXT("thu") = Request("hthux1" & i)
											end if
											if Request("hfrix1" & i) = "" then
												rsEXT("fri") = 0
											else
												rsEXT("fri") = Request("hfrix1" & i)
											end if
											if Request("hsatx1" & i) = "" then
												rsEXT("sat") = 0
											else
												rsEXT("sat") = Request("hsatx1" & i)
											end if
											if Request("hsunx1" & i) = "" then
												rsEXT("sun") = 0
											else
												rsEXT("sun") = Request("hsunx1" & i)
											end if	
										Else
											Session("MSG") = Session("MSG") & "<br>Notes are required if there are extended hours in week 1 for Consumer " & Request("hdept1" & i) & "."
														
										End If
									Else
										rsEXT("mon") = 0
										rsEXT("tue") = 0
										rsEXT("wed") = 0
										rsEXT("thu") = 0
										rsEXT("fri") = 0
										rsEXT("sat") = 0
										rsEXT("sun") = 0
									End If
							Else
								tmpHrschk = Z_CZero(Request("hmon1" & i)) + Z_CZero(Request("htue1" & i)) + Z_CZero(Request("hwed1" & i)) + _
									Z_CZero(Request("hthu1" & i)) + Z_CZero(Request("hfri1" & i)) + Z_CZero(Request("hsat1" & i)) + _
									Z_CZero(Request("hsun1" & i))
								If tmpHrschk = 0 Then
									Session("MSG") = Session("MSG") & "<br>Total hours for " &  Request("hdept1" & i) & " will become 0 in week 1. Please just delete the row."
								Else	
									rsEXT("mon") = 0
									rsEXT("tue") = 0
									rsEXT("wed") = 0
									rsEXT("thu") = 0
									rsEXT("fri") = 0
									rsEXT("sat") = 0
									rsEXT("sun") = 0 
									rsEXT("misc_notes") = Request("Mnotes1" & i)
								End If
							End If
							rsEXT.Update	
							rsEXT.Close
							Set rsEXT = Nothing
							End If
						'tblEMP.Update
						'End If
					ElseIf tmpHrschk = 0 Then
						Session("MSG") = Session("MSG") & "<br>Total hours for " &  Request("hdept1" & i) & " will become 0 in week 1. Please just delete the row."
					Else
						tblEMP("client") = Request("conid" & i)
						if Request("hmon1" & i)= "" then
							tblEMP("mon") = 0
						else
							tblEMP("mon") = Request("hmon1" & i)
						end if
						if Request("htue1" & i)= "" then
							tblEMP("tue") = 0
						else
							tblEMP("tue") = Request("htue1" & i)
						end if
						if Request("hwed1" & i) = "" then
							tblEMP("wed") = 0
						else
							tblEMP("wed") = Request("hwed1" & i)
						end if
						if Request("hthu1" & i) = "" then
							tblEMP("thu") = 0
						else
							tblEMP("thu") = Request("hthu1" & i)
						end if
						if Request("hfri1" & i) = "" then
							tblEMP("fri") = 0
						else
							tblEMP("fri") = Request("hfri1" & i)
						end if
						if Request("hsat1" & i) = "" then
							tblEMP("sat") = 0
						else
							tblEMP("sat") = Request("hsat1" & i)
						end if
						if Request("hsun1" & i) = "" then
							tblEMP("sun") = 0
						else
							tblEMP("sun") = Request("hsun1" & i)
						end if	
						tmpHRS2 = 0
						tmpHRS2 = Z_Cdbl(Request("hsunx1" & i)) + Z_Cdbl(Request("hmonx1" & i)) +Z_Cdbl(Request("htuex1" & i)) +Z_Cdbl(Request("hwedx1" & i)) + _
							Z_Cdbl(Request("hthux1" & i)) + Z_Cdbl(Request("hfrix1" & i)) + Z_Cdbl(Request("hsatx1" & i))
						If Request("chkEXT1" & i) = "" And tmpHRS2 = 0 Then 
							tblEMP("misc_notes") = Request("Mnotes1" & i)
						End If
						tblEMP("author") = Request("UserID")
						tblEMP.Update
						
						Set rsEXT = Server.CreateObject("ADODB.RecordSet")
						sqlEXT = "SELECT * FROM [tsheets_t] WHERE " & _
							"[emp_id] = '" & tblEMP("emp_ID") & "' AND " & _
							"[date] = #" & tblEMP("date") &  "# AND EXT = True AND client = '" & tblEMP("Client") & "' "
						rsEXT.Open sqlEXT, g_strCONN, 1, 3
						If Not rsEXT.EOF Then
							If Request("chkEXT1" & i) <> "" Then
									tmpHRS2 = 0
									tmpHRS2 = Z_Cdbl(Request("hsunx1" & i)) + Z_Cdbl(Request("hmonx1" & i)) +Z_Cdbl(Request("htuex1" & i)) +Z_Cdbl(Request("hwedx1" & i)) + _
										Z_Cdbl(Request("hthux1" & i)) + Z_Cdbl(Request("hfrix1" & i)) + Z_Cdbl(Request("hsatx1" & i))
									If NOT(Request("chkEXT1" & i) <> "" and tmpHRS2 <> 0) Then rsEXT("misc_notes") = Request("Mnotes1" & i)
									If tmpHRS2 <> 0 Then	 
										If Request("Mnotes1" & i) <> "" Then
											rsEXT("misc_notes") = Request("Mnotes1" & i)
											rsEXT("EXT") = True
											if Request("hmonx1" & i)= "" then
												rsEXT("mon") = 0
											else
												rsEXT("mon") = Request("hmonx1" & i)
											end if
											if Request("htuex1" & i)= "" then
												rsEXT("tue") = 0
											else
												rsEXT("tue") = Request("htuex1" & i)
											end if
											if Request("hwedx1" & i) = "" then
												rsEXT("wed") = 0
											else
												rsEXT("wed") = Request("hwedx1" & i)
											end if
											if Request("hthux1" & i) = "" then
												rsEXT("thu") = 0
											else
												rsEXT("thu") = Request("hthux1" & i)
											end if
											if Request("hfrix1" & i) = "" then
												rsEXT("fri") = 0
											else
												rsEXT("fri") = Request("hfrix1" & i)
											end if
											if Request("hsatx1" & i) = "" then
												rsEXT("sat") = 0
											else
												rsEXT("sat") = Request("hsatx1" & i)
											end if
											if Request("hsunx1" & i) = "" then
												rsEXT("sun") = 0
											else
												rsEXT("sun") = Request("hsunx1" & i)
											end if	
										Else
											Session("MSG") = Session("MSG") & "<br>Notes are required if there are extended hours in week 1 for Consumer " & Request("hdept1" & i) & "."
														
										End If
									Else
										rsEXT("mon") = 0
										rsEXT("tue") = 0
										rsEXT("wed") = 0
										rsEXT("thu") = 0
										rsEXT("fri") = 0
										rsEXT("sat") = 0
										rsEXT("sun") = 0
									End If
							Else
								tmpHrschk = Z_CZero(Request("hmon1" & i)) + Z_CZero(Request("htue1" & i)) + Z_CZero(Request("hwed1" & i)) + _
									Z_CZero(Request("hthu1" & i)) + Z_CZero(Request("hfri1" & i)) + Z_CZero(Request("hsat1" & i)) + _
									Z_CZero(Request("hsun1" & i))
								If tmpHrschk = 0 Then
									Session("MSG") = Session("MSG") & "<br>Total hours for " &  Request("hdept1" & i) & " will become 0 in week 1. Please just delete the row."
								Else	
									rsEXT("mon") = 0
									rsEXT("tue") = 0
									rsEXT("wed") = 0
									rsEXT("thu") = 0
									rsEXT("fri") = 0
									rsEXT("sat") = 0
									rsEXT("sun") = 0 
									rsEXT("misc_notes") = Request("Mnotes1" & i)
								End If
							End If
							rsEXT.Update	
							rsEXT.Close
							Set rsEXT = Nothing
						End If
						'tblEMP.Update
					End If
				End If 'end
			End if
			tblEMP.Close
			Set tblEMP = Nothing
			'''CHECK NOTES
			If tmpctr <> "" Then
				Set rsNotes = Server.CreateObject("ADODB.RecordSet")
				sqlNotes = "SELECT * FROM TSHEETS_T"
				rsNotes.Open sqlNotes, g_strCONN, 1, 3
				rsNotes.MoveFirst
				rsNotes.Find(strTmp)
				If Not rsNotes.EOF Then
					tmpNotes1 = rsNotes("misc_notes")
					Set rsNotes2 = Server.CreateObject("ADODB.RecordSet")
					sqlNotes2 = "SELECT * FROM [tsheets_t] WHERE " & _
								"[emp_id] = '" & rsNotes("emp_ID") & "' AND " & _
								"[date] = #" & rsNotes("date") &  "# AND EXT = True AND client = '" & rsNotes("Client") & "' "
						rsNotes2.Open sqlNotes2, g_strCONN, 3, 1
						If Not rsNotes2.EOF Then
							tmpNotes2 = rsNotes2("misc_notes")
						End If
						rsNotes2.Close
						Set rsNotes2 = Nothing	
						If tmpNotes2 <> tmpNotes1 Then
							rsNotes("misc_notes") = tmpNotes2
						End If
					rsNotes.update
				End If
				rsNotes.Close
				Set rsNotes = Nothing 
			End If		
			'''''
		Next
	End If 
	
	
	tmpHrsC = 0
	tmpHrsC = Z_CZero(Request("hmonX1")) + Z_CZero(Request("htueX1")) + Z_CZero(Request("hwedX1")) + _
			Z_CZero(Request("hthuX1")) + Z_CZero(Request("hfriX1")) + Z_CZero(Request("hsatX1")) + _
			Z_CZero(Request("hsunX1"))
	tmpHrs2C = 0
	tmpHrs2C = Z_CZero(Request("hmon1")) + Z_CZero(Request("htue1")) + Z_CZero(Request("hwed1")) + _
			Z_CZero(Request("hthu1")) + Z_CZero(Request("hfri1")) + Z_CZero(Request("hsat1")) + _
			Z_CZero(Request("hsun1"))
	if Request("hdept1") <>  ""  then
		If tmpHrsC <> 0 Or tmpHrs2C <> 0 then
			'''get max hrs
			Set tblCon = Server.CreateObject("ADODB.RecordSet")
			sqlCon = "SELECT * FROM Consumer_t WHERE medicaid_number = '" & Request("hdept1") & "' "
			tblCon.Open sqlCon, g_strCONN, 1, 3
			If Not tblCon.EOF Then
				tmpMax = Z_CZero(tblCon("MaxHrs"))
			End If
			tblCon.Close
			Set tblCon = Nothing
			''''
			''''get hours consumed
			sqlChk = "SELECT * FROM tSHEETS_T WHERE date = #" & Request("1day") & "# AND client = '" & _
			Request("hdept1") & "' AND emp_ID <> '" & Session("idemp") & "' AND Ext = false"
			Set rsChk = Server.CreateObject("ADODB.RecordSet")
			rsChk.Open sqlChk, g_strCONN, 3, 1
			If Not rsChk.EOF Then
				HrsCon = 0
				Do Until rsChk.EOF
					HrsCon = HrsCon + Z_CZero(rsChk("mon")) + Z_CZero(rsChk("tue")) + Z_CZero(rsChk("wed")) + Z_CZero(rsChk("thu")) + Z_CZero(rsChk("fri")) + _
						Z_CZero(rsChk("sat")) + Z_CZero(rsChk("sun"))
					rsChk.MoveNExt
				Loop
			End If	
			rsChk.Close
			Set rsChk = Nothing
			''''
			tmpHrs = HrsCon + Z_CZero(Request("hmon1")) + Z_CZero(Request("htue1")) + Z_CZero(Request("hwed1")) + _
					Z_CZero(Request("hthu1")) + Z_CZero(Request("hfri1")) + Z_CZero(Request("hsat1")) + _
					Z_CZero(Request("hsun1"))
			If tmpHrs > tmpMax Then
				xHRS = Z_DoEncrypt(Request("hmon1") & "|" & Request("htue1") & "|" & Request("hwed1") & "|" & _
						Request("hthu1") & "|" & Request("hfri1") & "|" & Request("hsat1") & "|" & _
						Request("hsun1") & "|" & Request("Mnotes1") & "|" & Request("hmonX1") & "|" & Request("htueX1") & "|" & Request("hwedX1") & "|" & _
						Request("hthuX1") & "|" & Request("hfriX1") & "|" & Request("hsatX1") & "|" & _
						Request("hsunX1"))
				Set rsConNam = Server.CreateObject("ADODB.RecordSet")
				sqlConNam = "SELECT * FROM Consumer_t WHERE Medicaid_Number = '" &  Request("hdept1") & "' "
				rsConNam.Open sqlConNam, g_strCONN, 1, 3
					ConNam = Request("hdept1")
				If Not rsConNam.EOF Then
					ConNam = rsConNam("LName") & ", " & rsConNam("FName")
				End If 	
				rsConNam.Close
				Set rsConNam = Nothing
				Session("MSG") = Session("MSG") & "<br>Total hours for " &  ConNam & " is over the allowed hours in week 1."
				'Response.Redirect "view.asp?Con=" & Request("hdept1") & "&xHRS=" & xHRS
			Else
				If Request("chkEXT") <> "" Then
					If Not Request("Mnotes1") <> "" then
						xHRS = Z_DoEncrypt(Request("hmon1") & "|" & Request("htue1") & "|" & Request("hwed1") & "|" & _
							Request("hthu1") & "|" & Request("hfri1") & "|" & Request("hsat1") & "|" & _
							Request("hsun1") & "|" & Request("Mnotes1") & "|" & Request("hmonX1") & "|" & Request("htueX1") & "|" & Request("hwedX1") & "|" & _
							Request("hthuX1") & "|" & Request("hfriX1") & "|" & Request("hsatX1") & "|" & _
							Request("hsunX1"))
						Set rsConNam = Server.CreateObject("ADODB.RecordSet")
						sqlConNam = "SELECT * FROM Consumer_t WHERE Medicaid_Number = '" &  Request("hdept1") & "' "
						rsConNam.Open sqlConNam, g_strCONN, 1, 3
							ConNam = Request("hdept1")
						If Not rsConNam.EOF Then
							ConNam = rsConNam("LName") & ", " & rsConNam("FName")
						End If 	
						rsConNam.Close
						Set rsConNam = Nothing
						Session("MSG") = Session("MSG") & "<br>Notes are required if there are extended hours in week 1 for consumer " & ConNam & "."
						'Response.Redirect "view.asp?Con=" & Request("hdept1") & "&xHRS=" & xHRS
					Else
						Set tblEMP = Server.CreateObject("ADODB.RecordSet")
						strSQL = "SELECT * FROM tsheets_T"
						tblEMP.Open strSQL, g_strCONN, 1, 3
						tblEMP.AddNew
						tblEMP("client") = Request("hdept1")
						if Request("hmon1")= "" then
							tblEMP("mon") = 0
						else
							tblEMP("mon") = Request("hmon1")
						end if
						if Request("htue1")= "" then
							tblEMP("tue") = 0
						else
							tblEMP("tue") = Request("htue1")
						end if
						if Request("hwed1") = "" then
							tblEMP("wed") = 0
						else
							tblEMP("wed") = Request("hwed1")
						end if
						if Request("hthu1") = "" then
							tblEMP("thu") = 0
						else
							tblEMP("thu") = Request("hthu1")
						end if
						if Request("hfri1") = "" then
							tblEMP("fri") = 0
						else
							tblEMP("fri") = Request("hfri1")
						end if
						if Request("hsat1") = "" then
							tblEMP("sat") = 0
						else
							tblEMP("sat") = Request("hsat1")
						end if
						if Request("hsun1") = "" then
							tblEMP("sun") = 0
						else
							tblEMP("sun") = Request("hsun1")
						end if	
						tblEMP("date") = Request("1day")
						tblEMP("emp_id") = Session("idemp") 
						tblEMP("misc_notes") = Request("Mnotes1")
						tblEMP("author") = session("UserID")
						tblEMP.Update
						tblEMP.Close
						Set tblEMP = Nothing
						Set tblEXT = Server.CreateObject("ADODB.Recordset")
						strSQL = "SELECT * FROM [tsheets_t]"
						tblEXT.Open strSQL, g_strCONN, 1, 3
						tblEXT.addnew
						tblEXT("EXT") = True
						tblEXT("client") = Request("hdept1")
						if Request("hmonx1")= "" then
							tblEXT("mon") = 0
						else
							tblEXT("mon") = Request("hmonx1")
						end if
						if Request("htuex1")= "" then
							tblEXT("tue") = 0
						else
							tblEXT("tue") = Request("htuex1")
						end if
						if Request("hwedx1") = "" then
							tblEXT("wed") = 0
						else
							tblEXT("wed") = Request("hwedx1")
						end if
						if Request("hthux1") = "" then
							tblEXT("thu") = 0
						else
							tblEXT("thu") = Request("hthux1")
						end if
						if Request("hfrix1") = "" then
							tblEXT("fri") = 0
						else
							tblEXT("fri") = Request("hfrix1")
						end if
						if Request("hsatx1") = "" then
							tblEXT("sat") = 0
						else
							tblEXT("sat") = Request("hsatx1")
						end if
						if Request("hsunx1") = "" then
							tblEXT("sun") = 0
						else
							tblEXT("sun") = Request("hsunx1")
						end if	
						tblEXT("date") = Request("1day")
						tblEXT("emp_id") = Session("idemp") 
						tblEXT("author") = session("UserID")
						tblEXT("misc_notes") = Request("Mnotes1")
						tblEXT.UPDATE
						tblEXT.Close
						Set tblEXT = Nothing
					End If
				Else
					Set tblEMP = Server.CreateObject("ADODB.RecordSet")
					strSQL = "SELECT * FROM tsheets_T"
					tblEMP.Open strSQL, g_strCONN, 1, 3
					tblEMP.AddNew
					tblEMP("client") = Request("hdept1")
					if Request("hmon1")= "" then
						tblEMP("mon") = 0
					else
						tblEMP("mon") = Request("hmon1")
					end if
					if Request("htue1")= "" then
						tblEMP("tue") = 0
					else
						tblEMP("tue") = Request("htue1")
					end if
					if Request("hwed1") = "" then
						tblEMP("wed") = 0
					else
						tblEMP("wed") = Request("hwed1")
					end if
					if Request("hthu1") = "" then
						tblEMP("thu") = 0
					else
						tblEMP("thu") = Request("hthu1")
					end if
					if Request("hfri1") = "" then
						tblEMP("fri") = 0
					else
						tblEMP("fri") = Request("hfri1")
					end if
					if Request("hsat1") = "" then
						tblEMP("sat") = 0
					else
						tblEMP("sat") = Request("hsat1")
					end if
					if Request("hsun1") = "" then
						tblEMP("sun") = 0
					else
						tblEMP("sun") = Request("hsun1")
					end if	
					tblEMP("date") = Request("1day")
					tblEMP("emp_id") = Session("idemp") 
					tblEMP("misc_notes") = Request("Mnotes1")
					tblEMP("author") = session("UserID")
					tblEMP.Update
					tblEMP.Close
					Set tblEMP = Nothing
					Set tblEXT = Server.CreateObject("ADODB.Recordset")
					strSQL = "SELECT * FROM [tsheets_t]"
					tblEXT.Open strSQL, g_strCONN, 1, 3
					tblEXT.addnew
					tblEXT("EXT") = True
					tblEXT("client") = Request("hdept1")
					tblEXT("mon") = 0
					tblEXT("tue") = 0
					tblEXT("wed") = 0
					tblEXT("thu") = 0
					tblEXT("fri") = 0
					tblEXT("sat") = 0
					tblEXT("sun") = 0
					tblEXT("date") = Request("1day")
					tblEXT("emp_id") = Session("idemp") 
					tblEXT("author") = session("UserID")
					tblEXT("misc_notes") = Request("Mnotes1")
					tblEXT.UPDATE
					tblEXT.Close
					Set tblEXT = Nothing
				End If
			End If
			'''PRocess table
			Set tblProc = Server.CreateObject("ADODB.RecordSet")
			sqlProc = "SELECT * FROM Process_t WHERE Wor = '" & Session("idemp") & "' AND Con = '" &  Request("hdept1") & "' " & _
				"AND Payd8 = #" & Request("1day") & "# "
			tblProc.Open sqlProc, g_strCONN, 1, 3
				If tblProc.EOF Then
				 	tblProc.AddNew
				 	tblProc("Wor") = Session("idemp")
				 	tblProc("Con") = Request("hdept1")
				 	tblProc("TSdate") = Request("1day")
				 	tblProc.Update
			 	End If
			tblProc.Close
			Set tblProc = Nothing
			''''
		End If	
		If tmpHrsC = 0 And tmpHrs2C = 0 then
			Session("MSG") = Session("MSG") & "<br>Total hours cannot be equal to 0 in new consumer in week 1." 
			
		End If
	Else
		If tmpHrsC <> 0 Or tmpHrs2C <> 0 then
			xHRS = Z_DoEncrypt(Request("hmon1") & "|" & Request("htue1") & "|" & Request("hwed1") & "|" & _
						Request("hthu1") & "|" & Request("hfri1") & "|" & Request("hsat1") & "|" & _
						Request("hsun1") & "|" & Request("Mnotes1") & "|" & Request("hmonX1") & "|" & Request("htueX1") & "|" & Request("hwedX1") & "|" & _
						Request("hthuX1") & "|" & Request("hfriX1") & "|" & Request("hsatX1") & "|" & _
						Request("hsunX1"))
			Session("MSG") = Session("MSG") & "<br>Please select a new consumer in week 1." 
		
		End If
	End If
End If
'''WEEK2
If Request("paychk2") = "" AND Request("medchk2") = "" Then
	ctrI = Request("count2")
	If Request("count2") <> 0 Then
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
					''''Get Max Hours
					Set tblCon = Server.CreateObject("ADODB.RecordSet")
					sqlCon = "SELECT * FROM Consumer_t WHERE medicaid_number = '" & Request("conid2" & i) & "' "
					tblCon.Open sqlCon, g_strCONN, 1, 3
					If Not tblCon.EOF Then
						tmpMax = Z_CZero(tblCon("MaxHrs"))
					End If
					tblCon.Close
					Set tblCon = Nothing
					''''
					'''get existing hours
					Set rsChk = Server.CreateObject("ADODB.RecordSet")
					sqlChk = "SELECT * FROM tSHEETS_T WHERE date = #" & Request("2day") & "# AND client = '" & _
					Request("conid2" & i) & "' AND emp_ID <> '" & Session("idemp") & "' AND Ext = false"
					rsChk.Open sqlChk, g_strCONN, 3, 1
					If Not rsChk.EOF Then
						HrsCon = 0
						Do Until rsChk.EOF
							HrsCon = HrsCon + Z_CZero(rsChk("mon")) + Z_CZero(rsChk("tue")) + Z_CZero(rsChk("wed")) + Z_CZero(rsChk("thu")) + Z_CZero(rsChk("fri")) + _
								Z_CZero(rsChk("sat")) + Z_CZero(rsChk("sun"))
							rsChk.MoveNExt
						Loop
					End If	
					rsChk.Close
					Set rsChk = Nothing
					''''
					tmpHrs = HrsCon + Z_CZero(Request("hmon2" & i)) + Z_CZero(Request("htue2" & i)) + Z_CZero(Request("hwed2" & i)) + _
							Z_CZero(Request("hthu2" & i)) + Z_CZero(Request("hfri2" & i)) + Z_CZero(Request("hsat2" & i)) + _
							Z_CZero(Request("hsun2" & i))
					tmpHrschk = Z_CZero(Request("hmon2" & i)) + Z_CZero(Request("htue2" & i)) + Z_CZero(Request("hwed2" & i)) + _
								Z_CZero(Request("hthu2" & i)) + Z_CZero(Request("hfri2" & i)) + Z_CZero(Request("hsat2" & i)) + _
								Z_CZero(Request("hsun2" & i)) + Z_CZero(Request("hmonX2" & i)) + Z_CZero(Request("htueX2" & i)) + Z_CZero(Request("hwedX2" & i)) + _
								Z_CZero(Request("hthuX2" & i)) + Z_CZero(Request("hfriX2" & i)) + Z_CZero(Request("hsatX2" & i)) + _
								Z_CZero(Request("hsunX2" & i))
					If tmpHrs > tmpMax Then
						Session("MSG") = Session("MSG") & "<br>Total hours for " &  Request("hdept2" & i) & " is over the allowed hours in week 2."
					ElseIf tmpHrschk = 0 Then
						Session("MSG") = Session("MSG") & "<br>Total hours for " &  Request("hdept2" & i) & " will become 0 in week 2. Please just delete the row."
					Else		
						tblEMP("client") = Request("conid2" & i)
						if Request("hmon2" & i)= "" then
							tblEMP("mon") = 0
						else
							tblEMP("mon") = Request("hmon2" & i)
						end if
						if Request("htue2" & i)= "" then
							tblEMP("tue") = 0
						else
							tblEMP("tue") = Request("htue2" & i)
						end if
						if Request("hwed2" & i) = "" then
							tblEMP("wed") = 0
						else
							tblEMP("wed") = Request("hwed2" & i)
						end if
						if Request("hthu2" & i) = "" then
							tblEMP("thu") = 0
						else
							tblEMP("thu") = Request("hthu2" & i)
						end if
						if Request("hfri2" & i) = "" then
							tblEMP("fri") = 0
						else
							tblEMP("fri") = Request("hfri2" & i)
						end if
						if Request("hsat2" & i) = "" then
							tblEMP("sat") = 0
						else
							tblEMP("sat") = Request("hsat2" & i)
						end if
						if Request("hsun2" & i) = "" then
							tblEMP("sun") = 0
						else
							tblEMP("sun") = Request("hsun2" & i)
						end if	
						tmpHRS2 = 0
						tmpHRS2 = Z_Cdbl(Request("hsunx2" & i)) + Z_Cdbl(Request("hmonx2" & i)) +Z_Cdbl(Request("htuex2" & i)) +Z_Cdbl(Request("hwedx2" & i)) + _
							Z_Cdbl(Request("hthux2" & i)) + Z_Cdbl(Request("hfrix2" & i)) + Z_Cdbl(Request("hsatx2" & i))
						If Request("chkEXT2" & i) = "" And tmpHRS2 = 0 Then 
							tblEMP("misc_notes") = Request("Mnotes2" & i)
						End If
						tblEMP("author") = Request("UserID")
						tblEMP.Update
						
						Set rsEXT = Server.CreateObject("ADODB.RecordSet")
						sqlEXT = "SELECT * FROM [tsheets_t] WHERE " & _
							"[emp_id] = '" & tblEMP("emp_ID") & "' AND " & _
							"[date] = #" & tblEMP("date") &  "# AND EXT = True AND client = '" & tblEMP("Client") & "' "
						rsEXT.Open sqlEXT, g_strCONN, 1, 3
						If Not rsEXT.EOF Then
							If Request("chkEXT2" & i) <> "" Then
								tmpHRS2 = 0
								tmpHRS2 = Z_Cdbl(Request("hsunx2" & i)) + Z_Cdbl(Request("hmonx2" & i)) +Z_Cdbl(Request("htuex2" & i)) +Z_Cdbl(Request("hwedx2" & i)) + _
									Z_Cdbl(Request("hthux2" & i)) + Z_Cdbl(Request("hfrix2" & i)) + Z_Cdbl(Request("hsatx2" & i))
								If NOT(Request("chkEXT2" & i) <> "" and tmpHRS2 <> 0) Then rsEXT("misc_notes") = Request("Mnotes2" & i)
								If tmpHRS2 <> 0 Then	 
									If Request("Mnotes2" & i) <> "" Then
										rsEXT("misc_notes") = Request("Mnotes2" & i)
										rsEXT("EXT") = True
											if Request("hmonx2" & i)= "" then
												rsEXT("mon") = 0
											else
												rsEXT("mon") = Request("hmonx2" & i)
											end if
											if Request("htuex2" & i)= "" then
												rsEXT("tue") = 0
											else
												rsEXT("tue") = Request("htuex2" & i)
											end if
											if Request("hwedx2" & i) = "" then
												rsEXT("wed") = 0
											else
												rsEXT("wed") = Request("hwedx2" & i)
											end if
											if Request("hthux2" & i) = "" then
												rsEXT("thu") = 0
											else
												rsEXT("thu") = Request("hthux2" & i)
											end if
											if Request("hfrix2" & i) = "" then
												rsEXT("fri") = 0
											else
												rsEXT("fri") = Request("hfrix2" & i)
											end if
											if Request("hsatx2" & i) = "" then
												rsEXT("sat") = 0
											else
												rsEXT("sat") = Request("hsatx2" & i)
											end if
											if Request("hsunx2" & i) = "" then
												rsEXT("sun") = 0
											else
												rsEXT("sun") = Request("hsunx2" & i)
											end if	
							
									Else
										Session("MSG") = Session("MSG") & "<br>Notes are required if there are extended hours in week 2 for Consumer " & Request("hdept2" & i) & "."
										'Response.Redirect "view.asp"
										'rsEXT("mon") = 0
									'	rsEXT("tue") = 0
										'rsEXT("wed") = 0
										'rsEXT("thu") = 0
										'rsEXT("fri") = 0
										'rsEXT("sat") = 0
										'rsEXT("sun") = 0
									End If
								End If
							Else
								tmpHrschk = Z_CZero(Request("hmon2" & i)) + Z_CZero(Request("htue2" & i)) + Z_CZero(Request("hwed2" & i)) + _
									Z_CZero(Request("hthu2" & i)) + Z_CZero(Request("hfri2" & i)) + Z_CZero(Request("hsat2" & i)) + _
									Z_CZero(Request("hsun2" & i))
								If tmpHrschk = 0 Then
									Session("MSG") = Session("MSG") & "<br>Total hours for " &  Request("hdept2" & i) & " will become 0 in week 2. Please just delete the row."
								Else	
									rsEXT("mon") = 0
									rsEXT("tue") = 0
									rsEXT("wed") = 0
									rsEXT("thu") = 0
									rsEXT("fri") = 0
									rsEXT("sat") = 0
									rsEXT("sun") = 0 
									rsEXT("misc_notes") = Request("Mnotes2" & i)
								End If
							End If
						rsEXT.Update	
						rsEXT.Close
						Set rsEXT = Nothing
					End If
					'tblEMP.Update
				End If
			End If
		End If
		tblEMP.Close
		Set tblEMP = Nothing
		'''CHECK NOTES
			If tmpctr <> "" Then
				Set rsNotes = Server.CreateObject("ADODB.RecordSet")
				sqlNotes = "SELECT * FROM TSHEETS_T"
				rsNotes.Open sqlNotes, g_strCONN, 1, 3
				rsNotes.MoveFirst
				rsNotes.Find(strTmp)
				If Not rsNotes.EOF Then
					tmpNotes1 = rsNotes("misc_notes")
					Set rsNotes2 = Server.CreateObject("ADODB.RecordSet")
					sqlNotes2 = "SELECT * FROM [tsheets_t] WHERE " & _
								"[emp_id] = '" & rsNotes("emp_ID") & "' AND " & _
								"[date] = #" & rsNotes("date") &  "# AND EXT = True AND client = '" & rsNotes("Client") & "' "
						rsNotes2.Open sqlNotes2, g_strCONN, 3, 1
						If Not rsNotes2.EOF Then
							tmpNotes2 = rsNotes2("misc_notes")
						End If
						rsNotes2.Close
						Set rsNotes2 = Nothing	
						If tmpNotes2 <> tmpNotes1 Then
							rsNotes("misc_notes") = tmpNotes2
						End If
					rsNotes.update
				End If
				rsNotes.Close
				Set rsNotes = Nothing 
			End If		
			'''''
	Next 
	End If	
	
	tmpHrs2C = 0
	tmpHrs2C = Z_CZero(Request("hmonX2")) + Z_CZero(Request("htueX2")) + Z_CZero(Request("hwedX2")) + _
			Z_CZero(Request("hthuX2")) + Z_CZero(Request("hfriX2")) + Z_CZero(Request("hsatX2")) + _
			Z_CZero(Request("hsunX2"))
	tmpHrsC = 0
	tmpHrsC = Z_CZero(Request("hmon2")) + Z_CZero(Request("htue2")) + Z_CZero(Request("hwed2")) + _
			Z_CZero(Request("hthu2")) + Z_CZero(Request("hfri2")) + Z_CZero(Request("hsat2")) + _
			Z_CZero(Request("hsun2"))
	if Request("hdept2") <> "" then
		If tmpHrsC <> 0 Or tmpHrs2C <> 0 then
			'''get max hrs
			Set tblCon = Server.CreateObject("ADODB.RecordSet")
			sqlCon = "SELECT * FROM Consumer_t WHERE medicaid_number = '" & Request("hdept2") & "' "
			tblCon.Open sqlCon, g_strCONN, 1, 3
			If Not tblCon.EOF Then
				tmpMax = Z_CZero(tblCon("MaxHrs"))
			End If
			tblCon.Close
			Set tblCon = Nothing
			''''
			''''get hours consumed
			sqlChk = "SELECT * FROM tSHEETS_T WHERE date = #" & Request("2day") & "# AND client = '" & _
			 Request("hdept2") & "' AND emp_ID <> '" & Session("idemp") & "' AND Ext = false"
			Set rsChk = Server.CreateObject("ADODB.RecordSet")
			rsChk.Open sqlChk, g_strCONN, 3, 1
			If Not rsChk.EOF Then
				HrsCon = 0
				Do Until rsChk.EOF
					HrsCon = HrsCon + Z_CZero(rsChk("mon")) + Z_CZero(rsChk("tue")) + Z_CZero(rsChk("wed")) + Z_CZero(rsChk("thu")) + Z_CZero(rsChk("fri")) + _
						Z_CZero(rsChk("sat")) + Z_CZero(rsChk("sun"))
					rsChk.MoveNExt
				Loop
			End If	
			rsChk.Close
			Set rsChk = Nothing
			''''
			tmpHrs = HrsCon + Z_CZero(Request("hmon2")) + Z_CZero(Request("htue2")) + Z_CZero(Request("hwed2")) + _
				Z_CZero(Request("hthu2")) + Z_CZero(Request("hfri2")) + Z_CZero(Request("hsat2")) + _
				Z_CZero(Request("hsun2"))
			If tmpHrs > tmpMax Then
				xHRS2 = Z_DoEncrypt(Request("hmon2") & "|" & Request("htue2") & "|" & Request("hwed2") & "|" & _
						Request("hthu2") & "|" & Request("hfri2") & "|" & Request("hsat2") & "|" & _
						Request("hsun2") & "|" & Request("Mnotes2") & "|" & Request("hmonX2") & "|" & Request("htueX2") & "|" & Request("hwedX2") & "|" & _
						Request("hthuX2") & "|" & Request("hfriX2") & "|" & Request("hsatX2") & "|" & _
						Request("hsunX2"))
				Set rsConNam = Server.CreateObject("ADODB.RecordSet")
				sqlConNam = "SELECT * FROM Consumer_t WHERE Medicaid_Number = '" &  Request("hdept2") & "' "
				rsConNam.Open sqlConNam, g_strCONN, 1, 3
					ConNam = Request("hdept2")
				If Not rsConNam.EOF Then
					ConNam = rsConNam("LName") & ", " & rsConNam("FName")
				End If 	
				rsConNam.Close
				Set rsConNam = Nothing
				Session("MSG") = Session("MSG") & "<br>Total hours for " &  ConNam & " is over the allowed hours in week 2."
				'Response.Redirect "view.asp?Con2=" & Request("hdept2") & "&xHRS2=" & xHRS2
			Else
				If Request("chkEXT2") <> "" Then
					If Not Request("Mnotes2") <> "" then
						xHRS2 = Z_DoEncrypt(Request("hmon2") & "|" & Request("htue2") & "|" & Request("hwed2") & "|" & _
							Request("hthu2") & "|" & Request("hfri2") & "|" & Request("hsat2") & "|" & _
							Request("hsun2") & "|" & Request("Mnotes2") & "|" & Request("hmonX2") & "|" & Request("htueX2") & "|" & Request("hwedX2") & "|" & _
							Request("hthuX2") & "|" & Request("hfriX2") & "|" & Request("hsatX2") & "|" & _
							Request("hsunX2"))
						Set rsConNam = Server.CreateObject("ADODB.RecordSet")
						sqlConNam = "SELECT * FROM Consumer_t WHERE Medicaid_Number = '" &  Request("hdept2") & "' "
						rsConNam.Open sqlConNam, g_strCONN, 1, 3
							ConNam = Request("hdept1")
						If Not rsConNam.EOF Then
							ConNam = rsConNam("LName") & ", " & rsConNam("FName")
						End If 	
						rsConNam.Close
						Set rsConNam = Nothing
						Session("MSG") = Session("MSG") & "<br>Notes are required if there are extended hours in week 2 for consumer " & ConNam & "."
						'Response.Redirect "view.asp?Con2=" & Request("hdept2") & "&xHRS2=" & xHRS2
					Else
						Set tblEMP = Server.CreateObject("ADODB.RecordSet")
						strSQL = "SELECT * FROM tsheets_T"
						tblEMP.Open strSQL, g_strCONN, 1, 3
						tblEMP.AddNew
						tblEMP("client") = Request("hdept2")
						if Request("hmon2")= "" then
							tblEMP("mon") = 0
						else
							tblEMP("mon") = Request("hmon2")
						end if
						if Request("htue2")= "" then
							tblEMP("tue") = 0
						else
							tblEMP("tue") = Request("htue2")
						end if
						if Request("hwed2") = "" then
							tblEMP("wed") = 0
						else
							tblEMP("wed") = Request("hwed2")
						end if
						if Request("hthu2") = "" then
							tblEMP("thu") = 0
						else
							tblEMP("thu") = Request("hthu2")
						end if
						if Request("hfri2") = "" then
							tblEMP("fri") = 0
						else
							tblEMP("fri") = Request("hfri2")
						end if
						if Request("hsat2") = "" then
							tblEMP("sat") = 0
						else
							tblEMP("sat") = Request("hsat2")
						end if
						if Request("hsun2") = "" then
							tblEMP("sun") = 0
						else
							tblEMP("sun") = Request("hsun2")
						end if	
						tblEMP("date") = Request("2day")
						tblEMP("emp_id") = Session("idemp") 
						tblEMP("misc_notes") = Request("Mnotes2")
						tblEMP("author") = session("UserID")
						tblEMP.Update
						tblEMP.Close
						Set tblEMP = Nothing
						Set tblEXT = Server.CreateObject("ADODB.Recordset")
						strSQL = "SELECT * FROM [tsheets_t]"
						tblEXT.Open strSQL, g_strCONN, 1, 3
						tblEXT.addnew
						tblEXT("EXT") = True
						tblEXT("client") = Request("hdept2")
						if Request("hmonx2")= "" then
							tblEXT("mon") = 0
						else
							tblEXT("mon") = Request("hmonx2")
						end if
						if Request("htuex2")= "" then
							tblEXT("tue") = 0
						else
							tblEXT("tue") = Request("htuex2")
						end if
						if Request("hwedx2") = "" then
							tblEXT("wed") = 0
						else
							tblEXT("wed") = Request("hwedx2")
						end if
						if Request("hthux2") = "" then
							tblEXT("thu") = 0
						else
							tblEXT("thu") = Request("hthux2")
						end if
						if Request("hfrix2") = "" then
							tblEXT("fri") = 0
						else
							tblEXT("fri") = Request("hfrix2")
						end if
						if Request("hsatx2") = "" then
							tblEXT("sat") = 0
						else
							tblEXT("sat") = Request("hsatx2")
						end if
						if Request("hsunx2") = "" then
							tblEXT("sun") = 0
						else
							tblEXT("sun") = Request("hsunx2")
						end if	
						tblEXT("date") = Request("2day")
						tblEXT("emp_id") = Session("idemp") 
						tblEXT("author") = session("UserID")
						tblEXT("misc_notes") = Request("Mnotes2")
						tblEXT.UPDATE
						tblEXT.Close
						Set tblEXT = Nothing
					End If
				Else
					Set tblEMP = Server.CreateObject("ADODB.RecordSet")
						strSQL = "SELECT * FROM tsheets_T"
						tblEMP.Open strSQL, g_strCONN, 1, 3
						tblEMP.AddNew
						tblEMP("client") = Request("hdept2")
						if Request("hmon2")= "" then
							tblEMP("mon") = 0
						else
							tblEMP("mon") = Request("hmon2")
						end if
						if Request("htue2")= "" then
							tblEMP("tue") = 0
						else
							tblEMP("tue") = Request("htue2")
						end if
						if Request("hwed2") = "" then
							tblEMP("wed") = 0
						else
							tblEMP("wed") = Request("hwed2")
						end if
						if Request("hthu2") = "" then
							tblEMP("thu") = 0
						else
							tblEMP("thu") = Request("hthu2")
						end if
						if Request("hfri2") = "" then
							tblEMP("fri") = 0
						else
							tblEMP("fri") = Request("hfri2")
						end if
						if Request("hsat2") = "" then
							tblEMP("sat") = 0
						else
							tblEMP("sat") = Request("hsat2")
						end if
						if Request("hsun2") = "" then
							tblEMP("sun") = 0
						else
							tblEMP("sun") = Request("hsun2")
						end if	
						tblEMP("date") = Request("2day")
						tblEMP("emp_id") = Session("idemp") 
						tblEMP("misc_notes") = Request("Mnotes2")
						tblEMP("author") = session("UserID")
						tblEMP.Update
						tblEMP.Close
						Set tblEMP = Nothing
						Set tblEXT = Server.CreateObject("ADODB.Recordset")
					strSQL = "SELECT * FROM [tsheets_t]"
					tblEXT.Open strSQL, g_strCONN, 1, 3
					tblEXT.addnew
					tblEXT("EXT") = True
					tblEXT("client") = Request("hdept2")
					tblEXT("mon") = 0
					tblEXT("tue") = 0
					tblEXT("wed") = 0
					tblEXT("thu") = 0
					tblEXT("fri") = 0
					tblEXT("sat") = 0
					tblEXT("sun") = 0
					tblEXT("date") = Request("2day")
					tblEXT("emp_id") = Session("idemp") 
					tblEXT("author") = session("UserID")
					tblEXT("misc_notes") = Request("Mnotes2")
					tblEXT.UPDATE
					tblEXT.Close
					Set tblEXT = Nothing
				End If
			End If
			'''PRocess table	
			Set tblProc = Server.CreateObject("ADODB.RecordSet")
			sqlProc = "SELECT * FROM Process_t WHERE Wor = '" & Session("idemp") & "' AND Con = '" &  Request("hdept2") & "' " & _
				"AND Payd8 = #" & Request("2day") & "# "
			tblProc.Open sqlProc, g_strCONN, 1, 3
			If tblProc.EOF Then
			 	tblProc.AddNew
			 	tblProc("Wor") = Session("idemp")
			 	tblProc("Con") = Request("hdept2")
			 	tblProc("TSdate") = Request("2day")
			 	tblProc.Update
		 	End If
			tblProc.Close
			Set tblProc = Nothing
			''''
		End If
		If tmpHrsC = 0 And tmpHrs2C = 0 then
			Session("MSG") = Session("MSG") & "<br>Total hours cannot be equal to 0 in new consumer in week 2." 
		End If
	Else
		If tmpHrsC <> 0 Or tmpHrs2C <> 0 then
			xHRS2 = Z_DoEncrypt(Request("hmon2") & "|" & Request("htue2") & "|" & Request("hwed2") & "|" & _
						Request("hthu2") & "|" & Request("hfri2") & "|" & Request("hsat2") & "|" & _
						Request("hsun2") & "|" & Request("Mnotes2") & "|" & Request("hmonX2") & "|" & Request("htueX2") & "|" & Request("hwedX2") & "|" & _
						Request("hthuX2") & "|" & Request("hfriX2") & "|" & Request("hsatX2") & "|" & _
						Request("hsunX2"))
			Session("MSG") = Session("MSG") & "<br>Please select a new consumer in week 2." 
			'Response.Redirect "view.asp?xHRS2=" & xHRS2
		End If
	End If
End If
'Response.redirect "view.asp"
Response.Redirect "view.asp?Con=" & Request("hdept1") & "&xHRS=" & xHRS & "&Con2=" & Request("hdept2") & "&xHRS2=" & xHRS2
%>