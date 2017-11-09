<%Language=VBScript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<%
		If Request("add") <> 1 then
			Set tblEMP = Server.CreateObject("ADODB.Recordset")
			strSQL = "SELECT * FROM [C_Site_Visit_Dates_t] WHERE Medicaid_Number ='" & Request("MNum") & "' "
			'On Error Resume Next
			tblEMP.Open strSQL, g_strCONN, 1, 3
			ctr = Request("ctr")
				For i = 0 to ctr - 1 
					tmpctr = Request("chkSV" & i)
					If tmpctr <> "" Then
						strTmp = "Index= " & tmpctr 
						tblEMP.Find(strTmp)
						If Not tblEMP.EOF Then
							If IsDate(Request("txtSVD" & i)) Then 
								tblEmp("Site_V_Date") = Request("txtSVD" & i)
							Else
								Session("MSG") = "Invalid date in Site Visit log. " & tblEmp("Site_V_Date") & " changed to " & Request("txtSVD" & i) & ".<br>"
							End If
							tblEmp("Comments") = Request("Vcom" & i)
							'tblEmp("MiscSV") = Request("txtSVmisc" & i)
							tblEmp.Update
						End If
					End If
					tblEmp.MoveFirst
					
				Next
				 
			tblEmp.Close
			Set tblEMP = Nothing
			response.write "ctr: " & Request("ctr2") & "<br>"
			Set tblEMP2 = Server.CreateObject("ADODB.Recordset")
			strSQL2 = "SELECT * FROM [C_Site_Visit_Dates_t] WHERE Medicaid_Number ='" & Request("MNum") & "' "
			'On Error Resume Next
			tblEMP2.Open strSQL2, g_strCONN, 1, 3
			ctr = Request("ctr2")
				For i = 0 to ctr - 1 
					tmpctr = Request("chkPC" & i)
					If tmpctr <> "" Then
						strTmp = "Index= " & tmpctr 
						tblEMP2.Find(strTmp)
						If Not tblEMP2.EOF Then
							If IsDate(Request("txtPCD" & i)) Then 
								tblEmp2("PhoneCon_last") = Request("txtPCD" & i)
							Else
								Session("MSG") = Session("MSG") & "Invalid date in Phone Call log. " & tblEmp2("PhoneCon_last") & " changed to " & Request("txtPCD" & i) & ".<br>"
							End If
							tblEmp2("PCom") = Request("Pcom" & i)
							'tblEmp2("MiscPC") = Request("txtPCmisc" & i)
							tblEmp2.Update
						End If
					End If
					tblEmp2.MoveFirst
					
				Next
				 
			tblEmp2.Close
			Set tblEMP2 = Nothing
			Set tblEMP3 = Server.CreateObject("ADODB.Recordset")
			strSQL3 = "SELECT * FROM [C_Site_Visit_Dates_t] WHERE Medicaid_Number ='" & Request("MNum") & "' "
			'On Error Resume Next
			tblEMP3.Open strSQL3, g_strCONN, 1, 3
			ctr = Request("ctr3")
				For i = 0 to ctr - 1 
					tmpctr = Request("chkMC" & i)
					If tmpctr <> "" Then
						strTmp = "Index= " & tmpctr 
						tblEMP3.Find(strTmp)
						If Not tblEMP3.EOF Then
							If IsDate(Request("txtMCD" & i)) Then 
								tblEmp3("miscCon") = Request("txtMCD" & i)
							Else
								Session("MSG") = Session("MSG") & "Invalid date in Misc. Contacts  log. " & tblEmp3("misccon") & " changed to " & Request("txtMCD" & i) & ".<br>"
							End If
							tblEmp3("misc") = Request("MCon" & i)
							'tblEmp2("MiscPC") = Request("txtPCmisc" & i)
							tblEmp3.Update
						End If
					End If
					tblEmp3.MoveFirst
					
				Next
				 
			tblEmp3.Close
			Set tblEMP3 = Nothing
			
			'Set tblEMP4 = Server.CreateObject("ADODB.Recordset")
			'strSQL4 = "SELECT * FROM [C_Site_Visit_Dates_t] WHERE Medicaid_Number ='" & Request("MNum") & "' "
			'On Error Resume Next
			'tblEMP4.Open strSQL4, g_strCONN, 1, 3
			'ctr = Request("ctr4")
			'For i = 0 to ctr - 1 
			'	tmpctr = Request("chkhosp" & i)
			'	If tmpctr <> "" Then
			'		strTmp = "Index= " & tmpctr 
			'		tblEMP4.Find(strTmp)
			'		If Not tblEMP4.EOF Then
			'			If IsDate(Request("txthospdate" & i)) Then 
			'				tblEMP4("hospdate") = Request("txthospdate" & i)
			'			Else
			'				Session("MSG") = Session("MSG") & "Invalid date in Hospitalization log. " & tblEMP4("hospdate") & " changed to " & Request("txthospdate" & i) & ".<br>"
			'			End If
			'			tblEMP4("hospcom") = Request("hospcom" & i)
			'			tblEMP4.Update
			'		End If
			'	End If
			'	tblEMP4.MoveFirst
			'Next
			'tblEMP4.Close
			'Set tblEMP4 = Nothing
		If Request("txtSVD") <> "" Then
			If IsDate(Request("txtSVD")) Then
				Set tblSV = Server.CreateObject("ADODB.RecordSet")
				sqlSV = "SELECT * FROM [C_Site_Visit_Dates_t] WHERE [index] = 0 "
				tblSV.Open sqlSV, g_strCONN, 1, 3
				tblSV.AddNew
				tblSV("Medicaid_Number") = Request("MNum")
				tblSV("Site_V_Date") = Request("txtSVD")
				If Request("txtSVD") <> "" Then 
					tblSV("Comments") = Request("txtSVC")
				End If
				tblSV.Update
				tblSV.Close
				Set tblSV = Nothing
			Else
				Session("MSG") = "Invalid Entry in Site Visit.<br>" 
			End If
		End If
		If Request("txtPCD") <> "" Then
			If IsDate(Request("txtPCD")) Then
				Set tblPC = Server.CreateObject("ADODB.RecordSet")
				sqlPC = "SELECT * FROM [C_Site_Visit_Dates_t] WHERE [index] = 0 "
				tblPC.Open sqlPC, g_strCONN, 1, 3
				tblPC.AddNew
				tblPC("Medicaid_Number") = Request("MNum")
				tblPC("phoneCon_Last") = Request("txtPCD")
				If Request("txtPCD") <> "" Then 
					tblPC("PCom") = Request("txtPCC")
				End If
				tblPC.Update
				tblPC.Close
				Set tblPC = Nothing
			Else
				Session("MSG") = Session("MSG") & "Invalid Entry in Phone Calls."
			End If
		End If
		If Request("txtMCD") <> "" Then
			If IsDate(Request("txtMCD")) Then
				Set tblPC = Server.CreateObject("ADODB.RecordSet")
				sqlPC = "SELECT * FROM [C_Site_Visit_Dates_t] WHERE [index] = 0 "
				tblPC.Open sqlPC, g_strCONN, 1, 3
				tblPC.AddNew
				tblPC("Medicaid_Number") = Request("MNum")
				tblPC("misccon") = Request("txtMCD")
				If Request("txtMCD") <> "" Then 
					tblPC("misc") = Request("mcon")
				End If
				tblPC.Update
				tblPC.Close
				Set tblPC = Nothing
			Else
				Session("MSG") = Session("MSG") & "Invalid Entry in Misc. Contacts."
			End If
		End If
		'If Request("txtHospdate") <> "" Then
		'	If IsDate(Request("txtHospdate")) Then
		'		Set tblPC = Server.CreateObject("ADODB.RecordSet")
		'		sqlPC = "SELECT * FROM [C_Site_Visit_Dates_t] WHERE [index] = 0 "
		'		tblPC.Open sqlPC, g_strCONN, 1, 3
		'		tblPC.AddNew
		'		tblPC("Medicaid_Number") = Request("MNum")
		'		tblPC("hospdate") = Request("txtHospdate")
		'		If Request("hospcom") <> "" Then 
		'			tblPC("hospcom") = Request("hospcom")
		'		End If
		'		tblPC.Update
		'		tblPC.Close
		'		Set tblPC = Nothing
		'	Else
		'		Session("MSG") = Session("MSG") & "Invalid Entry in Hospitalizations."
		'	End If
		'End If
		Response.Redirect "Log.asp?MNum=" & Request("MNum")
	Else
		Set tblEMP = Server.CreateObject("ADODB.Recordset")
		strSQL = "SELECT * FROM [w_log_t]"
		'On Error Resume Next
		tblEMP.Open strSQL, g_strCONN, 1, 3
		ctr = Request("ctr")
			For i = 0 to ctr - 1 
				tmpctr = Request("chkSV" & i)
				If tmpctr <> "" Then
					response.write tmpctr & " ===> " & i
					strTmp = "Index= " & tmpctr
					tblEMP.Find(strTmp)
					If Not tblEMP.EOF Then
						If IsDate(Request("txtSVD" & i)) Then 
							tblEmp("sitev") = Request("txtSVD" & i)
						Else
							Session("MSG") = "Invalid date in Site Visit log. " & tblEmp("sitev") & " changed to " & Request("txtSVD" & i) & ".<br>"
						End If
						tblEmp("scom") = Request("Vcom" & i)
						
						tblEmp.Update
					End If
				End If
				tblEmp.MoveFirst
				
			Next
			 
		tblEmp.Close
		Set tblEMP = Nothing
		response.write "ctr: " & Request("ctr2") & "<br>"
		Set tblEMP2 = Server.CreateObject("ADODB.Recordset")
		strSQL2 = "SELECT * FROM [w_log_t]"
		'On Error Resume Next
		tblEMP2.Open strSQL2, g_strCONN, 1, 3
		ctr = Request("ctr2")
			For i = 0 to ctr - 1 
				tmpctr = Request("chkPC" & i)
				If tmpctr <> "" Then
					strTmp = "Index= " & tmpctr 
					tblEMP2.Find(strTmp)
					If Not tblEMP2.EOF Then
						If IsDate(Request("txtPCD" & i)) Then 
							tblEmp2("phonec") = Request("txtPCD" & i)
						Else
							Session("MSG") = Session("MSG") & "Invalid date in Phone Call log. " & tblEmp2("phonec") & " changed to " & Request("txtPCD" & i) & ".<br>"
						End If
						tblEmp2("PCom") = Request("Pcom" & i)
						
						tblEmp2.Update
					End If
				End If
				tblEmp2.MoveFirst
				
			Next
			 
		tblEmp2.Close
		Set tblEMP2 = Nothing
		
		Set tblEMP3 = Server.CreateObject("ADODB.Recordset")
		strSQL3 = "SELECT * FROM [w_log_t]"
		'On Error Resume Next
		tblEMP3.Open strSQL3, g_strCONN, 1, 3
		ctr = Request("ctr3")
			For i = 0 to ctr - 1 
				tmpctr = Request("chkMC" & i)
				If tmpctr <> "" Then
					strTmp = "Index= " & tmpctr 
					tblEMP3.Find(strTmp)
					If Not tblEMP3.EOF Then
						If IsDate(Request("txtMCD" & i)) Then 
							tblEmp3("misc") = Request("txtMCD" & i)
						Else
							Session("MSG") = Session("MSG") & "Invalid date in misc. contacts log. " & tblEmp3("misc") & " changed to " & Request("txtMCD" & i) & ".<br>"
						End If
						tblEmp3("mcom") = Request("MCon" & i)
						
						tblEmp3.Update
					End If
				End If
				tblEmp3.MoveFirst
				
			Next
			 
		tblEmp3.Close
		Set tblEMP3 = Nothing
		
	If Request("txtSVD") <> "" Then
		If IsDate(Request("txtSVD")) Then
			Set tblSV = Server.CreateObject("ADODB.RecordSet")
			sqlSV = "SELECT * FROM [w_log_t]"
			tblSV.Open sqlSV, g_strCONN, 1, 3
			tblSV.AddNew
			tblSV("ssn") = Request("wID")
			tblSV("sitev") = Request("txtSVD")
			If Request("txtSVD") <> "" Then tblSV("scom") = Request("txtSVC")
			tblSV.Update
			tblSV.Close
			Set tblSV = Nothing
		Else
			Session("MSG") = "Invalid Entry in Site Visit.<br>" 
		End If
	End If
	If Request("txtPCD") <> "" Then
		If IsDate(Request("txtPCD")) Then
			Set tblPC = Server.CreateObject("ADODB.RecordSet")
			sqlPC = "SELECT * FROM [w_log_t]"
			tblPC.Open sqlPC, g_strCONN, 1, 3
			tblPC.AddNew
			tblPC("ssn") = Request("wID")
			tblPC("phonec") = Request("txtPCD")
			If Request("txtPCD") <> "" Then tblPC("PCom") = Request("txtPCC")
			tblPC.Update
			tblPC.Close
			Set tblPC = Nothing
		Else
			Session("MSG") = Session("MSG") & "Invalid Entry in Phone Calls."
		End If
	End If
	If Request("txtMCD") <> "" Then
		If IsDate(Request("txtMCD")) Then
			Set tblMC = Server.CreateObject("ADODB.RecordSet")
			sqlMC = "SELECT * FROM [w_log_t]"
			tblMC.Open sqlMC, g_strCONN, 1, 3
			tblMC.AddNew
			tblMC("ssn") = Request("wID")
			tblMC("misc") = Request("txtMCD")
			If Request("txtMCD") <> "" Then tblMC("mcom") = Request("MCon")
			tblMC.Update
			tblMC.Close
			Set tblMC = Nothing
		Else
			Session("MSG") = Session("MSG") & "Invalid Entry in Phone Calls."
		End If
	End If
	Response.Redirect "a_W_Log.asp?WID=" & Request("WID")
End If
%>