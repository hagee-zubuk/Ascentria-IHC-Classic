<%Language=VBScript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<%
	If Request("req") = 1 Then
		Set tblSite = Server.CreateObject("ADODB.RecordSet")
		sqlSite = "SELECT * FROM C_Site_Visit_Dates_t WHERE Medicaid_Number ='" & Request("MNum") & "' "
		tblSite.Open sqlSite, g_strCONN, 1, 3
		If Not tblSite.EOF Then
			If Request("ctr") <> "" Then 
				ctr = Request("ctr")
				For i = 0 to ctr 
					tmpctr = Request("chk" & i)
					If tmpctr <> "" Then
						strTmp = "index= " & tmpctr 
						tblSite.Movefirst
						tblSite.Find(strTmp)
						If Not tblSite.EOF Then
							tblSite.Delete
							tblSite.Update
						End If
					End If
				Next
			End If 
		End If
		tblSite.Close
		Set tblSite = Nothing
		Response.Redirect "A_C_Status.asp?MNum=" & Request("MNum")
	ElseIf Request("req") = 2 Then
		Set tblSite = Server.CreateObject("ADODB.RecordSet")
		sqlSite = "SELECT * FROM C_Diagnosis_t WHERE Medicaid_Number ='" & Request("MNum") & "' "
		tblSite.Open sqlSite, g_strCONN, 1, 3
		If Not tblSite.EOF Then
			If Request("ctr") <> "" Then 
				ctr = Request("ctr")
				For i = 0 to ctr 
					tmpctr = Request("chk" & i)
					If tmpctr <> "" Then
						strTmp = "index= " & tmpctr 
						tblSite.Movefirst
						tblSite.Find(strTmp)
						If Not tblSite.EOF Then
							tblSite.Delete
							tblSite.Update
						End If
					End If
				Next
			End If 
		End If
		tblSite.Close
		Set tblSite = Nothing
		Response.Redirect "A_C_Health.asp?MNum=" & Request("MNum")
	ElseIf Request("req") = 3 Then
		Set tblSite = Server.CreateObject("ADODB.RecordSet")
		sqlSite = "SELECT * FROM C_Site_Visit_Dates_t WHERE Medicaid_Number ='" & Request("MNum") & "' "
		tblSite.Open sqlSite, g_strCONN, 1, 3
		If Not tblSite.EOF Then
			If Request("ctr") <> "" Then 
				ctr = Request("ctr")
				For i = 0 to ctr 
					tmpctr = Request("chkSV" & i)
					If tmpctr <> "" Then
						strTmp = "index= " & tmpctr 
						tblSite.Movefirst
						tblSite.Find(strTmp)
						If Not tblSite.EOF Then
							tblSite.Delete
							tblSite.Update
						End If
					End If
				Next
			End If 
		End If
		tblSite.Close
		Set tblSite = Nothing
		Set tblSite = Server.CreateObject("ADODB.RecordSet")
		sqlSite = "SELECT * FROM C_Site_Visit_Dates_t WHERE Medicaid_Number ='" & Request("MNum") & "' "
		tblSite.Open sqlSite, g_strCONN, 1, 3
		If Not tblSite.EOF Then
			If Request("ctr2") <> "" Then 
				ctr = Request("ctr2")
				For i = 0 to ctr 
					tmpctr = Request("chkPC" & i)
					If tmpctr <> "" Then
						strTmp = "index= " & tmpctr 
						tblSite.Movefirst
						tblSite.Find(strTmp)
						If Not tblSite.EOF Then
							tblSite.Delete
							tblSite.Update
						End If
					End If
				Next
			End If 
		End If
		tblSite.Close
		Set tblSite = Nothing
		Set tblSite = Server.CreateObject("ADODB.RecordSet")
		sqlSite = "SELECT * FROM C_Site_Visit_Dates_t WHERE Medicaid_Number ='" & Request("MNum") & "' "
		tblSite.Open sqlSite, g_strCONN, 1, 3
		If Not tblSite.EOF Then
			If Request("ctr3") <> "" Then 
				ctr = Request("ctr3")
				For i = 0 to ctr 
					tmpctr = Request("chkMC" & i)
					If tmpctr <> "" Then
						strTmp = "index= " & tmpctr 
						tblSite.Movefirst
						tblSite.Find(strTmp)
						If Not tblSite.EOF Then
							tblSite.Delete
							tblSite.Update
						End If
					End If
				Next
			End If 
		End If
		tblSite.Close
		Set tblSite = Nothing
		
		Response.Redirect "log.asp?MNum=" & Request("MNum")
	ElseIf Request("req") = 4 Then
		Set tblSite = Server.CreateObject("ADODB.RecordSet")
		sqlSite = "SELECT * FROM C_Site_Visit_Dates_t WHERE Medicaid_Number ='" & Request("MNum") & "' "
		tblSite.Open sqlSite, g_strCONN, 1, 3
		If Not tblSite.EOF Then
			If Request("ctr4") <> "" Then 
				ctr = Request("ctr4")
				For i = 0 to ctr 
					tmpctr = Request("chkhosp" & i)
					If tmpctr <> "" Then
						strTmp = "index= " & tmpctr 
						tblSite.Movefirst
						tblSite.Find(strTmp)
						If Not tblSite.EOF Then
							tblSite.Delete
							tblSite.Update
						End If
					End If
				Next
			End If 
		End If
		tblSite.Close
		Set tblSite = Nothing
		Set tblSite = Server.CreateObject("ADODB.RecordSet")
		sqlSite = "SELECT * FROM C_Site_Visit_Dates_t WHERE Medicaid_Number ='" & Request("MNum") & "' "
		tblSite.Open sqlSite, g_strCONN, 1, 3
		If Not tblSite.EOF Then
			If Request("ctr5") <> "" Then 
				ctr = Request("ctr5")
				For i = 0 to ctr 
					tmpctr = Request("chksup" & i)
					If tmpctr <> "" Then
						strTmp = "index= " & tmpctr 
						tblSite.Movefirst
						tblSite.Find(strTmp)
						If Not tblSite.EOF Then
							tblSite.Delete
							tblSite.Update
						End If
					End If
				Next
			End If 
		End If
		tblSite.Close
		Set tblSite = Nothing
		Response.Redirect "a_C_Status.asp?MNum=" & Request("MNum")
	End If
%>