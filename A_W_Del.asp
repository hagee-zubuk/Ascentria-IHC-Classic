<%Language=VBScript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<%
Function GetWIndex(xxx)
	Set rsI = Server.CreateObject("ADODB.RecordSet")
	sqlI = "SELECT [index] FROM Worker_T WHERE Social_Security_Number = '" & xxx & "'"
	rsI.Open sqlI, g_strCONN, 3, 1
	If Not rsI.EOF Then
		GetWIndex = rsI("Index")
	End If
	rsI.Close
	Set rsI = Nothing
End Function
	response.write "PAGE: " & Request("page") 
	If Request("page") = 0 Then
		Set tblTowns = Server.CreateObject("ADODB.RecordSet")
		sqlTowns = "SELECT * FROM W_Towns_t WHERE SSN ='" & Request("WID") & "' "
		tblTowns.Open sqlTowns, g_strCONN, 1, 3
		If Not tblTowns.EOF Then
			If Request("ctr") <> "" Then 
				ctr = Request("ctr")
				For i = 0 to ctr 
					tmpctr = Request("chk" & i)
					If tmpctr <> "" Then
						strTmp = "index= " & tmpctr 
						tblTowns.Movefirst
						tblTowns.Find(strTmp)
						If Not tblTowns.EOF Then
							tblTowns.Delete
							tblTowns.Update
						End If
					End If
				Next
			End If 
		End If
		tblTowns.Close
		Set tblTowns = Nothing
	  Response.Redirect "A_Worker.asp?WID=" & Request("WID")
	ElseIf Request("page") = 1 Then
		Set tblTowns = Server.CreateObject("ADODB.RecordSet")
		sqlTowns = "SELECT * FROM ConWork_t WHERE WID ='" & GetWIndex(Request("WID")) & "' "
		tblTowns.Open sqlTowns, g_strCONN, 1, 3
		If Not tblTowns.EOF Then
			If Request("ctr") <> "" Then 
				ctr = Request("ctr")
				For i = 0 to ctr 
					tmpctr = Request("chk" & i)
					If tmpctr <> "" Then
						strTmp = "ID= " & tmpctr 
						tblTowns.Movefirst
						tblTowns.Find(strTmp)
						If Not tblTowns.EOF Then
							tblTowns.Delete
							tblTowns.Update
						End If
					End If
				Next
			End If 
		End If
		tblTowns.Close
		Set tblTowns = Nothing
		Response.Redirect "WorkCon.asp?WID=" & Request("WID")
	ElseIf Request("page") = 2 Then
		Set tblSite = Server.CreateObject("ADODB.RecordSet")
		sqlSite = "SELECT * FROM w_log_t WHERE SSN ='" & Request("WID") & "' "
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
		sqlSite = "SELECT * FROM W_log_t WHERE ssn ='" & Request("WID") & "' "
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
		sqlSite = "SELECT * FROM W_log_t WHERE ssn ='" & Request("WID") & "' "
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
		Response.Redirect "a_w_log.asp?WID=" & Request("WID")
	ElseIf Request("page") = 3 Then
		Set tblSite = Server.CreateObject("ADODB.RecordSet")
		sqlSite = "SELECT * FROM w_vio_t WHERE SSN ='" & Request("WID") & "' "
		tblSite.Open sqlSite, g_strCONN, 1, 3
		If Not tblSite.EOF Then
			If Request("ctr") <> "" Then 
				ctr = Request("ctr")
				For i = 0 to ctr 
					tmpctr = Request("chkwarn" & i)
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
		Response.Redirect "a_w_misc.asp?WID=" & Request("WID")
	End If
%>