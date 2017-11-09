<%Language=VBScript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<%
	If Request("act") = 1 Then
		Set tblNH = Server.CreateObject("ADODB.RecordSet")
		sqlNH = "SELECT * FROM NHSM_Staff_t"
		tblNH.Open sqlNH, g_strCONN, 1, 3
		If Not tblNH.EOF Then
			ctr = Request("ctr")
			For i = 0 to ctr 
				tmpctr = Request("chk" & i)
				If tmpctr <> "" Then
					strTmp = "index= " & tmpctr 
					tblNH.Movefirst
					tblNH.Find(strTmp)
					If Not tblNH.EOF Then
						tblNH.Delete
						tblNH.Update
					End If
				End If
			Next 
		End If
		tblNH.Close
		Set tblNH = Nothing
		Response.Redirect "A_Choice.asp?choice=Staff"
	ElseIf Request("act") = 2 Then
		Set tblRep = Server.CreateObject("ADODB.RecordSet")
		sqlRep = "SELECT * FROM Representative_t"
		tblRep.Open sqlRep, g_strCONN, 1, 3
		If Not tblRep.EOF Then
			ctr = Request("ctr")
			For i = 0 to ctr 
				tmpctr = Request("chk" & i)
				If tmpctr <> "" Then
					strTmp = "index= " & tmpctr 
					tblRep.Movefirst
					tblRep.Find(strTmp)
					If Not tblRep.EOF Then
						tblRep.Delete
						tblRep.Update
					End If
				End If
			Next 
		End If
		tblRep.Close
		Set tblRep = Nothing
		Response.Redirect "A_Choice.asp?choice=Rep"
	ElseIf Request("act") = 3 Then
		Set tblCase = Server.CreateObject("ADODB.RecordSet")
		sqlCase = "SELECT * FROM Case_Manager_t"
		tblCase.Open sqlCase, g_strCONN, 1, 3
		If Not tblCase.EOF Then
			ctr = Request("ctr")
			For i = 0 to ctr 
				tmpctr = Request("chk" & i)
				If tmpctr <> "" Then
					strTmp = "index= " & tmpctr 
					tblCase.Movefirst
					tblCase.Find(strTmp)
					If Not tblCase.EOF Then
						tblCase.Delete
						tblCase.Update
					End If
				End If
			Next 
		End If
		tblCase.Close
		Set tblCase = Nothing
		Response.Redirect "A_Choice.asp?choice=Case"
	ElseIf Request("act") = 4 Then
		Set tblConsumer = Server.CreateObject("ADODB.RecordSet")
		sqlConsumer = "DELETE FROM Consumer_t WHERE Medicaid_Number ='" & Session("CID") & "' "
		tblConsumer.Open sqlConsumer, g_strCONN, 1, 3
		sqlConsumer = "DELETE FROM C_Diagnosis_t WHERE Medicaid_Number ='" & Session("CID") & "' "
		tblConsumer.Open sqlConsumer, g_strCONN, 1, 3
		sqlConsumer = "DELETE FROM C_Files_t WHERE Medicaid_Number ='" & Session("CID") & "' "
		tblConsumer.Open sqlConsumer, g_strCONN, 1, 3
		sqlConsumer = "DELETE FROM C_health_t WHERE Medicaid_Number ='" & Session("CID") & "' "
		tblConsumer.Open sqlConsumer, g_strCONN, 1, 3
		sqlConsumer = "DELETE FROM C_Site_Visit_Dates_t WHERE Medicaid_Number ='" & Session("CID") & "' "
		tblConsumer.Open sqlConsumer, g_strCONN, 1, 3
		sqlConsumer = "DELETE FROM C_Status_t WHERE Medicaid_Number ='" & Session("CID") & "' "
		tblConsumer.Open sqlConsumer, g_strCONN, 1, 3
		sqlConsumer = "DELETE FROM ConWork_t WHERE CID ='" & Session("CID") & "' "
		tblConsumer.Open sqlConsumer, g_strCONN, 1, 3
		sqlConsumer = "DELETE FROM ConRep_t WHERE CID ='" & Session("CID") & "' "
		tblConsumer.Open sqlConsumer, g_strCONN, 1, 3
		sqlConsumer = "DELETE FROM CMCon_t WHERE CID ='" & Session("CID") & "' "
		tblConsumer.Open sqlConsumer, g_strCONN, 1, 3
		Set tblConsumer = Nothing
		Response.Redirect "A_Choice.asp?choice=Consumer"
	ElseIf Request("act") = 5 Then
		Set tblWorker = Server.CreateObject("ADODB.RecordSet")
		strSQL = "DELETE FROM Worker_t WHERE Social_Security_Number = '" & Session("WID") & "' "
		tblWorker.Open strSQL, g_strCONN, 1, 3
		'tblWorker.Close
		strSQL = "DELETE FROM W_Files_t WHERE SSN = '" & Session("WID") & "' "
		tblWorker.Open strSQL, g_strCONN, 1, 3
		'tblWorker.Close
		strSQL = "DELETE FROM W_Towns_t WHERE SSN = '" & Session("WID") & "' "
		tblWorker.Open strSQL, g_strCONN, 1, 3
		'tblWorker.Close
		'strSQL = "DELETE FROM Process_t WHERE Wor = '" & Session("WID") & "' "
		'tblWorker.Open strSQL, g_strCONN, 1, 3
		'tblWorker.Close
		'strSQL = "DELETE FROM Report_t WHERE empid = '" & Session("WID") & "' "
		'tblWorker.Open strSQL, g_strCONN, 1, 3
		'tblWorker.Close
		'strSQL = "DELETE FROM Tsheets_t WHERE emp_id = '" & Session("WID") & "' "
		'tblWorker.Open strSQL, g_strCONN, 1, 3
		strSQL = "DELETE FROM ConWork_t WHERE WID = '" & Session("WID") & "' "
		tblWorker.Open strSQL, g_strCONN, 1, 3
		'tblWorker.Close
		Set tblWorker = Nothing
		're-create worker list
		Set fso = CreateObject("Scripting.FileSystemObject")
		'If Not fso.FileExists(WorkerList) Then
			Set NewWorkList = fso.CreateTextFile(WorkerList, true)
			Set tblLWork = Server.CreateObject("ADODB.Recordset")
			strSQLd = "SELECT * FROM [Worker_t] WHERE Status = 'Active' OR Status = 'InActive' ORDER BY [lname], [Fname]"
			tblLWork.Open strSQLd, g_strCONN, 3, 1
			tblLWork.Movefirst
			Do Until tblLWork.EOF
				NewWorkList.WriteLine tblLWork("index") & "|<option value='" & tblLWork("index")& "'> "& tblLWork("Lname") & ", " & tblLWork("fname") & " </option>"
				tblLWork.Movenext
			Loop
			tblLWork.Close
			Set tblLWork = Nothing
			Set NewWorkList = Nothing
			set fso = nothing
		'End If
		Response.Redirect "A_Choice.asp?choice=Worker"
	ElseIf Request("act") = 6 Then
		Set tblStaff = Server.CreateObject("ADODB.RecordSet")
		sqlStaff = "SELECT * FROM NHSM_Staff_t WHERE Index =" & Session("SID") 
		tblStaff.Open sqlStaff, g_strCONN, 1, 3
		If Not tblStaff.EOF Then
			tblStaff.Delete
			tblStaff.Update
		End If
		tblStaff.Close
		Set tblStaff = Nothing
		Response.Redirect "A_Choice.asp?choice=Worker"
	ElseIf Request("act") = 7 Then
		Set tblRep = Server.CreateObject("ADODB.RecordSet")
		sqlRep = "SELECT * FROM Representative_t WHERE Index =" & Session("RID") 
		tblRep.Open sqlRep, g_strCONN, 1, 3
		If Not tblRep.EOF Then
			tblRep.Delete
			tblRep.Update
		End If
		tblRep.Close
		Set tblRep = Nothing
		Set rsCon = Server.CreateObject("ADODB.RecordSet")
		sqlCon = "SELECT * FROM [ConRep_t] WHERE RID = '" & Session("RID") & "' "
		rsCon.Open sqlCon, g_strCONN, 1, 3
		If Not rsCon.EOF Then
			Do Until rsCon.EOF
				rsCon.Delete
				rsCon.Update
				rsCon.MoveNext
			Loop
		End If
		rsCon.Close
		Set rsCon = Nothing
		Response.Redirect "A_Choice.asp?choice=Rep"
	ElseIf Request("act") = 8 Then
		Set tblCase = Server.CreateObject("ADODB.RecordSet")
		sqlCase = "SELECT * FROM Case_Manager_t WHERE [Index] = " & Session("CaID") 
		tblCase.Open sqlCase, g_strCONN, 1, 3
		If Not tblCase.EOF Then
			tblCase.Delete
			tblCase.Update
		End If
		tblCase.Close
		Set tblCase = Nothing
		Set rsCon = Server.CreateObject("ADODB.RecordSet")
		sqlCon = "SELECT * FROM [CMCon_t] WHERE CMID = '" & Session("CaID")  & "' "
		rsCon.Open sqlCon, g_strCONN, 1, 3
		If Not rsCon.EOF Then
			Do Until rsCon.EOF
				rsCon.Delete
				rsCon.Update
				rsCon.MoveNext
			Loop
		End If
		rsCon.Close
		Set rsCon = Nothing
		Response.Redirect "A_Choice.asp?choice=Case"
	End If
%>