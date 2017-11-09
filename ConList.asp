<%language=vbscript%>
<!-- #include file="_Files.asp" -->

<!-- #include file="_Utils.asp" -->
<%	
Function GetWIndex(xxx)
	Set rsWID = Server.CreateObject("ADODB.RecordSet")
	sqlWID = "SELECT [index] FROM Worker_T WHERE Social_Security_Number = '" & xxx & "'"
	rsWID.Open sqlWID, g_strCONN, 3, 1
	If Not rsWID.EOF Then
			GetWIndex = rsWID("Index")
	End If
	rsWID.CLose
	Set rsWID = Nothing
End Function
	If Request("del") <> 1 Then
		If Request("SelWor") <> "" Then
			Set rsWork = Server.CreateObject("ADODB.RecordSet")
			sqlWork = "SELECT * FROM [ConWork_t]"
			rsWork.Open sqlWork, g_strCONN, 1, 3
			rsWork.AddNew
			rsWork("WID") = GetWIndex(Request("SelWor"))
			rsWork("CID") = Request("MCnum")
			rsWork.Update
			rsWork.close
			Set rsWork = Nothing
		End If	
		If Request("SelCM") <> "" Then
			Set rsCM = Server.CreateObject("ADODB.RecordSet")
			sqlCM = "SELECT * FROM [CMCon_t]"
			rsCM.Open sqlCM, g_strCONN, 1, 3
			rsCM.AddNew
			rsCM("CMID") = Request("SelCM")
			rsCM("CID") = Request("MCnum")
			rsCM.Update
			rsCM.close
			Set rsCM = Nothing
		End If
		If Request("SelR") <> "" Then
			Set rsR = Server.CreateObject("ADODB.RecordSet")
			sqlR = "SELECT * FROM [ConRep_t]"
			rsR.Open sqlR, g_strCONN, 1, 3
			rsR.AddNew
			rsR("RID") = Request("SelR")
			rsR("CID") = Request("MCnum")
			rsR.Update
			rsR.close
			Set rsR = Nothing
		End If		
		If Request("SelWorBack") <> "" Then
			Set rsWork = Server.CreateObject("ADODB.RecordSet")
			sqlWork = "SELECT * FROM [ConWorkBack_t]"
			rsWork.Open sqlWork, g_strCONN, 1, 3
			rsWork.AddNew
			rsWork("WID") = Request("SelWorBack")
			rsWork("CID") = Request("MCnum")
			rsWork.Update
			rsWork.close
			Set rsWork = Nothing
		End If	
	Else
		
		If Request("chkCM0") <> "" Then
			Set rsCM = Server.CreateObject("ADODB.Recordset")
			sqlCM = "SELECT * FROM [CMCon_t] WHERE [CID] = '" & Request("MCnum") & "' AND [CMID] = '" & Request("chkCM0") & "' " 
			rsCM.Open sqlCM, g_strCONN, 1, 3
			If Not rsCM.EOF Then
				rsCM.Delete
				rsCM.Update
				'response.write sqlCM
			End If
			rsCM.Close
			Set rsCM = Nothing
		End If
		If Request("chkR0") <> "" Then
			Set rsR = Server.CreateObject("ADODB.Recordset")
			sqlR = "SELECT * FROM [ConRep_t] WHERE [CID] = '" & Request("MCnum") & "' AND [RID] = '" & Request("chkR0") & "' " 
			rsR.Open sqlR, g_strCONN, 1, 3
			If Not rsR.EOF Then
				rsR.Delete
				rsR.Update
				'response.write sqlCM
			End If
			rsR.Close
			Set rsR = Nothing
		End If
			'worker
			Set rsWork = Server.CreateObject("ADODB.Recordset")
			sqlWork = "SELECT * FROM [ConWork_t]"
			On Error Resume Next
			rsWork.Open sqlWork, g_strCONN, 1, 3
			ctrI = Request("ctr")
				For i = 0 to ctrI + 1 
					tmpctr = Request("chkW" & i)
					If tmpctr <> "" Then
						strTmp = "ID=" & tmpctr 
						rsWork.Find(strTmp)
						If Not rsWork.EOF Then
							rsWork.Delete
							rsWork.Update
						End If
					End If
					rsWork.MoveFirst
				Next
			rsWork.Close
			Set rsWork = Nothing
			'backup worker
			Set rsWork = Server.CreateObject("ADODB.Recordset")
			sqlWork = "SELECT * FROM [ConWorkBack_t]"
			On Error Resume Next
			rsWork.Open sqlWork, g_strCONN, 1, 3
			ctrI = Request("ctr2")
			response.write "CTR2: " & ctrI
				For i = 0 to ctrI + 1 
					tmpctr = Request("chkWBack" & i)
					response.write "tmpctr: " & tmpctr
					If tmpctr <> "" Then
						strTmp = "ID=" & tmpctr 
						rsWork.Find(strTmp)
						If Not rsWork.EOF Then
							rsWork.Delete
							rsWork.Update
						End If
					End If
					rsWork.MoveFirst
				Next
			rsWork.Close
			Set rsWork = Nothing
	End If
	
	Response.Redirect "A_Consumer.asp?MNum=" & Request("mcNum")
%>