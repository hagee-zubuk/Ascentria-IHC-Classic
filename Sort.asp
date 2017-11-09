<%language=vbscript%>
<!-- #include file="_Files.asp" -->

<!-- #include file="_Utils.asp" -->
<%
Function activeako(xxx)
	activeako = False
	Set rsCM = Server.CreateOBject("Adodb.recordset")
	sqlCM = "SELECT * FROM CMCon_T WHERE CMID = '" & xxx & "'"
	rsCM.OPen sqlCM, g_StrCONN, 1, 3
	Do Until rsCm.EOF
		if activecons(rsCm("CID")) then 
			activeako = true
			exit function
		end if
		rscm.MoveNext
	Loop
	rscm.close
	set rscm = Nothing
End Function
function activecons(xxx)
	activecon = false
	Set rsCM = Server.CreateOBject("Adodb.recordset")
	sqlCM = "SELECT * FROM c_status_T WHERE Medicaid_Number = '" & xxx & "'"
	rsCM.OPen sqlCM, g_StrCONN, 1, 3
	If Not rscm.eof then
		activecons = rscm("active")
	end If
	rscm.close
	set rscm = Nothing	
end function
function inactiveako(xxx)
	inactiveako = true
	Set rsCM = Server.CreateOBject("Adodb.recordset")
	sqlCM = "SELECT * FROM CMCon_T WHERE CMID = '" & xxx & "'"
	rsCM.OPen sqlCM, g_StrCONN, 1, 3
	if rscm.eof then
		inactiveako = true
		exit function
	end if
	Do Until rsCm.EOF
		if activecons(rsCm("CID")) then 
			inactiveako = false
			exit function
		end if
		rscm.MoveNext
	Loop
	rscm.close
	set rscm = Nothing
end function
	If Request("type") = "con" Then	
		Set tblConsumer = Server.CreateObject("ADODB.Recordset")
		If Request("chk") = 1 Then sqlConsumer = "SELECT lname, fname, Medicaid_Number as mednum FROM Consumer_t ORDER BY Lname, fname ASC"
		If Request("chk") = 2 Then sqlConsumer = "SELECT lname, fname, Consumer_t.[Medicaid_Number] as mednum FROM Consumer_t, C_Status_t WHERE [Consumer_t].[Medicaid_Number] = " & _
			"[C_Status_t].[Medicaid_Number] AND Active = 1 ORDER BY Lname, fname ASC"
		If Request("chk") = 3 Then sqlConsumer = "SELECT lname, fname, Consumer_t.[Medicaid_Number] as mednum FROM Consumer_t, C_Status_t WHERE [Consumer_t].[Medicaid_Number] = " & _
			"[C_Status_t].[Medicaid_Number] AND Active = 0 ORDER BY Lname, fname ASC"
		tblConsumer.Open sqlConsumer, g_strCONN, 3, 1
			If Not tblConsumer.EOF Then
				Do Until tblConsumer.EOF
					If Request("chk") = 1 Then 
						tmp = tblConsumer("mednum")
					Else
						tmp = tblConsumer("mednum")
					End If
					If tmp <> "" Then
						tmpMN = tmp
						strConsumer = strConsumer & "<option value='" & tmpMN & "'>" &  tblConsumer("Lname") & ", " & tblConsumer("Fname") & "</option>"
						ctr = ctr + 1 
					End If
					tblConsumer.MoveNext
				Loop
			End If
		tblConsumer.Close
		Set tblConsumer = Nothing
		Session("strConsumer") = strConsumer 
		Response.Redirect "Admin2.asp?choice=Consumer&chk=" & Request("chk")
	ElseIf Request("type") = "wor" Then
		Set tblWorker = Server.CreateObject("ADODB.Recordset")
		If Request("chk") = 1 Then sqlWorker = "SELECT * FROM Worker_t ORDER BY Lname, fname ASC"
		If Request("chk") = 2 Then sqlWorker = "SELECT * FROM Worker_t WHERE STATUS = 'Active' ORDER BY Lname, fname ASC"
		If Request("chk") = 3 Then sqlWorker = "SELECT * FROM Worker_t WHERE STATUS = 'InActive' ORDER BY Lname, fname ASC"
		tblWorker.Open sqlWorker, g_strCONN, 3, 1
			If Not tblWorker.EOF Then
				Do Until tblWorker.EOF
					If tblWorker("Social_Security_Number") <> "" Then
						tmpMN = tblWorker("Social_Security_Number")
						strWorker = strWorker & "<option value='" & tmpMN & "'>" &  tblWorker("Lname") & ", " & tblWorker("Fname") & "</option>"
						ctr = ctr + 1
					End If
						tblWorker.MoveNext
				Loop
			End If
		tblWorker.Close
		Set tblWorker = Nothing
		Session("strWorker") = strWorker 
	   Response.Redirect "Admin2.asp?choice=Worker&chk=" & Request("chk")
	ElseIf Request("type") = "Rep" Then
		Set tblWorker = Server.CreateObject("ADODB.Recordset")
		If Request("chk") = 1 Then sqlWorker = "SELECT Representative_t.[index] as repID, Representative_t.[Lname] as replname, Representative_t.[Fname] as repFname FROM Representative_t, conrep_T, consumer_T WHERE RID = Representative_t.[index] " & _
				"AND CID = consumer_T.Medicaid_Number ORDER BY Representative_t.Lname, Representative_t.fname ASC"
		If Request("chk") = 2 Then sqlWorker = "SELECT Representative_t.[index] as repID, Representative_t.[Lname] as replname, Representative_t.[Fname] as repFname FROM Representative_t, conrep_T, consumer_T, C_status_T WHERE RID = Representative_t.[index] " & _
				"AND CID = consumer_T.Medicaid_Number AND consumer_T.Medicaid_Number = C_status_T.Medicaid_Number AND Active = 1 " & _
				"ORDER BY Representative_t.Lname, Representative_t.fname ASC"
		If Request("chk") = 3 Then sqlWorker = "SELECT Representative_t.[index] as repID, Representative_t.[Lname] as replname, Representative_t.[Fname] as repFname FROM Representative_t, conrep_T, consumer_T, C_status_T WHERE RID = Representative_t.[index] " & _
				"AND CID = consumer_T.Medicaid_Number AND consumer_T.Medicaid_Number = C_status_T.Medicaid_Number AND Active = 0 " & _
				"ORDER BY Representative_t.Lname, Representative_t.fname ASC"
			response.write "TEST: " & sqlWorker
			tblWorker.Open sqlWorker, g_strCONN, 3, 1
			If Not tblWorker.EOF Then
				Do Until tblWorker.EOF
					If tblWorker("repID") <> "" Then
						tmpMN = tblWorker("repID")
						strRep = strRep & "<option>" & tmpMN & " - " & tblWorker("replname") & ", " & tblWorker("repFname") & "</option>"
						ctr = ctr + 1
					End If
						tblWorker.MoveNext
				Loop
			End If
		tblWorker.Close
		Set tblWorker = Nothing
		Session("strRep") = strRep 
		Response.Redirect "Admin2.asp?choice=Rep&chk=" & Request("chk")
	ElseIf Request("type") = "case" Then
		Set tblWorker = Server.CreateObject("ADODB.Recordset")
		sqlWorker = "SELECT * FROM Case_Manager_t ORDER BY Lname, fname ASC"
		tblWorker.Open sqlWorker, g_strCONN, 3, 1
			If Not tblWorker.EOF Then
				Do Until tblWorker.EOF
					If tblWorker("index") <> "" Then
						If Request("chk") = 2 Then
							If activeako(tblWorker("index")) = true Then
								tmpMN = tblWorker("index")
								strRep = strRep & "<option>" & tmpMN & " - " & tblWorker("Lname") & ", " & tblWorker("Fname") & "</option>"
								
							End If
							response.write tblWorker("index") & " - " & activeako(tblWorker("index")) & "<br>"
						End If
						If Request("chk") = 3 Then
							If inactiveako(tblWorker("index")) = true Then
								response.write tblWorker("index") & " - " & inactiveako(tblWorker("index")) & "<br>"
								tmpMN = tblWorker("index")
								strRep = strRep & "<option>" & tmpMN & " - " & tblWorker("Lname") & ", " & tblWorker("Fname") & "</option>"
						End If
							
						End If
						If Request("chk") = 1 Then
							tmpMN = tblWorker("index")
							strRep = strRep & "<option>" & tmpMN & " - " & tblWorker("Lname") & ", " & tblWorker("Fname") & "</option>"
						End If
						ctr = ctr + 1
					End If
						tblWorker.MoveNext
				Loop
			End If
		tblWorker.Close
		Set tblWorker = Nothing
		Session("strCase") = strRep 
		Response.Redirect "Admin2.asp?choice=Case&chk=" & Request("chk")
	End If
%>