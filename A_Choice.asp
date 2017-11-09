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
	Response.write Request("choice") & vbCrLf 
	If Request("choice") = "Worker" Then
		Session("type") = "Worker"
		Set tblWorker = Server.CreateObject("ADODB.Recordset")
		'sqlWorker = "SELECT * FROM Worker_t ORDER BY Lname ASC"
		sqlworker = "SELECT * FROM Worker_t WHERE STATUS = 'Active' ORDER BY Lname, Fname ASC"
		tblWorker.Open sqlWorker, g_strCONN, 3, 1
			If Not tblWorker.EOF Then
				Do Until tblWorker.EOF
					If tblWorker("Social_Security_Number") <> "" Then
						tmpMN = tblWorker("Social_Security_Number")
						strWorker = strWorker & "<option value='" & tmpMN & "'>" & tblWorker("Lname") & ", " & tblWorker("Fname") & "</option>"
						ctr = ctr + 1
					End If
						tblWorker.MoveNext
				Loop
			End If
		tblWorker.Close
		Set tblWorker = Nothing
		Session("strWorker") = strWorker 
	    Response.Redirect "Admin2.asp?choice=Worker"
	ElseIf Request("choice") = "Consumer" Then
		Session("type") = "Consumer"
		Set tblConsumer = Server.CreateObject("ADODB.Recordset")
		sqlconsumer = "SELECT Consumer_t.Medicaid_Number as mednum, lname, fname " & _
			"FROM Consumer_t, C_Status_t " & _
			"WHERE Consumer_t.Medicaid_Number = C_Status_t.Medicaid_Number " & _
			"AND Active = 1 " & _
			"ORDER BY Lname, Fname"
		tblConsumer.Open sqlConsumer, g_strCONN, 3, 1
			If Not tblConsumer.EOF Then
				Do Until tblConsumer.EOF
					If tblConsumer("mednum") <> "" Then
						tmpMN = tblConsumer("mednum")
						strConsumer = strConsumer & "<option value='" & tmpMN & "'>" & tblConsumer("Lname") & ", " & tblConsumer("Fname") & "</option>"
						ctr = ctr + 1 
					End If
					tblConsumer.MoveNext
				Loop
			End If
		tblConsumer.Close
		Set tblConsumer = Nothing
		Session("strConsumer") = strConsumer 
	    Response.Redirect "Admin2.asp?choice=Consumer"
	ElseIf Request("choice") = "Case" Then
		Session("type") = "Case"
		Set tblCase = Server.CreateObject("ADODB.Recordset")
		sqlCase = "SELECT * FROM Case_Manager_t ORDER BY Lname, fname ASC"
		tblCase.Open sqlCase, g_strCONN, 3, 1
			If Not tblCase.EOF Then
				Do Until tblCase.EOF
					If tblCase("Index") <> "" Then	
						If activeako(tblCase("index")) = true Then
								tmpMN = tblCase("index")
								strCase = strCase & "<option>" & tmpMN & " - " & tblCase("Lname") & ", " & tblCase("Fname") & "</option>"
								
							End If
						ctr = ctr + 1
					End If
					tblCase.MoveNext
				Loop
			End If
		tblCase.Close
		Set tblCase = Nothing
		Session("strCase") = strCase 
	    Response.Redirect "Admin2.asp?choice=Case"
	ElseIf Request("choice") = "Rep" Then
		Session("type") = "Rep"
		Set tblRep = Server.CreateObject("ADODB.Recordset")
		sqlRep = "SELECT Representative_t.[index] as repID, Representative_t.[Lname] as replname, Representative_t.[Fname] as repFname FROM Representative_t, conrep_T, consumer_T, C_status_T WHERE RID = Representative_t.[index] " & _
				"AND CID = consumer_T.Medicaid_Number AND consumer_T.Medicaid_Number = C_status_T.Medicaid_Number AND Active = 1 " & _
				"ORDER BY Representative_t.Lname, Representative_t.fname ASC"
		tblRep.Open sqlRep, g_strCONN, 3, 1
			If Not tblRep.EOF Then
				Do Until tblRep.EOF
					If tblRep("repID") <> "" Then
						tmpMN = tblRep("repID")
						strRep = strRep & "<option>" & tmpMN & " - " & tblRep("replname") & ", " & tblRep("repFname") & "</option>"
						ctr = ctr + 1
					End If
					tblRep.MoveNext
				Loop
			End If
		tblRep.Close
		Set tblRep = Nothing
		Session("strRep") = strRep
	    Response.Redirect "Admin2.asp?choice=Rep"
	ElseIf Request("choice") = "Staff" Then
		Session("type") = "Staff"
		Set tblStaff = Server.CreateObject("ADODB.Recordset")
		sqlStaff = "SELECT * FROM NHSM_Staff_t ORDER BY Lname, fname ASC"
		tblStaff.Open sqlStaff, g_strCONN, 3, 1
			If Not tblStaff.EOF Then
				Do Until tblStaff.EOF
					'If len(tblStaff("Index")) = 1 Then
					'	tmpMN = "0" & tblStaff("index")
					'Else
						tmpMN = tblStaff("Index")
					'End If
					strStaff = strStaff & "<option>" & tmpMN & " - " & tblStaff("Lname") & ", " & tblStaff("Fname") & "</option>"
					tblStaff.MoveNext
					ctr = ctr + 1
				Loop
			End If
		tblStaff.Close
		Set tblStaff = Nothing
		Session("strStaff") = strStaff
	   Response.Redirect "Admin2.asp?choice=Staff"
	Else
		Response.Redirect "Admin2.asp"
	End If
%>