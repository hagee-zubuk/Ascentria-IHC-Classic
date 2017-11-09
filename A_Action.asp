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
						tblNH("Fname") = Request("Fname" & i)
						tblNH("Lname") = Request("lname" & i)
						tblNH.Update
					End If
				End If
			Next 
		End If
		On Error Resume Next
		If Request("lname") <> "" Then
			x = 0
			tblNH.Movefirst
			Do Until tblNH.EOF
				If tblNH("Lname") = Request("lname") Then
					x = 1
				End If
				tblNH.MoveNext
			Loop
			If x <> 1 then
				tblNH.AddNew
				tblNH("Lname") = Request("lname")
				tblNH("Fname") = Request("Fname")
				tblNH.Update
			Else
				Session("MSG") = Request("lname") & " already exist!"
			End If
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
						tblRep("Fname") = Request("Fname" & i)
						tblRep("Lname") = Request("lname" & i)
						tblRep("Address") = Request("Addr" & i)
						tblRep("PhoneNo") = Request("Phone" & i)
						tblRep.Update
					End If
				End If
			Next 
		End If
		On Error Resume Next
		If Request("lname") <> "" Then
			x = 0
			tblRep.Movefirst
			Do Until tblRep.EOF
				If tblRep("Lname") = Request("lname") Then
					x = 1
				End If
				tblRep.MoveNext
			Loop
			If x <> 1 then
				tblRep.AddNew
				tblRep("Lname") = Request("lname")
				tblRep("Fname") = Request("Fname")
				tblRep("Address") = Request("Addr")
				tblRep("PhoneNo") = Request("Phone")
				tblRep.Update
			Else
				Session("MSG") = Request("lname") & " already exist!"
			End If
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
						tblCase("Fname") = Request("Fname" & i)
						tblCase("Lname") = Request("lname" & i)
						tblCase("Agency") = Request("Agency" & i)
						tblCase("Address") = Request("Add" & i)
						tblCase("OfficeNo") = Request("OPhone" & i)
						tblCase("CellNo") = Request("CPhone" & i)
						tblCase("FaxNo") = Request("FPhone" & i)
						tblCase.Update
					End If
				End If
			Next 
		End If
		On Error Resume Next
		If Request("lname") <> "" Then
			x = 0
			tblCase.Movefirst
			Do Until tblCase.EOF
				If tblCase("Lname") = Request("lname") Then
					x = 1
				End If
				tblCase.MoveNext
			Loop
			If x <> 1 then
				tblCase.AddNew
				tblCase("Fname") = Request("Fname")
				tblCase("Lname") = Request("lname")
				tblCase("Agency") = Request("Agency")
				tblCase("Address") = Request("Add")
				tblCase("OfficeNo") = Request("OPhone")
				tblCase("CellNo") = Request("CPhone")
				tblCase("FaxNo") = Request("FPhone")
				tblCase.Update
			Else
				Session("MSG") = Request("lname") & " already exist!"
			End If
		End If
		tblCase.Close
		Set tblCase = Nothing
		Response.Redirect "A_Choice.asp?choice=Case"
	End If
%>