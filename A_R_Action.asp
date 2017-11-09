<%language=vbscript%>
<!-- #include file="_Files.asp" -->

<!-- #include file="_Utils.asp" -->
<%
	If Request("page") = "" Then
		Set tblRep = Server.CreateObject("ADODB.Recordset")
		If Request("new") <> 1 Then
			sqlRep = "SELECT * FROM Representative_t WHERE [Index] = " & Session("RID") 
		Else
			Session("RDet") = Z_DoEncrypt(Request("Lname") & "|" & Request("Fname") & "|" & Request("Addr") & "|" & Request("cty") & _
				"|" & Request("ste") & "|" & Request("zcode") & "|" & Request("PhoneNo") & "|" & Request("wPhoneNo") & "|" & Request("cPhoneNo") & "|" & Request("remail"))
			If Request("Lname") = "" Then
				Session("MSG") =  "Last Name required."
				Response.Redirect "A_New_Rep.asp"
			End If
			sqlRep = "SELECT * FROM Representative_t "
		End If
		response.write "SQL:" & sqlRep & vbCrLf
		tblRep.Open sqlRep, g_strCONN, 1, 3
			If Not tblRep.EOF Then
				If Request("new") = 1 Then
					tblRep.AddNew
					
				End If
			
				response.write "new:" & request("new")
				tblRep("Lname") = Request("Lname")
				tblRep("Fname") = Request("Fname")
				tblRep("Address") = Request("Addr")
				tblRep("City") = Request("cty")
				tblRep("State") = UCase(Request("ste"))
				tblRep("zip") = Request("zcode")
				tblRep("PhoneNo") = Request("PhoneNo")
				tblRep("wPhoneNo") = Request("wPhoneNo")
				tblRep("cPhoneNo") = Request("cPhoneNo")
				tblRep("email") = Request("remail")
						tblRep.Update
				Session("RID") = tblRep("Index")
		
			End If
		tblRep.Close
		Set tblRep = Nothing
		Response.Redirect "A_Rep.asp"
	ElseIf Request("page") = 1 Then
		If Request("SelCon") <> "" Then
			Set rsWork = Server.CreateObject("ADODB.RecordSet")
			sqlWork = "SELECT * FROM [ConRep_t] WHERE CID = '" & Request("SelCon") & "' "
			rsWork.Open sqlWork, g_strCONN, 1, 3
			If rsWork.EOF Then
				rsWork.AddNew
				rsWork("CID") = Request("SelCon")
				rsWork("RID") = Request("RID")
				rsWork.Update
			Else
				Session("MSG") = "Consumer selected already has a representative."
			End If
			rsWork.Close
			Set rsWork = Nothing
		End If	
		Response.Redirect "RepCon.asp?ID=" & Request("RID")
	ElseIf Request("page") = 2 Then	
		Set tblTowns = Server.CreateObject("ADODB.RecordSet")
		sqlTowns = "SELECT * FROM ConRep_t WHERE RID ='" & Request("RID") & "' "
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
		Response.Redirect "RepCon.asp?ID=" & Request("RID")
	End If
%>