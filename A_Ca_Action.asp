<%language=vbscript%>
<!-- #include file="_Files.asp" -->

<!-- #include file="_Utils.asp" -->
<%
	Set tblCase = Server.CreateObject("ADODB.Recordset")
	If Request("new") <> 1 Then
		sqlCase = "SELECT * FROM Case_Manager_t WHERE [Index] = " & Session("CaID") 
	Else
		Session("CaDet") = Z_DoEncrypt(Request("Lname") & "|" & Request("Fname") & "|" & Request("Addr") & "|" & Request("cty") & _
			"|" & Request("ste") & "|" & Request("zcode") & "|" & Request("Agency") & "|" & Request("PhoneNo") & "|" & _
			Request("CelNo") & "|" & Request("FaxNo") & "|" & Request("ext") & "|" & Request("eMail")) 
		If Request("Lname") = "" Then
			Session("MSG") =  "Last Name required."
			Response.Redirect "A_New_Case.asp"
		End If
		sqlCase = "SELECT * FROM Case_Manager_t "
	End If
	response.write "SQL:" & sqlCase & vbCrLf
	tblCase.Open sqlCase, g_strCONN, 1, 3
		If Not tblCase.EOF Then
			If Request("new") = 1 Then
				tblCase.AddNew
				
			End If
		
			response.write "new:" & request("new")
			tblCase("Lname") = Request("Lname")
			tblCase("Fname") = Request("Fname")
			'tblCase("Address") = Request("Addr")
			'tblCase("City") = Request("cty")
			'tblCase("State") = UCase(Request("ste"))
			'tblCase("zip") = Request("zcode")
			'tblCase("Agency") = Request("Agency")
			tblCase("OfficeNo") = Request("PhoneNo")
			tblCase("CellNo") = Request("CelNo")
			tblCase("FaxNo") = Request("FaxNo")
			tblCase("ext") = Request("ext")
			tblCase("eMail") = Request("eMail")
			tblCase("cmcid") = Request("selCMC")
			tblCase.Update
			
			Session("CaID") = tblCase("Index")
		End If
	tblCase.Close
	Set tblCase = Nothing
	Response.Redirect "A_Case.asp"
%>