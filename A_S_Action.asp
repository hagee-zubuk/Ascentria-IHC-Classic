<%Language=VBScript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<%
	Set tblStaff = Server.CreateObject("ADODB.Recordset")
	If Request("new") <> 1 Then
		sqlStaff = "SELECT * FROM NHSM_Staff_t WHERE Index = " & Session("SID") 
	Else
		sqlStaff = "SELECT * FROM NHSM_Staff_t "
	End If
	response.write "SQL:" & sqlStaff & vbCrLf
	tblStaff.Open sqlStaff, g_strCONN, 1, 3
		If Not tblStaff.EOF Then
			If Request("new") = 1 Then
				tblStaff.AddNew
				Session("SID") = tblStaff("Index")
			End If
		
			response.write "new:" & request("new")
			tblStaff("Lname") = Request("Lname")
			tblStaff("Fname") = Request("Fname")
			tblStaff.Update
		End If
	tblStaff.Close
	Set tblStaff = Nothing
	Response.Redirect "A_Staff.asp"
%>