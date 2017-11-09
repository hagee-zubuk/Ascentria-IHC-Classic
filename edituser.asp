<%@Language=VBScript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<%
Function CheckUName(strUname)
	CheckUname = False
	Set rsUser = Server.CreateObject("ADODB.RecordSet")
	sqlUser = "SELECT * FROM Input_T WHERE Upper([username]) = '" & UCase(strUname) & "' "
	rsUser.Open sqlUser, g_strCONN, 3, 1
	If Not rsUser.EOF Then CheckUname = True
	rsUser.Close
	Set rsUser = Nothing
End Function
Set tblEMP = Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT * FROM [input_t]"
tblEMP.Open strSQL, g_strCONN, 1, 3
ctrI = Request("count")
	For i = 0 to ctrI 
		tmpctr = Request("chk" & i)
		If tmpctr <> "" Then
			strTmp = "index= " & tmpctr
			tblEMP.MoveFirst
			tblEMP.Find(strTmp)
			If Not tblEMP.EOF Then
				If Request("l_name" & i) <> "" Then
					tblEMP("lname") = UCase(Request("l_name" & i))
				Else
					Session("MSG") = Session("MSG")  & "<br>Please supply a 'Last Name'."  
				End If
				If Request("f_name" & i) <> "" Then
					tblEMP("fname") = UCase(Request("f_name" & i))
				Else
					Session("MSG") = Session("MSG")  & "<br>Please supply a 'First Name'."  
				End If
				If Request("u_name" & i) <> "" Then
					If Request("u_name" & i) <> tblEMP("username") Then
						If CheckUname(Request("u_name" & i)) = False Then 
							tblEMP("username") = Request("u_name" & i)
						Else
							Session("MSG") = Session("MSG")  & "<br>'Username' already exists." 
						End If
					End If
				Else
					Session("MSG") = Session("MSG")  & "<br>Please supply a 'Username'."  
				End If
				If Request("p_word" & i) <> "" Then
					tblEMP("password") = Z_DoEncrypt(Request("p_word" & i))
				Else
					Session("MSG") = Session("MSG")  & "<br>Please supply a 'Password'."  
				End If
				tblEMP("type") =  request("seltype" & i) 
				
				If Session("MSG") = "" Then tblEmp.Update
			End If
		End If
	Next 
	'On Error Resume Next
	if Request("u_name") <> "" Then
		tblEMP.AddNew
		If Request("l_name") <> "" Then
			tblEMP("lname") = UCase(Request("l_name"))
		Else
			Session("MSG") = Session("MSG")  & "<br>Please supply a 'Last Name'."  
		End If
		If Request("u_name") <> "" Then
			tblEMP("fname") = UCase(Request("f_name"))
		Else
			Session("MSG") = Session("MSG")  & "<br>Please supply a 'First name'."  
		End If
		If Request("u_name") <> "" Then
			If CheckUname(Request("u_name")) = False Then 
				tblEMP("username") = Request("u_name")
			Else
				Session("MSG") = Session("MSG")  & "<br>'Username' already exists." 
			End If
		Else
			Session("MSG") = Session("MSG")  & "<br>Please supply a 'Username'."  
		End If
		If Request("p_word") <> "" Then
			tblEMP("password") = Z_DoEncrypt(Request("p_word"))
		Else
			Session("MSG") = Session("MSG")  & "<br>Please supply a 'Password'."  
		End If
		tblEMP("type") = Request("seltype") 
		If Session("MSG") = "" then tblEMP.Update
		End If
tblEmp.Close
Set tblEMP = Nothing
Response.Redirect "admin.asp"
%>
