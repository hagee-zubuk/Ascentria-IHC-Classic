<%@Language=VBScript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<%


DIM		lngType, tblEMP, strSQL, dtDATE, strLName, strFName, empid, strEUser, strEPass, tmpPW

Set tblEMP = Server.CreateObject("ADODB.Recordset")
strEUser = LCase(Request("UN"))
strSQL = "SELECT * FROM [Worker_t] WHERE [username] = '" & strEUser & "'"

tblEMP.Open strSQL, g_strCONN, 3, 1
Session("MSG") = "Invalid username or password"
If Not tblEMP.EOF Then
	tmpPW = Z_Dodecrypt(tblEMP("password"))
	if tmpPW = Request("PW") then
		If IsDate(Request("inDate")) Then
			Session("dtDate") = Request("inDATE")
		Else
			Session("dtDate") = Date
		End If
		Session("namel") = tblEMP("LName")
		Session("namef") = tblEMP("FName")
		Session("idemp") = tblEMP("Social_Security_number") 
		lngType = z_clng(tblEMP("type")) 	
		Session("MSG") = ""
		tblEmp.Close
		Set tblEMP = Nothing
		' Create the cookie
		Response.Cookies("TSHEET").Expires = Now + 0.125		' expire in 3 hours
		Response.Cookies("TSHEET")("User") = strEuser
		'Response.Cookies("TSHEET")("Pass") = strEpass
		'Response.Write "[VID]: " & Request("vid") & "<BR>"
		if Request("vid") <> "" then
			Response.Cookies("VTSHEETS")("VEid") = Request("vid")
			Response.Cookies("VTSHEETS")("VEname") = Request("vname")
			Response.Cookies("VTSHEETS")("VEdate") = Request("vdate")
			Response.Redirect "verify.asp"
		end if
		If lngType = 1 then 
			Response.redirect "admin2.asp"
		end if
		Response.Redirect "view.asp"
	end if
End if
On Error Resume Next


Response.Redirect "default.asp"
tblEmp.Close
Set tblEMP = Nothing
%>


