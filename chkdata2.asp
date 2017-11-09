<%@Language=VBScript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<%


DIM		lngType, tbluser, strSQL, dtDATE, strLName, strFName, userid, strEUser, strEPass, tmpPW
on Error resume next
if request("cwork") <> 1 Then 
	Set tbluser = Server.CreateObject("ADODB.Recordset")
	strEUser = LCase(Request("UN"))
		
	strSQL = "SELECT * FROM [Input_t] WHERE [username] = '" & strEUser & "'"
	Session("MSG") = Session("MSG") & "Invalid username or password"
	tbluser.Open strSQL, g_strCONN, 3, 1
	'Session("MSG") = "Invalid username or password"
	If Not tbluser.EOF Then Session("lngType") = tbluser("type")
	If Not tbluser.EOF Then
		tmpPW = Z_Dodecrypt(tbluser("password"))
		session("UserID") = tbluser("index")
		Session("MSG") = Session("MSG") & "Invalid username or password" 
		if tmpPW = Request("PW") then
			Set fso = CreateObject("Scripting.FileSystemObject")
			Set ALog = fso.OpenTextFile(UserLog, 8, True)
			Alog.WriteLine Now & vbtab & "Successful Sign in :: User: " & tbluser("lname") & ", " & tbluser("fname")
			Set Alog = Nothing
			Set fso = Nothing 
			Set tblWork = Server.CreateObject("ADODB.RecordSet")
			sqlWork = "SELECT * FROM Worker_t WHERE Status = 'Active' ORDER BY lname, fname ASC"
			tblWork.Open sqlWork, g_strCONN, 1, 3
				if Not tblwork.eof Then
				tmpID = tblWork("Social_Security_Number")
				end if
			tblWork.Close
			sqlWork = "SELECT * FROM Worker_t WHERE Social_Security_Number = '" & tmpID & "' "
			tblWork.Open sqlWork, g_strCONN, 1, 3
			if Not tblwork.eof Then
				If IsDate(Request("indate")) Then
					Session("dtDate") = Request("indate")
				Else
					Session("dtDate") = Date
				End If
				Session("namel") = tblWork("LName")
				Session("namef") = tblWork("FName")
				Session("idemp") = tblWork("Social_Security_number") 
				'lngType = z_clng(tblWork("type")) 	
				Session("MSG") = ""
			End If
			tblWork.Close
			Set tblWork = Nothing
			' Create the cookie
			Response.Cookies("TSHEET").Expires = Now + 0.34		' expire in 8 hours
			Response.Cookies("TSHEET")("User") = strEuser
			'Response.Cookies("TSHEET")("Pass") = strEpass
			'Response.Write "[VID]: " & Request("vid") & "<BR>"
			
			if Request("vid") <> "" And lngType <> "False" then
				Response.Cookies("VTSHEETS")("VEid") = Request("vid")
				Response.Cookies("VTSHEETS")("VEname") = Request("vname")
				Response.Cookies("VTSHEETS")("VEdate") = Request("vdate")
				Response.Redirect "verify.asp"
			end if
			'If lngType <> "False" then 
			'	session("adminka") = "true"
			'	
			'	Response.redirect "admin.asp"
			'end if
			'Response.Redirect "view.asp"
			response.redirect "admin2.asp"
		end if
	End if
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set ALog = fso.OpenTextFile(UserLog, 8, True)
	Alog.WriteLine Now & vbtab & "Error in Sign in :: Username: " & strEuser
	Set Alog = Nothing
	Set fso = Nothing 		
	
	Response.Redirect "default.asp"
	tbluser.Close
	Set tbluser = Nothing
	response.write "<!-- Session: " & Request("vid") & " -->"
	response.write "<!-- lngtype: " & lngtype & " -->"
Else
	response.write "<!-- 2 -->"
	'Response.Cookies("TSHEET").Expires = Now + 0.125		' expire in 3 hours
	'		Response.Cookies("TSHEET")("User") = strEuser
	set tblWork = Server.CreateObject("ADODB.RecordSet")
	If Request("tmpID")  = "" Then
		sqlWork = "SELECT * FROM Worker_t WHERE Status = 'Active' order by lname"
	Else
		sqlWork = "SELECT * FROM Worker_t WHERE Social_Security_Number = " & Request("tmpID") 
	End If
	response.write "SQL:" & sqlwork
	tblWork.Open sqlWork, g_strCONN, 1, 3
	If IsDate(Request("indate")) Then
		Session("dtDate") = Request("indate")
	Else
		Session("dtDate") = Date
	End If
	Session("namel") = tblWork("LName")
	Session("namef") = tblWork("FName")
	Session("idemp") = tblWork("Social_Security_number") 

	'Session("idate") = 
	'lngType = z_clng(tblWork("type")) 	
	Session("MSG") = ""
	if Request("vid") <> "" then
			Response.Cookies("VTSHEETS")("VEid") = Request("vid")
			Response.Cookies("VTSHEETS")("VEname") = Request("vname")
			Response.Cookies("VTSHEETS")("VEdate") = Request("vdate")
			Response.Redirect "verify.asp"
		end if
	tblWork.Close
	Set tblWork = Nothing
	'response.write Request("indate") & "<br>" & Session("dtDate")
	Response.Redirect "view.asp"
End If
%>


