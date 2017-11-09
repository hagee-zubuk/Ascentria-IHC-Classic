<%@Language=VBScript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<%

DIM		lngI, tblEMP, strSQL, tmpID, ctrI, tmpctr	

Set tblEMP = Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT * FROM [Consumer_t]"

tblEMP.Open strSQL, g_strCONN, 1, 3
ctrI = Request("countj")
	For i = 0 to ctrI 
		tmpctr = Request("hchk" & i)
		If tmpctr <> "" Then
			strTmp = "index= '" & tmpctr & "' "
			tblEMP.Movefirst
			tblEMP.Find(strTmp)
			If Not tblEMP.EOF Then
				tblEMP("Fname") = Request("h_dept" & i)
			tblEMP.Update
			End If
		End If
	Next 
	On Error Resume Next
	x = 0
	Do Until tblEMP.EOF
		if tblEMP("Fname") = Request("h_dept") then
			x = 1
		end if
		tblEMP.movenext
	Loop
	if x <> 1 then
		tblEMP.AddNew
		tblEMP("Fname") = Request("h_dept")
		tblEMP.Update
	else
		Session("MSG") = Request("h_dept") & " Already Exist!"
	end if
tblEmp.Close
Set tblEMP = Nothing
Response.Redirect "admin.asp"
%>
