<%@Language=VBScript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<%

DIM		lngI, tblEMP, strSQL, tmpID, ctrI, tmpctr

Set tblEMP = Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT * FROM [Consumer_t]"
On Error Resume Next
tblEMP.Open strSQL, g_strCONN, 1, 3
ctrI = Request("countj")
	For i = 0 to ctrI - 1 
		tmpctr = Request("hchk" & i)
		If tmpctr <> "" Then
			strTmp = "medicaid_number='" & tmpctr & "'"
			tblEMP.Find(strTmp)
			If Not tblEMP.EOF Then
				tblEmp.DELETE
				tblEmp.Update
			End If
		End If
		tblEmp.MoveFirst
		
	Next
	 
tblEmp.Close
Set tblEMP = Nothing
Response.Redirect "admin.asp"
%>