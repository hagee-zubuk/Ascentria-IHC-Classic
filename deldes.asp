<%@Language=VBScript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<%

DIM		lngI, tblEMP, strSQL, tmpID, ctrI, tmpctr

Set tblEMP = Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT * FROM [tdes_t]"
On Error Resume Next
tblEMP.Open strSQL, g_strCONN, 1, 3
ctrI = Request("countb")
	For i = 0 to ctrI
		tmpctr = Request("dchk" & i)
		If tmpctr <> "" Then
			strTmp = "index='" & tmpctr & "'"
			tblEMP.Find(strTmp)
			If Not tblEMP.EOF Then
				tblEMP.DELETE
				tblEMP.Update
			End If
		End If
		tblEMP.MoveFirst
	Next 
tblEmp.Close
Set tblEMP = Nothing
Response.Redirect "admin.asp"
%>