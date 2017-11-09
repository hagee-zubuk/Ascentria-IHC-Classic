<%@Language=VBScript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->

<%
DIM		tblAD, strSQLAD, strTmp, lngI, tblEMP, strSQL, strTableScript, tmpID, ctrI, tmpctr

Set tblAD = Server.CreateObject("ADODB.Recordset")
strSQLAD = "SELECT * FROM [input_t] WHERE [Index] =" & Request("empID") 
tblAD.OPEN strSQLAD, g_strCONN, 1, 3
if tblAD.EOF then
	Session("MSG") = "Invalid ID."
	tblAD.Close
	set tblAD = Nothing
	Response.Redirect "verify.asp"
else
	Set tblrep = Server.CreateObject("ADODB.Recordset")
	sqlrep = "SELECT * from [report_t] WHERE empid = '" & Request("tmpID") & "' " & _
			" AND d8 = #" & Request("tmpd8") & "# "
			
	tblrep.OPEN sqlrep, g_strCONN, 1, 3
	'response.write sqlrep
			 
	Set tblEMP = Server.CreateObject("ADODB.Recordset")
	strSQL = "SELECT * FROM [tsheets_t]"
	tblEMP.Open strSQL, g_strCONN, 1, 3
	ctrI = Request("count")
	i = 0
	for i = 0 to ctrI 
	'do until i = ctrI + 1
		tmpctr = Request("tmpID" & i)
			If tmpctr <> "" Then
				strTmp = "ID='" & tmpctr & "'"
				'if not tblemp.bof then tblemp.movefirst
				tblEmp.MoveFirst
				tblEMP.Find(strTmp)
				Response.Write "Finding: " & strTmp & " |  tblEMP.EOF=" & tblEMP.EOF & "<BR>Req(app" & i & ")=" & Request("app" & i) & "<BR>"
				If NOt(tblEMP.EOF) Then
					Response.Write "REQ is? " & request("app" & i) & " --> "
					if request("app" & i) <> "" then
						Response.Write "Verify set to true<br>"
						tblEMP("verify") = true
						'tblrep("stat") = "Approved"	
					else
						Response.Write "Verify set to FALSE<br>"
						tblEMP("verify") = false
						'tblrep("stat") = "For Approval"	
					end if
					tblEmp.Update
					tblrep.Update
				End If
			
			End If
	next
	'loop
end if

tblEmp.Close
Set tblEMP = Nothing

tblrep.CLOSE
Set tblrep = NOTHING
Response.Redirect "verify.asp"
'else

'end if

%>

