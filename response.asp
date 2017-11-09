<%@Language=VBScript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_security.asp" -->
<%

DIM		tblSTAT, strSQLs, lngI, tblEMP, strSQL, strTableScript, tmpID, ctrI, tmpctr, mlMail

'''''''''''''''Save Row
Set tblEMP = Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT * FROM [tsheets_t]"

tblEMP.Open strSQL, g_strCONN, 1, 3
on error resume next
ctrI = Request("count")
	For i = 0 to ctrI 
		tblEMP.Movefirst
		tmpctr = Request("chk" & i)
		If tmpctr <> "" Then
			strTmp = "ID='" & tmpctr & "' "
			tblEMP.Find(strTmp)
			If Not tblEMP.EOF Then
				Set tblCon = Server.CreateObject("ADODB.RecordSet")
				sqlCon = "SELECT * FROM Consumer_t WHERE medicaid_number = '" & Request("conid" & i) & "' "
				tblCon.Open sqlCon, g_strCONN, 1, 3
				If Not tblCon.EOF Then
					tmpMax = Z_CZero(tblCon("MaxHrs"))
				End If
				tblCon.Close
				Set tblCon = Nothing
				tmpHrs = Z_CZero(Request("hmon" & i)) + Z_CZero(Request("htue" & i)) + Z_CZero(Request("hwed" & i)) + _
						Z_CZero(Request("hthu" & i)) + Z_CZero(Request("hfri" & i)) + Z_CZero(Request("hsat" & i)) + _
						Z_CZero(Request("hsun" & i))
				If tmpHrs > tmpMax Then
					Session("MSG") = "Total hours for " &  Request("hdept" & i) & " is over the allowed hours."
					Response.Redirect "view.asp"
				End if		
				tblEMP("client") = Request("conid" & i)
				if Request("hmon" & i)= "" then
					tblEMP("mon") = 0
				else
					tblEMP("mon") = Request("hmon" & i)
				end if
				if Request("htue" & i)= "" then
					tblEMP("tue") = 0
				else
					tblEMP("tue") = Request("htue" & i)
				end if
				if Request("hwed" & i) = "" then
					tblEMP("wed") = 0
				else
					tblEMP("wed") = Request("hwed" & i)
				end if
				if Request("hthu" & i) = "" then
					tblEMP("thu") = 0
				else
					tblEMP("thu") = Request("hthu" & i)
				end if
				if Request("hfri" & i) = "" then
					tblEMP("fri") = 0
				else
					tblEMP("fri") = Request("hfri" & i)
				end if
				if Request("hsat" & i) = "" then
					tblEMP("sat") = 0
				else
					tblEMP("sat") = Request("hsat" & i)
				end if
				if Request("hsun" & i) = "" then
					tblEMP("sun") = 0
				else
						tblEMP("sun") = Request("hsun" & i)
					end if	
				tblEMP("date") = Request("1day")
				tblEMP("emp_id") = Session("idemp")
				tblEMP("misc_notes") = Request("Mnotes" & i)
			tblEMP.Update
			End If
		End If
	Next 
On Error Resume Next

'response.write Request("hdept") & "<-----------"
if not Request("hdept") = "" then
	tblEMP.AddNew
	Set tblCon = Server.CreateObject("ADODB.RecordSet")
	sqlCon = "SELECT * FROM Consumer_t WHERE medicaid_number = '" & Request("hdept") & "' "
	tblCon.Open sqlCon, g_strCONN, 1, 3
	If Not tblCon.EOF Then
		tmpMax = Z_CZero(tblCon("MaxHrs"))
	End If
	tblCon.Close
	Set tblCon = Nothing
	tmpHrs = Z_CZero(Request("hmon")) + Z_CZero(Request("htue")) + Z_CZero(Request("hwed")) + _
			Z_CZero(Request("hthu")) + Z_CZero(Request("hfri")) + Z_CZero(Request("hsat")) + _
			Z_CZero(Request("hsun"))
	If tmpHrs > tmpMax Then
		Session("MSG") = "Total hours for " &  Request("hdept") & " is over the allowed hours."
		Response.Redirect "view.asp"
	End if	
	If tmpHrs = 0 Then Response.Redirect "view.asp"
	tblEMP("client") = Request("hdept")

	if Request("hmon")= "" then
		tblEMP("mon") = 0
	else
		tblEMP("mon") = Request("hmon")
	end if
	if Request("htue")= "" then
		tblEMP("tue") = 0
	else
		tblEMP("tue") = Request("htue")
	end if
	if Request("hwed") = "" then
		tblEMP("wed") = 0
	else
		tblEMP("wed") = Request("hwed")
	end if
	if Request("hthu") = "" then
		tblEMP("thu") = 0
	else
		tblEMP("thu") = Request("hthu")
	end if
	if Request("hfri") = "" then
		tblEMP("fri") = 0
	else
		tblEMP("fri") = Request("hfri")
	end if
	if Request("hsat") = "" then
		tblEMP("sat") = 0
	else
		tblEMP("sat") = Request("hsat")
	end if
	if Request("hsun") = "" then
		tblEMP("sun") = 0
	else
		tblEMP("sun") = Request("hsun")
	end if	
	tblEMP("date") = Request("1day")
	tblEMP("emp_id") = Session("idemp") 
	tblEMP("misc_notes") = Request("Mnotes")
	tblEMP.UPDATE
	tblEmp.Close
	Set tblEMP = Nothing
	
	
	Set tblProc = Server.CreateObject("ADODB.RecordSet")
	sqlProc = "SELECT * FROM Process_t WHERE id ='" & Session("idemp") & "' AND d8 =# " & Request("1day") & " #"
	tblProc.Open sqlProc, g_strCONN, 1, 3
		If tblProc.EOF Then
		 	tblProc.AddNew
		 	tblProc("id") = Session("idemp")
		 	tblProc("d8") = Request("1day")
		 	tblProc.Update
	 	End If
	tblProc.Close
	Set tblProc = Nothing
End If
'''''''''''''''''''''''''''''
ENCdate = Z_DoEncrypt(Request("1day"))
ENCname = Z_DoEncrypt(Request("ename"))
ENCid = Z_DoEncrypt(Request("eid"))

on error resume next
	Set mlMail = CreateObject("CDONTS.NewMail")
    	mlMail.bodyformat = 0
    	mlMail.mailformat = 0
    	mlMail.From= "dev@zubuk.com"
    	'mlMail.To= request("email")
    	mlMail.To= "patrick@zubuk.com"
    	mlMail.Subject="Verify Timesheets"
	strSQL = "Website visited on: " & Now & "<BR><BR>" & vbCrLf & vbCrLf & _
			"From:" & Request("ename") & "<BR>" & vbCrLf & _
			"Date Submitted: " & date & "<BR>" & vbCrLf & _
			"Start of Timesheet: " & request("1day") & "<BR>" & vbCrLf & _
			"Link: <a href=""" & g_strURL & "default.asp?edte=" & ENCdate & "&enme=" & ENCname & "&id=" & ENCid & """ target=""_blank"">timesheet</a><BR><BR>" & vbCrLf & vbCrLf & _
			
			"<BR><font size='1' face='trebuchet MS'>(insert standard disclaimer here)</font>"
			'response.write strsql
	mlMail.Body = strSQL			
	mlMail.Send
	set mlMail=nothing

Set tblSTATUS = Server.CreateObject("ADODB.Recordset")
strSQLs = "SELECT * from [report_t] WHERE [empid] = '" & Request("eid")& "' AND [d8_sub] = #" & Request("1day") & "# "
tblSTATUS.Open strSQLs, g_strCONN, 1, 3
if tblSTATUS.EOF then
	tblSTATUS.Addnew
	tblSTATUS("tdes") = Request("tdes")
	tblSTATUS("stat") = "For Approval"
	tblSTATUS("d8_sub") = date
	tblSTATUS("empid") = Request("eid")
	tblSTATUS("thours") = Request("thrs")
	tblSTATUS("d8") = Request("1day")
else
	tblSTATUS("tdes") = Request("tdes")
	tblSTATUS("stat") = "For Approval"
	tblSTATUS("d8_sub") = date
	tblSTATUS("empid") = Request("eid")
	tblSTATUS("thours") = Request("thrs")
	tblSTATUS("d8") = Request("1day")
end if
tblSTATUS.Update

tblSTATUS.close
set tblSTATUS = Nothing
%>

<html>
<title>Timesheet - Response</title>
<link href="styles.css" type="text/css" rel="stylesheet" media="print">
<SCRIPT LANGUAGE="JavaScript">
function logout()
{
	document.frmRes.action = "default.asp";
	document.frmRes.submit();
}
function tsheet()
{
	document.frmRes.action = "view.asp";
	document.frmRes.submit();
}
</script>
</head>
<body bgcolor='white'  LEFTMARGIN='0' TOPMARGIN='0'>
<form name='frmRes'>
	<!-- #include file="_boxup.asp" -->
	<!-- #include file="_TopBanner.asp" -->
<table bgcolor='#A4CADB' align='left' border='0' width='100%'>	
			<tr >
				<td>
					<img src='images/Welcome.gif' border='0'>
					<a href='admin2.asp' ><img src='images/Home.gif' border='0'></a>
					<a href='view.asp' ><img src='images/Tsheet.gif' border='0'></a>
					<a href='default.asp' ><img src='images/SignOut.gif' border='0'></a>
				</td>
				
			</tr>
		</table>

	<center>
	
	<h2><br><br><font face='trebuchet MS'>Your timesheet has been sent for approval on </font></h2><br>
	<h3><font face='trebuchet MS'><%=now%></font>
	<hr><br>
		
	<font face='trebuchet MS'>Total hours for the week:</font></h3><h2>
		<%
		dim thrs
		if Request("thtot") = "" then
			thrs = 0
		else
			thrs = Request("thtot")
		end if
		Response.Write "<font face='trebuchet MS'>" & thrs  & " hrs.</font>" 
		
		%>
		<br>
		
		
		</h2>
<!-- #include file="_boxdown.asp" -->
</center>
</body>
</html>
