<%Language=VBScript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<%
'GET CODES
Set rsReg = Server.CreateObject("ADODB.RecordSet")
sqlReg = "SELECT * FROM activity_T ORDER BY [Desc]"
rsReg.Open sqlReg, g_strCONN, 1, 3
ctrReg = 0
Do Until rsReg.EOF
	if Z_IsOdd(ctrReg) = true then 
		kulay = "transparent" 
	else 
		kulay = "#F8F8FF"
	end if
	strReg = strReg & "<tr bgcolor='" & kulay & "'><td align='center'><font size='1' face='trebuchet MS'>" & rsReg("code") & "</font></td><td><font size='1' face='trebuchet MS'>" & rsReg("desc") & "</font></td></tr>"
	ctrReg = ctrReg + 1
	rsReg.MoveNext
Loop
rsReg.Close
Set rsReg = Nothing
%>
<html>
	<head>
		<title>LSS - In-Home Care - Activity Codes</title>
		</head>
		<body style="filter:progid:DXImageTransform.Microsoft.gradient(startColorstr=#FFFFFF, endColorstr=#C4B464);" bgcolor='#C4B464'>
			<table border='0'>
			<tr>
				<td align='center'><font size='1' face='trebuchet MS'><b><u>Code</u></b></font></td>
				<td align='center'><font size='1' face='trebuchet MS'><b><u>Description</u></b></font></td>
			</tr>
			<tr>
				<%=strReg%>
			</tr>
		</table>
		</body>
</html>