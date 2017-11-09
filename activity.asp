<%Language=VBScript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_security.asp" -->
<%
If UCase(Session("lngType")) <> "2" AND UCase(Session("lngType")) <> "1" Then
	Session("MSG") = "Invalid User Type. Please Sign In again."
	Response.Redirect "default.asp"
End If
Session("adminka") = "true"
'GET CODES
Set rsReg = Server.CreateObject("ADODB.RecordSet")
sqlReg = "SELECT * FROM activity_T ORDER BY [Desc]"
rsReg.Open sqlReg, g_strCONN, 1, 3
ctrReg = 0
Do Until rsReg.EOF
	strReg = strReg & "<tr><td><input type='checkbox' name='chkReg" & ctrReg & "' value='" & rsReg("index") & "'></td>" & _
		"<td align='center'><font size='1' face='trebuchet MS'>" & rsReg("code") & "</font></td><td><font size='1' face='trebuchet MS'>" & rsReg("desc") & "</font></td></tr>"
	ctrReg = ctrReg + 1
	rsReg.MoveNext
Loop
rsReg.Close
Set rsReg = Nothing
%>
<html>
<head>
<link href="styles.css" type="text/css" rel="stylesheet" media="print">
<title>LSS - In-Home Care - Administrator Tools - Activity Codes</title>
<SCRIPT LANGUAGE="JavaScript">
function DelMe()
{
	var ans = window.confirm("Delete Activity?");
	if (ans)
	{
		document.frmADmin.action = "adminactivity.asp?ctrl=2";
		document.frmADmin.submit();
	}
}
function SaveMe()
{
	if (document.frmADmin.txtcode.value != "" && document.frmADmin.txtdesc.value != "")
	{
		document.frmADmin.action = "adminactivity.asp?ctrl=1";
		document.frmADmin.submit();
	}
	else
	{
		alert("ERROR: Code and Description are needed.");
	}
}
</script>
<style>
	Input.btn{
	font-size: 7.5pt;
	font-family: arial;
	color:#000000;
	font-weight:bolder;
	background-color:#d4d0c8;
	border:2px solid;
	text-align: center;
	border-top-color:#d4d0c8;
	border-left-color:#d4d0c8;
	border-right-color:#b6b3ae;
	border-bottom-color:#b6b3ae;
	filter:progid:DXImageTransform.Microsoft.Gradient(GradientType=0,StartColorStr='#ffffffff',EndColorStr='#d4d0c8');
}
INPUT.hovbtn{
	font-size: 7.5pt;
	font-family: arial;
	color:#000000;
	font-weight:bolder;
	background-color:#b6b3ae;
	border:2px solid;
	text-align: center;
	border-top-color:#b6b3ae;
	border-left-color:#b6b3ae;
	border-right-color:#d4d0c8;
	border-bottom-color:#d4d0c8;
	filter:progid:DXImageTransform.Microsoft.Gradient(GradientType=0,StartColorStr='#ffffffff',EndColorStr='#b6b3ae');
}  
</style>
</head>
<body bgcolor='white'  LEFTMARGIN='0' TOPMARGIN='0'>
		<!-- #include file="_boxup.asp" -->
		<!-- #include file="_NavHeader.asp" -->
		
		<br>
		<form method='post' name='frmADmin' >
<table border='0' align='center'>
<tr valign='center'><td valign='top'>
<table id='empT' align='center' bgcolor='white' border='0' cellpadding='0' cellspacing= '1' >
<tr bgcolor='#040C8B'><td border='0' colspan='7' align='center' style="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=0, startColorstr=#040C8B, endColorstr=#c7c9e5);">
	<font face='trebuchet MS' size='2' color='white'><b>Activity Codes</font></td></tr>
<tr><td colspan='7' align='center'><span class='error'><font face='trebuchet MS' size='1' color='red'><%=Session("MSG")%></font></span></td
></tr>
<tr>
	<td valign='top' colspan='3'>
		<table border='0'>
			<tr>
				<td>&nbsp;</td>
				<td align='center'><font size='1' face='trebuchet MS'><u>Code</u></font></td>
				<td align='center'><font size='1' face='trebuchet MS'><u>Description</u></font></td>
			</tr>
			<tr>
				<%=strReg%>
			</tr>
		</table>
	</td>
	
</tr>
<tr bgcolor='#040C8B'>
	<td colspan='3' align='center' style="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=0, startColorstr=#040C8B, endColorstr=#c7c9e5);">
		&nbsp;
	</td>
</tr>
<tr>
<td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
	<td>
		
		<input type='text' name='txtcode' style='font-size: 10px; height: 20px;' maxlength='2' size='2'>
	</td>
	<td>
		<input type='text' name='txtdesc' style='font-size: 10px; height: 20px; width: 200px;' maxlength='50'>
	</td>
</tr>
<table border='0' align='center'>
<tr><td>&nbsp;</td></tr>	
<tr>
	<td align='center' colspan='3'>
		<input type='button' class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" style='width: 50px;' value="Save" onclick="javascript:SaveMe()">
		<input type='button' class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" style='width: 50px;' value="Delete" onclick="javascript:DelMe()">
		<input type='hidden' name='ctr1' value='<%=ctrReg%>'>
	</td>
</tr>
<tr><td></td></tr>
</table>
</td></tr>
</table>
<br>
<center>

</center>
</form>
<!-- #include file="_boxdown.asp" -->
</body>
</html>
<%
Session("MSG") = ""
%>