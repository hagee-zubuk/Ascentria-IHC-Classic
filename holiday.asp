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
'GET REG HOLIDAY
Set rsReg = Server.CreateObject("ADODB.RecordSet")
sqlReg = "SELECT * FROM RegHoliday_T ORDER BY month, day"
rsReg.Open sqlReg, g_strCONN, 1, 3
ctrReg = 0
Do Until rsReg.EOF
	strReg = strReg & "<tr><td><input type='checkbox' name='chkReg" & ctrReg & "' value='" & rsReg("index") & "'></td>" & _
		"<td><font size='1' face='trebuchet MS'>" & MonthName(rsReg("month"), True) & " / " & rsReg("day") & "</font></td></tr>"
	ctrReg = ctrReg + 1
	rsReg.MoveNext
Loop
rsReg.Close
Set rsReg = Nothing
'GET SPEC HOLIDAY
Set rsSpec = Server.CreateObject("ADODB.RecordSet")
sqlSpec = "SELECT * FROM SpecHoliday_T ORDER BY month, day"
rsSpec.Open sqlSpec, g_strCONN, 1, 3
ctrSpec = 0
Do Until rsSpec.EOF
	strSpec = strSpec & "<tr><td><input type='checkbox' name='chkSpec" & ctrSpec & "' value='" & rsSpec("index") & "'></td>" & _
		"<td><font size='1' face='trebuchet MS'>" & MonthName(rsSpec("month"), True) & " / " & rsSpec("day") & " / " & rsSpec("year") & "</font></td></tr>"
	ctrSpec = ctrSpec + 1
	rsSpec.MoveNext
Loop
rsSpec.Close
Set rsSpec = Nothing
%>
<html>
<head>
<link href="styles.css" type="text/css" rel="stylesheet" media="print">
<title>LSS - In-Home Care - Administrator Tools - Holiday Tools</title>
<SCRIPT LANGUAGE="JavaScript">
function DelMe()
{
	var ans = window.confirm("Delete Holiday?");
	if (ans)
	{
		document.frmADmin.action = "adminholiday.asp?ctrl=2";
		document.frmADmin.submit();
	}
}
function SaveMe()
{
	document.frmADmin.action = "adminholiday.asp?ctrl=1";
	document.frmADmin.submit();
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
	<font face='trebuchet MS' size='2' color='white'><b>Holiday Tools<b></font></td></tr>
<tr><td colspan='7' align='center'><span class='error'><font face='trebuchet MS' size='1' color='red'><%=Session("MSG")%></font></span></td
></tr>
<tr>
	<td colspan='1' align='center'>
		<font size='1' face='trebuchet MS'><u>Regular Holiday</u></font>
	</td>
	<td>&nbsp;</td>
	<td colspan='1' align='center'>
		<font size='1' face='trebuchet MS'><u>Special Holiday</u></font>
	</td>
</tr>
<tr>
	<td valign='top'>
		<table border='0'>
			<tr>
				<%=strReg%>
			</tr>
		</table>
	</td>
	<td>&nbsp;</td>
	<td valign='top'>
		<table border='0'>
			<tr>
				<%=strSpec%>
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
	<td>
		<select name='selMonth' style='font-size: 10px; height: 20px;'>
			<option value='0'>&nbsp;</option>
			<option value='1'>Jan</option>
			<option value='2'>Feb</option>
			<option value='3'>Mar</option>
			<option value='4'>Apl</option>
			<option value='5'>May</option>
			<option value='6'>Jun</option>
			<option value='7'>Jul</option>
			<option value='8'>Aug</option>
			<option value='9'>Sep</option>
			<option value='10'>Oct</option>
			<option value='11'>Nov</option>
			<option value='12'>Dec</option>
		</select>
		<select name='selDay' style='font-size: 10px; height: 20px;'>
			<option value='0'>&nbsp;</option>
			<option value='1'>1</option>
			<option value='2'>2</option>
			<option value='3'>3</option>
			<option value='4'>4</option>
			<option value='5'>5</option>
			<option value='6'>6</option>
			<option value='7'>7</option>
			<option value='8'>8</option>
			<option value='9'>9</option>
			<option value='10'>10</option>
			<option value='11'>11</option>
			<option value='12'>12</option>
			<option value='13'>13</option>
			<option value='14'>14</option>
			<option value='15'>15</option>
			<option value='16'>16</option>
			<option value='17'>17</option>
			<option value='18'>18</option>
			<option value='19'>19</option>
			<option value='20'>20</option>
			<option value='21'>21</option>
			<option value='22'>22</option>
			<option value='23'>23</option>
			<option value='24'>24</option>
			<option value='25'>25</option>
			<option value='26'>26</option>
			<option value='27'>27</option>
			<option value='28'>28</option>
			<option value='29'>29</option>
			<option value='30'>30</option>
			<option value='31'>31</option>
		</select>
	</td>
	<td bgcolor='#040C8B'>&nbsp;</td>
	<td>
	<select name='selMonth1' style='font-size: 10px; height: 20px;'>
			<option value='0'>&nbsp;</option>
			<option value='1'>Jan</option>
			<option value='2'>Feb</option>
			<option value='3'>Mar</option>
			<option value='4'>Apl</option>
			<option value='5'>May</option>
			<option value='6'>Jun</option>
			<option value='7'>Jul</option>
			<option value='8'>Aug</option>
			<option value='9'>Sep</option>
			<option value='10'>Oct</option>
			<option value='11'>Nov</option>
			<option value='12'>Dec</option>
		</select>
		<select name='selDay1' style='font-size: 10px; height: 20px;'>
			<option value='0'>&nbsp;</option>
			<option value='1'>1</option>
			<option value='2'>2</option>
			<option value='3'>3</option>
			<option value='4'>4</option>
			<option value='5'>5</option>
			<option value='6'>6</option>
			<option value='7'>7</option>
			<option value='8'>8</option>
			<option value='9'>9</option>
			<option value='10'>10</option>
			<option value='11'>11</option>
			<option value='12'>12</option>
			<option value='13'>13</option>
			<option value='14'>14</option>
			<option value='15'>15</option>
			<option value='16'>16</option>
			<option value='17'>17</option>
			<option value='18'>18</option>
			<option value='19'>19</option>
			<option value='20'>20</option>
			<option value='21'>21</option>
			<option value='22'>22</option>
			<option value='23'>23</option>
			<option value='24'>24</option>
			<option value='25'>25</option>
			<option value='26'>26</option>
			<option value='27'>27</option>
			<option value='28'>28</option>
			<option value='29'>29</option>
			<option value='30'>30</option>
			<option value='31'>31</option>
		</select>
		<input type='text' name='txtYear' style='font-size: 10px; height: 20px; width: 50px;' maxlength='4'>
	</td>
</tr>
<table border='0' align='center'>
<tr><td>&nbsp;</td></tr>	
<tr>
	<td align='center' colspan='3'>
		<input type='button' class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" style='width: 50px;' value="Save" onclick="javascript:SaveMe()">
		<input type='button' class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" style='width: 50px;' value="Delete" onclick="javascript:DelMe()">
		<input type='hidden' name='ctr1' value='<%=ctrReg%>'>
		<input type='hidden' name='ctr2' value='<%=ctrSpec%>'>
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