<%Language=VBScript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_security.asp" -->
<%
If UCase(Session("lngType")) <> "2" Then
	Session("MSG") = "Invalid User Type. Please Sign In again."
	Response.Redirect "default.asp"
End If
DIM		tblEMP, strSQL, strTableScript, tblTYPE, tblHDEPT, strSQLT, strSQLH
DIM		tblLOC, strSQLl, tblJOB, strSQLj, tblSTAT, tblDES, strSQLs, strSQLd

Set tblEMP = Server.CreateObject("ADODB.Recordset")

strSQL = "SELECT * FROM [input_t] ORDER BY [lname]" 

tblEMP.Open strSQL, g_strCONN, 3, 1
'on error resume next
tblEMP.Movefirst
lngI = 0
Do While Not tblEMP.EOF
	if Z_IsOdd(lngI) = true then 
		kulay = "#FFFAF0" 
	else 
		kulay = "#FFFFFF"
	end if
	sc = ""
	fin = ""
	admin = ""
	if tblEMP("type") = 0 then sc = "selected"	
	if tblEMP("type") = 1 then fin = "selected"
	if tblEMP("type") = 2 then admin = "selected"
	pwd = Z_DoDecrypt(tblEMP("password"))
strTableScript = strTableScript & "<tr bgcolor='" & kulay & "'><td align='center'>" & vbCrLf & _
			"<input type='checkbox' name='chk" & lngI & "' value='" & tblEMP("index") & "'></td>" & _
			"<td width='85px' align='center'><input type='text' size='8' maxlength='25' name='l_name" & lngI & "' value='" & tblEMP("lname") & "'></td><td align='center'>" & _
			"<input type='text' size='7' name='f_name" & lngI & "'" & _
			"value='" & tblEMP("fname") & "' ></td" & _
			"><td width='85px' align='center'><input type='text' size='8' maxlength='25' name='u_name" & lngI & "' " & _
			"value='" & tblEMP("username") & "' ></td" & _
			"><td width='85px' align='center'><input type='password' size='8' maxlength='25' name='p_word" & lngI & "' value=""" & pwd & """></td>" & _ 
			"<td width='85px' align='center'><select name='seltype" & lngI & "'> " & _
			"<option value='0' " & sc & ">IHC</option>" & _
			"<option value='1' " & fin & ">finance</option>" & _
			"<option value='2' " & admin & ">administrator</option>" & _
			"</td></tr>" & vbCrLf
	a = ""		
	v = ""	
	
lngI = lngI + 1			
tblEMP.Movenext
loop
tblEMP.CLose
set tblEMP = Nothing

Session("adminka") = "true"
%>
<html>
<head>
<link href="styles.css" type="text/css" rel="stylesheet" media="print">
<title>LSS - In-Home Care - Administrator Tools - User Tools</title>
<SCRIPT LANGUAGE="JavaScript">
function deluser()
{
	document.frmADmin.action = "deluser.asp";
	document.frmADmin.submit();
}
function edituser()
{
	document.frmADmin.action = "edituser.asp";
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
	<font face='trebuchet MS' size='2' color='white'><b>User Tools<b></font></td></tr>
<tr><td colspan='7' align='center'><span class='error'><font face='trebuchet MS' size='1' color='red'><%=Session("MSG")%></font></span></td
></tr>
<tr class='title' bgcolor='white'><td width='25px'>&nbsp;</td
><td class='title' align='center'><font face='trebuchet MS' size='1'><u>Last Name</u></td
><td class='title' align='center'><font face='trebuchet MS' size='1'><u>First Name</u></td
><td class='title' align='center'><font face='trebuchet MS' size='1'><u>Username</u></td
><td class='title' align='center'><font face='trebuchet MS' size='1'><u>Password</u></td
><td class='title' align='center'><font face='trebuchet MS' size='1'><u>Type</u></td
></tr><tr><td>&nbsp;</td></tr>
<%=strTableScript%>
<tr class='info'><td width='35px' align='center'>&nbsp;</td
><td width='85px' align='center'><input type='textbox' size='7' maxlength='25' name='l_name'></td 
><td width='85px' align='center'><input type='textbox' size='7' maxlength='25' name='f_name'></td 
><td width='85px' align='center'><input type='textbox' size='7' maxlength='25' name='u_name'></td 
><td width='85px' align='center'><input type='textbox' size='7' maxlength='15' name='p_word'></td 
><td width='85px' align='center'><select name='seltype'>
		<option value='0'>IHC</option>
		<option value='1'>finance</option>
		<option value='2'>administrator</option>
	</td
></tr></table><td>

<input type='hidden' name='count' value='<%=lngI%>'>

<table border='0'>
<tr><td align='left'><input type='button' class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" style='width: 50px;' value="Save" onclick="javascript:edituser()"></td></tr>
<tr><td align='left'><input type='button' class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" style='width: 50px;' value="Delete" onclick="javascript:deluser()"></td></tr>
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