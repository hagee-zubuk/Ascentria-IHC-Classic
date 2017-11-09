<%Language=VBScript%>
<!-- #include file="_Utils.asp" -->
<%
DIM DECname, DECid, DECdate

Session.Abandon
Response.Cookies("TSHEET").Expires = Now - 1
If request("enme") <> "" Then

	DECname = Z_DoDecrypt(Request("enme"))
	DECid = Z_DoDecrypt(Request("id"))
	DECdate = Z_DoDecrypt(Request("edte"))
	
	Session("vid") = DECid
	Session("vname") = DECname 
	Session("vdate") = DECdate

end if
%>
<html>
<head>
<link href="styles.css" type="text/css" rel="stylesheet" media="print">
<title>LSS - In-Home Care - Sign In</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"><style type="text/css">
<!--
body {
		background-color: #C4B464;
		filter:progid:DXImageTransform.Microsoft.Gradient(GradientType=0,StartColorStr='#c7c9e5',EndColorStr='#C4B464');
}
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
			table.bg {
			background-color: #FFFFFF;
			border-radius: 5px;
			}  
-->
</style></head>
<body onload="javascript: document.frmLog.UN.focus();">
<center>

<form method='post' name='frmLog' action='chkdata2.asp'>
	
  <p><br>
  </p>
  <p>&nbsp;</p>
  <p>&nbsp;</p>
  <p>&nbsp;</p>
  <p>&nbsp;</p>
  <p>&nbsp;</p>
  <p>&nbsp;</p>
  <p>&nbsp;</p>
  <p>&nbsp;</p>
  <table width='406' height='212' border='0' align="center" class="bg">
  	<tr><td align="center"><img src="images/printlogo.jpg" alt="In Home Care" border='0' /></td></tr>
	<tr><td width="400" align='center'><h5 align="center"><font face='trebuchet MS'>&nbsp;</font></h5>

  <table border='0' cellspacing='1' cellpadding='1'>
    <tr><td colspan='2' align='center'><span class='error'><font color='red'><%=Session("MSG")%></font></span></td></tr>
    <tr><td align='right'><font face='trebuchet MS' size='1'>USERNAME:</font></td
		><td width='100px'><input type='text' size='25' name='UN'></td
		></tr>
		    <tr><td align='right'><font face='trebuchet MS' size='1'>PASSWORD:</font></td
		><td width='100px'><input type='password' size='25' name='PW'></td
		></tr>
		<tr><td>&nbsp;</td></tr>
		<tr>
			<td colspan='2' align='right'>
				<input type='submit' style='width: 100px;' value='Sign In' class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='document.frmLog.submit();'>      
				<input type='reset' style='width: 100px;' value='Reset' class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'">      
			</td>
		</tr>
	</table>
</td>
	</tr></table>
</form>
</center>
</body>
</html>
<%
Session("MSG") = ""
%>