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
<title>LSS - Sign In</title>
</head>
<body bgcolor='#A4CADB'' onload="javascript: document.frmLog.UN.focus();">
<center>

<form method='post' name='frmLog' action='chkdata2.asp'>
	
<br>
<table border='1' width='500px' height='300px' bgcolor='white'>
	<tr><td align='center'>
		<img src='images/lssne.gif' border="0" >
 <br><br>
<p align='center'><h5><font face='trebuchet MS'>Sign In</font></h5></p>
<table border='0'>
<tr><td colspan='2' align='center'><span class='error'><%=Session("MSG")%></span></td></tr>
<tr><td align='right'><font face='trebuchet MS' size='1'>USERNAME:</font></td
><td width='100px'><input type='text' name='UN'></td
></tr>
<tr><td align='right'><font face='trebuchet MS' size='1'>PASSWORD:</font></td
><td width='100px'><input type='password' size='22%' name='PW' style="font-family:Times New Roman"></td
></tr>
</table>
<br>

<input style='background-image: url("images/submit.gif"); border: 0px; height: 30px; width: 90px;' type='submit' value=''>
<input style='background-image: url("images/reset.gif"); border: 0px; height: 30px; width: 90px;' type='reset' value=''>

<a href="help.asp" target="_blank"><p align='center' class='note'><font face='trebuchet MS' size='1'>help</font></p></a> 
<input type='hidden' name='vid' value='<%=DECid%>'>
<input type='hidden' name='vname' value='<%=DECname%>'>
<input type='hidden' name='vdate' value='<%=DECdate%>'>
</td></tr></table>
</form>



<font size='1' face='trebuchet MS'>Powered by:</font><br>
<img src='images/zubuk-gear.gif' width='80px' border='0'>
</center>
</body>
</html>
<%
Session("MSG") = ""
%>