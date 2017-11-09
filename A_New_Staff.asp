<%Language=VBScript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<%
%>
<html>
	<head>
		<title>LSS - NH In-Home Care Staff Details</title>
		<script language='JavaScript'>
			function SDN_Edit()
			{
				document.frmStafDet.action = "A_S_Action.asp?new=1";
				document.frmStafDet.submit();
			}
		</script>
	</head>
	<body bgcolor='#F5F5F5'>
		<form method='post' name='frmStafDet' >
			<!-- #include file="_NavHeader.asp" -->
			<br>
			<center>
				<table border='0'>
					<tr>
						<td colspan='4' align='center' width='600px'>
							<font size='3'><b><u>Staff Details</u></b></font>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr><td>&nbsp;</td></tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td>
							<font size='2'><u>Name:</u>&nbsp;<input style='font-size: 10px; height: 20px; width: 80px;' type='text' maxlength='20' name='lname' value="<%=Lname%>">,&nbsp; 
							<input style='font-size: 10px; height: 20px; width: 80px;' type='text' maxlength='20' name='fname' value="<%=Fname%>">
							<font size='1'>(Last Name, First Name)</font>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
				</table>	
				<table border='0'>
					<tr>
						<td align='center' colspan='2'>
							<input type='button' value='Save' style='width: 110px;' onclick='JavaScript:SDN_Edit();'>
							
						</td>
					</tr>
				</table>
			</center>	
		</form>
	</body>
</html>