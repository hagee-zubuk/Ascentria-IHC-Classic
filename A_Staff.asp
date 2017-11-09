<%Language=VBScript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<%
	If Session("SID") = "" Then
		tmpMCNum = split(Request("Staff"), " - ") 
		ID = tmpMCNum(0)
		Session("SID") = ID
	End If
	Set tblStaff = Server.CreateObject("ADODB.Recordset")
	sqlStaff = "SELECT * FROM NHSM_Staff_t WHERE Index = " & Session("SID") 
	tblStaff.Open sqlStaff, g_strCONN, 1, 3
		If Not tblStaff.EOF Then
			Index = tblStaff("Index")
			fname = tblStaff("Fname")
			lname = tblStaff("Lname")
		Else
			tblStaff.AddNew
			tblStaff("Index") = Session("SID")
			tblStaff.Update
		End If
	tblStaff.Close
	Set tblStaff = Nothing	
%>
<html>
	<head>
		<title>LSS - NH In-Home Care Staff Details</title>
		<script language='JavaScript'>
			function SD_Del()
			{
				var ans = window.confirm("Click OK to continue deletion of Staff. Click Cancel to stop.");
				if (ans){
				document.frmStafDet.action = "A_Del.asp?act=6";
				document.frmStafDet.submit();
				}
			}
		</script>
	</head>
	<body bgcolor='white'>
		<form method='post' name='frmStafDet' action='A_S_Action.asp'>
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
							<font size='2'><u>Name:</u>&nbsp;<input maxlength='20' style='font-size: 10px; height: 20px; width: 80px;' type='text' name='lname' value="<%=Lname%>">,&nbsp; 
							<input maxlength='20' style='font-size: 10px; height: 20px; width: 80px;' type='text' name='fname' value="<%=Fname%>">
							<font size='1'>(Last Name, First Name)</font>
							<input type='hidden' name='MCNum' value='<%=Index%>'>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
				</table>	
				<table border='0'>
					<tr>
						<td align='center' colspan='2'>
							<input type='submit' value='Save' style='width: 110px;' onclick='JavaScript:CSD_Edit();'>
							<input type='button' value='Delete' style='width: 110px;' onclick='JavaScript:SD_Del();'>
						</td>
					</tr>
				</table>
			</center>	
		</form>
	</body>
</html>