<%Language=VBScript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_security.asp" -->
<%
If UCase(Session("lngType")) = "0" Then
	Session("MSG") = "Invalid User Type. Please Sign In again."
	Response.Redirect "default.asp"
End If
%>
<html>
	<head>
		<title>LSS - In-Home Care - Administrator Tools</title>
		
	</head>
	<body bgcolor='white'  LEFTMARGIN='0' TOPMARGIN='0'>
		<!-- #include file="_boxup.asp" -->
		<!-- #include file="_NavHeader.asp" -->
		<br><br>
		<table border='0' align='center'>
			<tr>
				<td colspan='2' align='center'>
					<h3>ADMIN TOOLS</h3>
				</td>
			</tr>
			<tr>
				<td align='right'>
					<a href='activity.asp' style='text-decoration: none;'><font face='Trebuchet MS' Size='2'>[Activity Codes]</font></a>
				</td>
				<td>
					<font face='Trebuchet MS' Size='1'>- add/edit/delete activity codes for workers.</font>
				</td>
			</tr>
			<tr>
				<td align='right'>
					<a href='changeprime.asp' style='text-decoration: none;'><font face='Trebuchet MS' Size='2'>[Database Tools]</font></a>
				</td>
				<td>
					<font face='Trebuchet MS' Size='1'>- change SNN or medicaid of workers or consumers.</font>
				</td>
			</tr>
			<tr>
				<td align='right'>
					<a href='holiday.asp' style='text-decoration: none;'><font face='Trebuchet MS' Size='2'>[Holiday Tools]</font></a>
				</td>
				<td>
					<font face='Trebuchet MS' Size='1'>- set holidays for timesheet.</font>
				</td>
			</tr>
			<% If UCase(Session("lngType")) = "2" Then %>
			<tr>
				<td align='right'>
					<a href='Rate.asp' style='text-decoration: none;'><font face='Trebuchet MS' Size='2'>[Rate Tools]</font></a>
				</td>
				<td>
					<font face='Trebuchet MS' Size='1'>- change rate of consumers.</font>
				</td>
			</tr>
			<tr>
				<td align='right'>
					<a href='mileRate.asp' style='text-decoration: none;'><font face='Trebuchet MS' Size='2'>[Mileage Rate Tools]</font></a>
				</td>
				<td>
					<font face='Trebuchet MS' Size='1'>- change mileage rate of workers.</font>
				</td>
			</tr>
			<tr>
				<td align='right'>
					<a href='badge.asp'style='text-decoration: none;'><font face='Trebuchet MS' Size='2'>[Ultipro Tools]</font></a>
				</td>
				<td>
					<font face='Trebuchet MS' Size='1'>- converts ADP badge to Ultipro badge on punch files.</font>
				</td>
			</tr>
			<tr>
				<td align='right'>
					<a href='admin.asp'style='text-decoration: none;'><font face='Trebuchet MS' Size='2'>[User Tools]</font></a>
				</td>
				<td>
					<font face='Trebuchet MS' Size='1'>- add/edit/delete user accounts.</font>
				</td>
			</tr>
			<% End If %>
		</table>
	</td></tr>
		<!-- #include file="_boxdown.asp" -->
	</body>
</html>