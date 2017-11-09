<%Language=VBScript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_security.asp" -->
<%
	'response.write "USER:" & m_User
	chk1 = ""
	chk2 = "CHECKED"
	chk3 = ""
	If Request("chk") <> "" Then
		chk1 = ""
		chk2 = ""
		chk3 = ""
		chkr = Request("chk")
		If chkr = 1 Then
			chk1 = "CHECKED"
		ElseIf chkr = 2 Then
			chk2 = "CHECKED"
		ElseIf chkr = 3 Then
			chk3 = "CHECKED"
		End If	
	End If
%>
<html>
	<head>
		<title>LSS - In-Home Care - Home</title>
		<script language="JavaScript">
			function C_Go()
			{
				document.frmADmin.action = "A_Consumer.asp";
				document.frmADmin.submit();
			}
			function W_Go()
			{
				document.frmADmin.action = "A_Worker.asp";
				document.frmADmin.submit();
			}
			function S_Go()
			{
				document.frmADmin.action = "A_Staff.asp";
				document.frmADmin.submit();
			}
			function R_Go()
			{
				document.frmADmin.action = "A_Rep.asp";
				document.frmADmin.submit();
			}
			function Ca_Go()
			{
				document.frmADmin.action = "A_Case.asp";
				document.frmADmin.submit();
			}
			function ConSort()
			{
				document.frmADmin.action = "Sort.asp";
				document.frmADmin.submit();
			}
			function PopMe(zzz)
			{
				newwindow = window.open(zzz,'name','height=600,width=550,scrollbars=1,directories=0,status=0,toolbar=0,resizable=0');
				if (window.focus) {newwindow.focus()}
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
	<body bgcolor='#FFFFFF' LEFTMARGIN='0' TOPMARGIN='0'>
		<form method='post' name='frmADmin' action='Admin2.asp'>
			<!-- #include file="_boxup.asp" -->
			<!-- #include file="_NavHeader.asp" --> 
			<% If Request("choice") = "Staff" Then %>
			<center>
				<table width="100%">
					<tr>
						<td colspan='3' align='center'><font size='3'><u><b>NH In-Home Care STAFF<b></u></font></td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td align='center'>
							<select name='staff' style='font-size: 10pt;'>
								<%=Session("strStaff")%>
							</select>
						</td>
						<td align='left'>
							<input type='button' value='GO' onclick='JavaScript:S_Go();'>
						</td>
						<td align='left'>
							<a href='A_New_Staff.asp' style='text-decoration:none'><font size='1'>[New]</font></a>
						</td>
					</tr>
				</table> 
			<% ElseIf Request("choice") = "Rep" Then %>
				<center>
				<table>
					<tr>
						<td colspan='3' align='center'><font size='2' face='trebuchet MS'><u><b>REPRESENTATIVE<b></u></font>
								<font size='2' face='trebuchet MS'>&nbsp;&nbsp;All:</font><input type='radio' <%=chk1%> name='chk' value='1' onclick='JavaScript:ConSort();'>
							
								<font size='2' face='trebuchet MS'>&nbsp;&nbsp;Active:</font><input type='radio' <%=chk2%> name='chk' value='2' onclick='JavaScript:ConSort();'>
						
								<font size='2' face='trebuchet MS'>&nbsp;&nbsp;Inactive:</font><input type='radio' <%=chk3%> name='chk' value='3' onclick='JavaScript:ConSort();'>
								<input type='hidden' name='type' value='Rep'>
							
							</td>
				
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td align='center'>
							<select name='Rep' style='font-size: 10pt; font-face: trebuchet MS;'>
								<%=Session("strRep")%>
							</select>
						</td>
						<td align='left'>
							<input type='button' value='GO' class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='JavaScript:R_Go();'>
						</td>
						<td align='left'>
							<a href='A_New_Rep.asp' style='text-decoration:none'><font size='1' face='trebuchet MS'>[New]</font></a>
						</td>
					</tr>
				</table> 
			</center>
			<% ElseIf Request("choice") = "Case" Then %>
				<center>
				<table>
					<tr>
						<td colspan='3' align='center'><font size='2' face='trebuchet MS'><u><b>CASE MANAGER<b></u></font>
						<font size='2' face='trebuchet MS'>&nbsp;&nbsp;All:</font><input type='radio' <%=chk1%> name='chk' value='1' onclick='JavaScript:ConSort();'>
							
								<font size='2' face='trebuchet MS'>&nbsp;&nbsp;Active:</font><input type='radio' <%=chk2%> name='chk' value='2' onclick='JavaScript:ConSort();'>
						
								<font size='2' face='trebuchet MS'>&nbsp;&nbsp;Inactive:</font><input type='radio' <%=chk3%> name='chk' value='3' onclick='JavaScript:ConSort();'>
								<input type='hidden' name='type' value='case'>	
						</td>
						
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td align='center'>
							<select name='case' style='font-size: 10pt;'>
								<%=Session("strCase")%>
							</select>
						</td>
						<td align='left'>
							<input type='button' value='GO' class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='JavaScript:Ca_Go();'>
						</td>
						<td align='left'>
							<a href='A_New_Case.asp' style='text-decoration:none'><font size='1' face='trebuchet MS'>[New]</font></a>
						</td>
					</tr>
				</table> 
			<% ElseIf Request("choice") = "Consumer" Then%>
				<center>
					<table border='0'>
						<tr>
							<td colspan='3' align='center'><font size='2' face='trebuchet MS'><u><b>CONSUMER LIST</b></u>&nbsp;&nbsp;</font>
								<font size='2' face='trebuchet MS'>&nbsp;&nbsp;All:</font><input type='radio' <%=chk1%> name='chk' value='1' onclick='JavaScript:ConSort();'>
							
								<font size='2' face='trebuchet MS'>&nbsp;&nbsp;Active:</font><input type='radio' <%=chk2%> name='chk' value='2' onclick='JavaScript:ConSort();'>
						
								<font size='2' face='trebuchet MS'>&nbsp;&nbsp;Inactive:</font><input type='radio' <%=chk3%> name='chk' value='3' onclick='JavaScript:ConSort();'>
								<input type='hidden' name='type' value='con'>
							</td>
						</tr>
						<tr><td>&nbsp;</td></tr>
						<tr>
							<td align='center'>
								<select name='consumer' style='font-size: 10pt;'>
									<%=Session("strConsumer")%>
								</select>
							</td>
							<td align='left'>
								<input type='button' value='GO' class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='JavaScript:C_Go();'>
							</td>
							<td align='left'>
								<a href='A_New_Consumer.asp' style='text-decoration:none'><font size='1' face='trebuchet MS'>[New]</font></a>
							</td>
						</tr>
					</table> 
			<% ElseIf Request("choice") = "Worker" Then%>
				<center>
					<table>
						<tr>
							<td colspan='3' align='center'><font size='2' face='trebuchet MS'><u><b>PCSP WORKER LIST</b></u></font>
								<font size='2' face='trebuchet MS'>&nbsp;&nbsp;All:</font><input type='radio' <%=chk1%> name='chk' value='1' onclick='JavaScript:ConSort();'>
							
								<font size='2' face='trebuchet MS'>&nbsp;&nbsp;Active:</font><input type='radio' <%=chk2%> name='chk' value='2' onclick='JavaScript:ConSort();'>
						
								<font size='2' face='trebuchet MS'>&nbsp;&nbsp;Inactive:</font><input type='radio' <%=chk3%> name='chk' value='3' onclick='JavaScript:ConSort();'>
								<input type='hidden' name='type' value='wor'>
								</td>
						</tr>
						<tr><td>&nbsp;</td></tr>
						<tr>
							<td align='center'>
								<select name='worker' style='font-size: 10pt;'>
									<%=Session("strWorker")%>
								</select>
							</td>
							<td align='left'>
								<input type='button' value='GO' class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='JavaScript:W_Go();'>
							</td>
							<td align='left'>
								<a href='A_New_Worker.asp' style='text-decoration:none'><font size='1' face='trebuchet MS'>[New]</font></a>
							</td>
						</tr>
					</table> 
			<% Else %>
				<center>
					<br><BR>
				<table border='0'>
					<tr>
						<td align='center'><br><br>This is the <b>upgraded version/system for LSS In-Home Care</b>.<br><br> If you wish to view the old version/system,
							 click <a href="http://webapp1.lssnorth.org/archive.lssdb/">HERE</a>.<br><br>
						<font size="2">
							 *you will have to login again
							<br>*changes made in the old system/version will NOT affect the upgraded version/system, and vice-versa.</font>
						</td>
					</tr>
				</table>
				</center>
			<% End If%>
		</form>
		<!-- #include file="_boxDown.asp" -->
	</body>
</html>
<%Session("MSG") = "" %>
