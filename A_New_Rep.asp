<%Language=VBScript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<%
		LN = ""
		FN = ""
		Adr = ""
		C = ""
		S = ""
		Z = ""
		P = "" 
		wP = ""
		cP = ""
		em = ""
		'response.write "FALSE"
	If Session("Rdet") <> "" Then
		tmpRdet = Split(Z_DoDecrypt(Session("Rdet")), "|")
		LN = tmpRdet(0)
		FN = tmpRdet(1)
		Adr = tmpRdet(2)
		C = tmpRdet(3)
		S = tmpRdet(4)
		Z = tmpRdet(5)
		P = tmpRdet(6)
		wP = tmpRdet(7)
		cP = tmpRdet(8)
		em = tmpRdet(9)
		'response.write "TRUE"
	End If
%>
<html>
	<head>
		<title>LSS - In-Home Care - Representative Details - NEW</title>
		<script language='JavaScript'>
			function RDN_Edit()
			{
				document.frmRepDet.action = "A_R_Action.asp?new=1";
				document.frmRepDet.submit();
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
	<body bgcolor='white' LEFTMARGIN='0' TOPMARGIN='0'>
		<form method='post' name='frmRepDet' >
			<!-- #include file="_boxup.asp" -->
			<!-- #include file="_NavHeader.asp" -->
			<br>
			<center>
				<table border='0'>
					<tr>
						<td colspan='4' align='center' width='500px'>
							<font size='2' face='trebuchet MS'><b><u>Representative Details</u></b></font>
							<font size='1' face='trebuchet MS'>[Details]</font>
							
								<font size='1' face='trebuchet MS'>[List]</font>
						</td>
					</tr>
					
				<tr><td colspan='2' align='center'><font color='red' face='trebuchet MS' size='1'>&nbsp;<%=Session("MSG")%>&nbsp;</font></td></tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td>
							<font size='1' face='trebuchet MS'><u>Name:</u>&nbsp;<input style='font-size: 10px; height: 20px; width: 90px;' type='text' maxlength='20' name='lname' value="<%=LN%>">,&nbsp; 
							<input style='font-size: 10px; height: 20px; width: 90px;' type='text' name='fname' maxlength='20' value="<%=FN%>">
							<font size='1' face='trebuchet MS'>(Last Name, First Name)</font>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td>
							<font size='1' face='trebuchet MS'><u>Address:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 350px;' maxlength='49' name='Addr' value='<%=Adr%>'></font>
						</td>
						<tr>
						<td colspan='2'>
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							<font size='1' face='trebuchet MS'><u>City:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 100px;' name='cty' maxlength='49' value="<%=C%>">
							<font size='1' face='trebuchet MS'><u>State:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 50px;' name='ste' maxlength='2' value="<%=S%>">
							<font size='1' face='trebuchet MS'><u>Zip Code:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 100px;' name='zcode' maxlength='10' value="<%=Z%>">
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td><font size='1' face='trebuchet MS'><u>Home Phone:</u>&nbsp;
							<input style='font-size: 10px; height: 20px; width: 80px;' maxlength='12' type='text' value='<%=P%>' name='PhoneNo'></font></td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td><font size='1' face='trebuchet MS'><u>Work Phone:</u>&nbsp;<input style='font-size: 10px; height: 20px; width: 80px;' maxlength='12' type='text' name='wPhoneNo' value="<%=wP%>"></font></td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td><font size='1' face='trebuchet MS'><u>Mobile Phone:</u>&nbsp;<input style='font-size: 10px; height: 20px; width: 80px;' maxlength='12' type='text' name='cPhoneNo' value="<%=cP%>"></font></td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td><font size='1' face='trebuchet MS'><u>eMail:</u>&nbsp;<input style='font-size: 10px; height: 20px; width: 125px;' maxlength='50' type='text' name='Remail' value="<%=em%>"></font></td>
					</tr>
					<tr><td>&nbsp;</td></tr>
				</table>	
				<table border='0'>
					<tr>
						<td align='center' colspan='2'>
							<input type='button' value='Save' style='width: 110px;' class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'"  onclick='JavaScript:RDN_Edit();'>
							
						</td>
					</tr>
				</table>
			</center>	
			<!-- #include file="_boxdown.asp" -->
		</form>
	</body>
</html>
<%
Session("MSG") = ""
Session("Rdet") = ""
 %>