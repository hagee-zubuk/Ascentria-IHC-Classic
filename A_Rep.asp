<%language=vbscript%>
<!-- #include file="_Files.asp" -->

<!-- #include file="_Utils.asp" -->
<%
	'If Request("RID") <> "" Then 
	'	If Request("Rep") = "" Then Session("RID") = Request("RID")
	'Else
		
	'End If
	If Request("Rep") <> "" Then
		tmpMCNum = split(Request("Rep"), " - ") 
		ID = tmpMCNum(0)
		Session("RID") = ID
	Else
		If Request("RID") <> "" Then Session("RID") = Request("RID")
	End If
	Set tblRep = Server.CreateObject("ADODB.Recordset")
	sqlRep = "SELECT * FROM Representative_t WHERE [Index] = " & Session("RID") 
	tblRep.Open sqlRep, g_strCONN, 1, 3
		If Not tblRep.EOF Then
			Index = tblRep("Index")
			fname = tblRep("Fname")
			lname = tblRep("Lname")
			Session("Rname") = lname & ", " & fname
			adr = tblRep("Address")
			cty = tblRep("City")
			ste = tblRep("State")
			zcode = tblRep("Zip")
			Pno = tblRep("PhoneNo")
			wPno = tblRep("wPhoneNo")
			cPno = tblRep("cPhoneNo")
			email = tblRep("email")
		Else
			tblRep.AddNew
			tblRep("Index") = Session("RID")
			tblRep.Update
		End If
	tblRep.Close
	Set tblRep = Nothing	
%>
<html>
	<head>
		<title>LSS - In-Home Care - Representative Details</title>
		<script language='JavaScript'>
			function RD_Del()
			{
				var ans = window.confirm("Click OK to continue deletion of Representative. Click Cancel to stop.");
				if (ans){
				document.frmRepDet.action = "A_Del.asp?act=7";
				document.frmRepDet.submit();
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
	<body bgcolor='white' LEFTMARGIN='0' TOPMARGIN='0'>
		<form method='post' name='frmRepDet' action='A_R_Action.asp'>
			<!-- #include file="_boxup.asp" -->
			<!-- #include file="_NavHeader.asp" -->
			<br>
			<center>
				<table border='0'>
					<tr>
						<td colspan='4' align='center' width='500px'>
							<font size='2' face='trebuchet MS'><b><u>Representative Details</u></b></font>
							<font size='2' face='trebuchet MS'>[Details]</font>
							<a href="RepCon.asp?ID=<%=Index%>" style='text-decoration: none;'>
								<font color='blue' face='trebuchet MS' size='1'>[List]</font>
							</a>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr><td>&nbsp;</td></tr>
					
					<tr>
						<td>
							<font size='1' face='trebuchet MS'><u>Name:</u>&nbsp;<input style='font-size: 10px; height: 20px; width: 90px;' type='text' maxlength='20' name='lname' value="<%=Lname%>">,&nbsp; 
							<input maxlength='20' style='font-size: 10px; height: 20px; width: 90px;' type='text' name='fname' value="<%=Fname%>">
							<font size='1' face='trebuchet MS'>(Last Name, First Name)</font>
							<input type='hidden' name='MCNum' value='<%=Index%>'>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td>
							<font size='1' face='trebuchet MS'><u>Address:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 350px;' maxlength='49' name='Addr' value="<%=Adr%>"></font>
						</td>
						<tr>
						<td colspan='2'>
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							<font size='1' face='trebuchet MS'><u>City:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 100px;' name='cty' maxlength='49' value="<%=cty%>">
							<font size='1' face='trebuchet MS'><u>State:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 50px;' name='ste' maxlength='2' value="<%=ste%>">
							<font size='1' face='trebuchet MS'><u>Zip Code:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 100px;' name='zcode' maxlength='10' value="<%=zcode%>">
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
					<td><font size='1' face='trebuchet MS'><u>Home Phone:</u>&nbsp;<input style='font-size: 10px; height: 20px; width: 80px;' maxlength='12' type='text' name='PhoneNo' value="<%=pno%>"></font></td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td><font size='1' face='trebuchet MS'><u>Work Phone:</u>&nbsp;<input style='font-size: 10px; height: 20px; width: 80px;' maxlength='12' type='text' name='wPhoneNo' value="<%=wpno%>"></font></td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td><font size='1' face='trebuchet MS'><u>Mobile Phone:</u>&nbsp;<input style='font-size: 10px; height: 20px; width: 80px;' maxlength='12' type='text' name='cPhoneNo' value="<%=cpno%>"></font></td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td><font size='1' face='trebuchet MS'><u>eMail:</u>&nbsp;<input style='font-size: 10px; height: 20px; width: 125px;' maxlength='50' type='text' name='Remail' value="<%=email%>"></font></td>
					</tr>
					<tr><td>&nbsp;</td></tr>
				</table>	
				<table border='0'>
					<tr>
						<td align='center' colspan='2'>
							<input type='button' value='Save' style='width: 110px;' class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='document.frmRepDet.submit();'>
							<input type='button' value='Delete' style='width: 110px;' class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'"  onclick='JavaScript:RD_Del();'>
							
						</td>
					</tr>
				</table>
			</center>	
			<!-- #include file="_boxdown.asp" -->
		</form>
	</body>
</html>
<% Session("Rdet") = ""%>