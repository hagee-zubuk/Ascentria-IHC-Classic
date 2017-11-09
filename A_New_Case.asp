<%Language=VBScript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<%
	If Session("Cadet") <> "" Then
		tmpCadet = Split(Z_DoDecrypt(Session("Cadet")), "|")
		LN = tmpCadet(0)
		FN = tmpCadet(1)
		Adr = tmpCadet(2)
		C = tmpCadet(3)
		S = tmpCadet(4)
		Z = tmpCadet(5)
		A = tmpCadet(6)
		O = tmpCadet(7)
		Ce = tmpCadet(8)
		F = tmpCadet(9)
		Ext = tmpCadet(10)
		Em = tmpCadet(11)
	End If
	'''''''''''''''CMC
		Set rsPM = Server.CreateObject("ADODB.RecordSet")
		sqlPM = "SELECT * FROM CaseMngmt_T ORDER BY cmcname"
		rsPM.Open sqlPM, g_strCONN, 1, 3
		Do Until rsPM.EOF
			cmcname = rsPM("cmcname")
			Selcmc = ""
			If rsPM("cmcid") = cmc Then Selcmc = "SELECTED"
			strCMC = strCMC & "<option " & Selcmc & " value='" & rsPM("CMCid") & "' >" & cmcname & "</option>" 
			rsPM.MoveNext
		Loop
		rsPM.Close
		Set rsPM = Nothing
%>
<html>
	<head>
		<title>LSS - In-Home Care - Case Manager Details - NEW</title>
		<script language='JavaScript'>
			function CaDN_Edit()
			{
				document.frmCaseDet.action = "A_Ca_Action.asp?new=1";
				document.frmCaseDet.submit();
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
		<form method='post' name='frmCaseDet' >
			<!-- #include file="_boxup.asp" -->
			<!-- #include file="_NavHeader.asp" -->
			<br>
			<center>
				<table border='0'>
					<tr>
						<td colspan='4' align='center' width='500px'>
							<font size='2' face='trebuchet MS'><b><u>Case Manager Details</u></b></font>
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
					<!--<tr>
						<td>
							<font size='1' face='trebuchet MS'><u>Address:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 350px;' maxlength='49' name='Addr' value="<%=Adr%>"></font>
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
					<tr><td>&nbsp;</td></tr>//-->
					<tr>
						<td>
							<font size='1' face='trebuchet MS'><u>Case Management Company:</u></font>
							<select style='font-size: 10px; height: 20px; width: 150px;' name='selCMC'>
								<option value='0'></option>
								<%=strCMC%>
							</select>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
				<!--<tr>
						<td><font size='1' face='trebuchet MS'><u>Agency:</u>&nbsp;<input style='font-size: 10px; height: 20px; width: 120px;' maxlength='40' type='text' name='Agency' value="<%=Agency%>"></font></td>	
					</tr>
					<tr><td>&nbsp;</td></tr>//-->
					<tr>
						<td>
							<font size='1' face='trebuchet MS'><u>Office No:</u>&nbsp;<input style='font-size: 10px; height: 20px; width: 80px;' maxlength='12' type='text' name='PhoneNo' value="<%=P%>"></font>
							&nbsp;
							<font size='1' face='trebuchet MS'><u>Ext:</u>&nbsp;<input style='font-size: 10px; height: 20px; width: 80px;' type='text' maxlength='12' name='Ext' value="<%=Ext%>"></font>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td><font size='1' face='trebuchet MS'><u>Cell No:</u>&nbsp;<input style='font-size: 10px; height: 20px; width: 80px;' type='text' maxlength='12' name='CelNo' value="<%=Ce%>" ></font></td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td><font size='1' face='trebuchet MS'><u>Fax No:</u>&nbsp;<input style='font-size: 10px; height: 20px; width: 80px;' type='text' maxlength='12' name='FaxNo' value="<%=F%>"></font></td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td><font size='1' face='trebuchet MS'><u>eMail:</u>&nbsp;<input style='font-size: 10px; height: 20px; width: 120px;' type='text' maxlength='50' name='email' value="<%=Em%>"></font></td>
					</tr>
					<tr><td>&nbsp;</td></tr>//-->
				</table>	
				<table border='0'>
					<tr>
						<td align='center' colspan='2'>
							<input type='button' value='Save' style='width: 110px;' class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'"  onclick='JavaScript:CaDN_Edit();'>
							
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
	Session("CaDet") = ""
%>