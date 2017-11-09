<%Language=VBScript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_security.asp" -->
<%
	If Request("MNum") <> "" Then
		Set tblFiles = Server.CreateObject("ADODB.Recordset")
		sqlFiles = "SELECT * FROM C_Files_t WHERE Medicaid_Number = '" & Request("MNum") & "' "
		tblFiles.Open sqlFiles, g_strCONN, 1, 3
			If Not tblFiles.EOF Then
				MCNum = tblFiles("Medicaid_Number")
				If tblFiles("DEAS_Service_Plan") = True Then DEAS = "checked"
				If tblFiles("LSS_Care_Plan") = True Then LSS = "checked"
				If tblFiles("Vehicle_Release") = True Then VR = "checked"
				If tblFiles("Privacy_Statement") = True Then PS = "checked"
				If tblFiles("A_Representative_Form") = True Then ARF = "checked"
				If tblFiles("Roles_Respon_Outline") = True Then RRO = "checked"
				If tblFiles("C_Site_Safety_Check") = True Then CSSC = "checked"
				If tblFiles("Training_Requirements") = True Then TR = "checked"
				If tblFiles("Photo") = True Then Photo = "checked"
				If tblFiles("AckStmt") = True Then ASt = "checked"
				If tblFiles("ARepForm") = True Then ARep = "checked"
				DFax = tblFiles("DteFax")
			Else
				tblFiles.AddNew
				tblFiles("Medicaid_Number") = Request("MNum")
				tblFiles.Update
			End If
		tblFiles.Close
		Set tblFiles = Nothing
	Else
		Session("MSG") = "Sorry. To maintain integrity of database, please Sign-in again."
		Response.Redirect "default.asp"
	End If
%>
<html>
	<head>
		<title>LSS - In-Home Care - Consumer Details - Files</title>
		<script language='JavaScript'>
			function CDF_Edit()
			{
				document.frmConDetHFil.action = "A_C_Action.asp?page=4";
				document.frmConDetHFil.submit();
			}
			function onhold()
			{
				if (document.frmConDetHFil.chkARepForm.checked == true)
					{document.frmConDetHFil.DFax.disabled = false}
				else
					{document.frmConDetHFil.DFax.disabled = true};
			}
			function maskMe(str,textbox,loc,delim)
			{
				var locs = loc.split(',');
				for (var i = 0; i <= locs.length; i++)
				{
					for (var k = 0; k <= str.length; k++)
					{
						 if (k == locs[i])
						 {
							if (str.substring(k, k+1) != delim)
						 	{
						 		str = str.substring(0,k) + delim + str.substring(k,str.length);
			     			}
						}
					}
			 	}
				textbox.value = str
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
	<body bgcolor='white' onLoad='JavaScript: onhold();' LEFTMARGIN='0' TOPMARGIN='0'>
		<form method='post' name='frmConDetHFil'>
			<!-- #include file="_boxup.asp" -->
			<!-- #include file="_NavHeader.asp" -->
			<br>
			<center>
				<table border='0'>
					<tr>
						<td colspan='4' align='center' width='500px'>
							<font size='2' face='trebuchet MS'><b><u>Consumer Details - Files</u></b></font>
							<a href='A_Consumer.asp?MNum=<%=Request("MNum")%>' style='text-decoration:none'><font size='1' face='trebuchet MS'>[Details]</font></a>
							<a href='A_C_Status.asp?MNum=<%=Request("MNum")%>' style='text-decoration:none'><font size='1' face='trebuchet MS'>[Status]</font></a>
							<a href='A_C_Health.asp?MNum=<%=Request("MNum")%>' style='text-decoration:none'><font size='1' face='trebuchet MS'>[Health]</font></a>
							<font size='2' face='trebuchet MS'>[Files]</font>
							<a href='Log.asp?MNum=<%=Request("MNum")%>' style='text-decoration:none'><font size='1' face='trebuchet MS'>[Log]</font></a>
							<a href='cimport.asp?MNum=<%=Request("MNum")%>' style='text-decoration:none'><font size='1' face='trebuchet MS'>[Uploads]</font></a>
						</td>
					</tr>
					<tr><td colspan='2' align='center'><font color='red' face='trebuchet MS' size='1'>&nbsp;<%=Session("MSG")%>&nbsp;</font></td></tr>
					<tr>
						<td>
							<font size='1' face='trebuchet MS'><u>Name:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 120px;' readonly name='Cname' value="<%=Session("Cname")%>"></font>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td align='left'>
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							<input type='checkbox' name='chkDEAS' <%=DEAS%>>
							<font size='1' face='trebuchet MS'><u>DEAS Service Plan:</u>&nbsp;
						</td>
						<td align='left'>
							<input type='checkbox' name='chkLSS' <%=LSS%>>
							<font size='1' face='trebuchet MS'><u>LSS Care Plan:</u>&nbsp;
						</td>
					</tr>
					<tr>
						<td align='left'>
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							<input type='checkbox' name='chkVR' <%=VR%>>
							<font size='1' face='trebuchet MS'><u>Vehicle Release:</u>&nbsp;
						</td>
						<td align='left'>
							<input type='checkbox' name='chkPS' <%=PS%>>
							<font size='1' face='trebuchet MS'><u>Privacy Statement:</u>&nbsp;
						</td>
					</tr>
					<tr>
						<td align='left'>
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							<input type='checkbox' name='chkARF' <%=ARF%>>
							<font size='1' face='trebuchet MS'><u>Authorized Representative Form:</u>&nbsp;
						</td>
						<td align='left'>
							<input type='checkbox' name='chkRRO' <%=RRO%>>
							<font size='1' face='trebuchet MS'><u>Roles & Responsibilities Outline:</u>&nbsp;
						</td>
					</tr>
					<tr>
						<td align='left'>
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							<input type='checkbox' name='chkCSSC' <%=CSSC%>>
							<font size='1' face='trebuchet MS'><u>Consumer Site & Safety Checklist:</u>&nbsp;
						</td>
						<td align='left'>
							<input type='checkbox' name='chkTR' <%=TR%>>
							<font size='1' face='trebuchet MS'><u>Training Requirments:</u>&nbsp;
						</td>
					</tr>
					<tr>
						<td align='left'>
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							<input type='checkbox' name='chkPhoto' <%=Photo%>>
							<font size='1' face='trebuchet MS'><u>Photo Release</u>&nbsp;
						</td>
						<td align='left'>
							<input type='checkbox' name='chkAS' <%=ASt%>>
							<font size='1' face='trebuchet MS'><u>Acknowledgement Statement:</u>&nbsp;
						</td>
					</tr>
					<tr>
						<td align='left'>
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							<input type='checkbox' name='chkARepForm' <%=ARep%> onclick='JavaScript: onhold();'>
							<font size='1' face='trebuchet MS'><u>Authorized Rep.Form/Care Plan:</u>
						</td>
					</tr>
					<tr>
						<td colspan='1'>
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							<font size='1' face='trebuchet MS'><u>Faxed to Case Manager:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 80px;' maxlength='10' name='DFax' value="<%=DFax%>" onKeyUp="javascript:return maskMe(this.value,this,'2,5','/');" onBlur="javascript:return maskMe(this.value,this,'2,5','/');"></font>
							<font size='1' face='trebuchet MS'>(mm/dd/yyyy)</font>
						</td>
						
					</tr>
					<tr><td>&nbsp;</td></tr>
				</table>
				<table border='0'>
					<tr>
						<td align='center' colspan='2'>
							<input type='button' value='Save' style='width: 110px;' class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'"  onclick='JavaScript:CDF_Edit();'>
							<input type='hidden' name='Mnum' value="<%=Request("MNum")%>">
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
				</table>
			</center>	
			<!-- #include file="_boxdown.asp" -->
		</form>
	</body>
</html>
<%Session("MSG") = "" %>