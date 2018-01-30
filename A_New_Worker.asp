<%language=vbscript%>
<!-- #include file="_Files.asp" -->

<!-- #include file="_Utils.asp" -->
<!-- #include file="_security.asp" -->
<%
	Session("WID") = ""
	response.write "<!--tmpFiles:" & Z_DoDecrypt(Session("WFiles")) & "-->"
	myFin = ""
	'If UCase(Session("lngType")) = "0" Then myFin = "READONLY"
	If Session("WFiles") <> "" Then 
		tmpWFiles = Split(Z_DoDecrypt(Session("WFiles")),"|")

		SSN = tmpWFiles(0)
		LN = tmpWFiles(1)
		FN = tmpWFiles(2)
		AdR = tmpWFiles(3)
		G = tmpWFiles(4)
		DOB = tmpWFiles(5)
		PN = tmpWFiles(6)
		CN = tmpWFiles(7)
		DH = tmpWFiles(8)
		S = tmpWFiles(9)
		SC = tmpWFiles(10)
		CDv = tmpWFiles(11)
		CF = tmpWFiles(12)
		LsN = tmpWFiles(13)
		LED = tmpWFiles(14)
		CI = tmpWFiles(15)
		ID = tmpWFiles(16)
		CT = tmpWFiles(17)
		T = tmpWFiles(18)
		cty = tmpWFiles(19) 
		ste = tmpWFiles(20)
		zcode = tmpWFiles(21)
		TD = tmpWFiles(22)
		mAdR = tmpWFiles(23)
		mcty = tmpWFiles(24) 
		mste = tmpWFiles(25)
		mzcode = tmpWFiles(26)
		Sal = tmpWFiles(27)
		chkMail = ""
		If tmpWFiles(28) <> "" Then chkMail = "checked"
		Ema = tmpWFiles(29)
		tmpMan = tmpWFiles(30)
		FNum = tmpWFiles(31)
		pp = ""
		If tmpWFiles(32) <> "" Then pp = "checked"
	End If
	Set rsPM = Server.CreateObject("ADODB.RecordSet")
		sqlPM = "SELECT * FROM Proj_Man_T ORDER BY Lname, Fname"
		rsPM.Open sqlPM, g_strCONN, 1, 3
		Do Until rsPM.EOF
			PMname = rsPM("Lname") & ", " & rsPM("Fname")
			SelPM = ""
			If rsPM("ID") = PM1 Then SelPM = "SELECTED"
			SelPM2 = ""
			If rsPM("ID") = PM2 Then SelPM2 = "SELECTED"
			strPM = StrPM & "<option " & SelPM & " value='" & rsPM("ID") & "' >" & PMname & "</option>" 
			strPM2 = StrPM2 & "<option " & SelPM2 & " value='" & rsPM("ID") & "' >" & PMname & "</option>" 
			rsPM.MoveNext
		Loop
		rsPM.Close
		Set rsPM = Nothing
		''''county
		Set rsPM = Server.CreateObject("ADODB.RecordSet")
		sqlPM = "SELECT * FROM county_T ORDER BY county"
		rsPM.Open sqlPM, g_strCONN, 1, 3
		Do Until rsPM.EOF
			
			strcount = strcount & "<option value='" & rsPM("uid") & "' >" & rsPM("county") & "</option>" 
			rsPM.MoveNext
		Loop
		rsPM.Close
		Set rsPM = Nothing
%>
<html>
	<head>
		<title>LSS - In-Home Care - PCSP WORKER Details - NEW</title>
		<script language='JavaScript'>
			function WD_Edit()
			{
				if (document.frmWorDet.chkDrive.checked == true)
				{
					if (document.frmWorDet.chkonFile.checked == false)
					{
						alert ("ERROR: 'Copy of License on File' should be checked.")
						document.frmWorDet.chkonFile.focus();
						return;
					}
					if (document.frmWorDet.LisNo.value == "")
					{
						alert ("ERROR: 'License #' is blank.")
						document.frmWorDet.LisNo.focus();
						return;
					}
					if (document.frmWorDet.LisExpDte.value == "")
					{
						alert ("ERROR: 'License Expiration Date' is blank.")
						document.frmWorDet.LisExpDte.focus();
						return;
					}
					if (document.frmWorDet.chkIns.checked == false)
					{
						alert ("ERROR: 'Copy of Auto Insurance' should be checked.")
						document.frmWorDet.chkIns.focus();
						return;
					}
					if (document.frmWorDet.Insdate.value == "")
					{
						alert ("ERROR: 'Insurance Expiration Date:' is blank.")
						document.frmWorDet.Insdate.focus();
						return;
					}
				}
				document.frmWorDet.action = "A_W_Action.asp?page=1&new=1";
				document.frmWorDet.submit();
			}
			function Sep()
			{
				if (document.frmWorDet.Stat.value == "Inactive")
					{document.frmWorDet.SepCode.disabled = false}
				else
					{document.frmWorDet.SepCode.disabled = true};
			}
			function Driver()
			{
				if (document.frmWorDet.chkDrive.checked == true)
					{document.frmWorDet.chkonFile.disabled = false
					document.frmWorDet.LisNo.disabled = false
					document.frmWorDet.chkIns.disabled = false
					document.frmWorDet.LisExpDte.disabled = false
					document.frmWorDet.Insdate.disabled = false}
				else
					{document.frmWorDet.chkonFile.disabled = true
					document.frmWorDet.LisNo.disabled = true
					document.frmWorDet.chkIns.disabled = true
					document.frmWorDet.LisExpDte.disabled = true
					document.frmWorDet.Insdate.disabled = true};
			}
			function InterTown()
			{
				if (document.frmWorDet.chkTown.checked == true)
					{document.frmWorDet.Towns.disabled = false}
				else
					{document.frmWorDet.Towns.disabled = true};
					
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
	<body bgcolor='white' LEFTMARGIN='0' TOPMARGIN='0' onload='JavaScript:Sep(); Driver(); InterTown();'>
		<form method='post' name='frmWorDet'>
			<!-- #include file="_boxup.asp" -->
			<!-- #include file="_NavHeader.asp" -->
			<br>
			<center>
				<table border='0'>
					<tr>
						<td colspan='4' align='center' width='500px'>
							<font size='2' face='trebuchet MS'><b><u>PCSP Worker Details</u></b></font>
							<font size='1' face='trebuchet MS'>[Details]</font>
							<font size='1' face='trebuchet MS'>[Files]</font>
							<font size='1' face='trebuchet MS'>[Skills]</font>
							<font size='1' face='trebuchet MS'>[Log]</font>
							<font size='1' face='trebuchet MS'>[Violations]</font>
							<font size='1' face='trebuchet MS'>[Uploads]</font>
						</td>
					</tr>
					<tr><td colspan='2' align='center'><font color='red' face='trebuchet MS' size='1'>&nbsp;<%=Session("MSG")%>&nbsp;</font></td></tr>
					<tr>
						<td>
							<font size='1' face='trebuchet MS'><u>SSN:</u>&nbsp;
							<input style='font-size: 10px; height: 20px; width: 80px;' type='text' maxlength='11' name='SSN' onKeyUp="javascript:return maskMe(this.value,this,'3,6','-');" onBlur="javascript:return maskMe(this.value,this,'3,6','-');">
							<font size='1' face='trebuchet MS'>*Required
							</td>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td>
							<font size='1' face='trebuchet MS'><u>Name:</u>&nbsp;<input style='font-size: 10px; height: 20px; width: 90px;' maxlength='20' type='text' name='lname' value='<%=LN%>'>,&nbsp; 
							<input style='font-size: 10px; height: 20px; width: 90px;' type='text' maxlength='20' name='fname' value='<%=FN%>'>
							<font size='1' face='trebuchet MS'>(Last Name, First Name)</font>
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							<input type='checkbox' name='chkflt' <%=flt%> value='1'><u>Float Worker</u>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td>
							<font size='1' face='trebuchet MS'><u>Residence Address:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 250px;' name='Addr' maxlength='49' value="<%=ADr%>"></font>
						</td>
					</tr>
					<tr>
						<td colspan='2'>
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							<font size='1' face='trebuchet MS'><u>City:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 100px;' name='cty' maxlength='49' value="<%=cty%>">
							<font size='1' face='trebuchet MS'><u>State:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 50px;' name='ste' maxlength='2' value="<%=ste%>">
							<font size='1' face='trebuchet MS'><u>Zip Code:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 100px;' name='zcode' maxlength='10' value="<%=zcode%>" onKeyUp="javascript:return maskMe(this.value,this,'5','-');" onBlur="javascript:return maskMe(this.value,this,'5','-');">
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td>
							<font size='1' face='trebuchet MS'><u>Mailing Address:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 250px;' name='mAddr' maxlength='49' value="<%=mADr%>"></font>
							<input type='checkbox' name='chkMail' <%=chkMail%>><font size='1' face='trebuchet MS'><u>Same as Residence:</u></font>
						</td>
					</tr>
					<tr>
						<td colspan='2'>
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							<font size='1' face='trebuchet MS'><u>City:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 100px;' name='mcty' maxlength='49' value="<%=mcty%>">
							<font size='1' face='trebuchet MS'><u>State:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 50px;' name='mste' maxlength='2' value="<%=mste%>">
							<font size='1' face='trebuchet MS'><u>Zip Code:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 100px;' name='mzcode' maxlength='10' value="<%=mzcode%>" onKeyUp="javascript:return maskMe(this.value,this,'5','-');" onBlur="javascript:return maskMe(this.value,this,'5','-');">
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td>
							<font size='1' face='trebuchet MS'><u>County:</u>&nbsp;
								<select style='font-size: 10px; height: 20px; width: 100px;' name='selcount'></font>
								<%=strcount%>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
						<tr>
						<td>
							<font size='1' face='trebuchet MS'><u>Gender:</u>&nbsp;
							<select style='font-size: 10px; height: 20px; width: 80px;' name='Gen'></font>
								<option>Male</option>
								<option>Female</option>
							</select>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td colspan='2'>
							<font size='1' face='trebuchet MS'><u>DOB:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 80px;' maxlength='10' name='DOB' value='<%=DOB%>' onKeyUp="javascript:return maskMe(this.value,this,'2,5','/');" onBlur="javascript:return maskMe(this.value,this,'2,5','/');"></font>
							<font size='1' face='trebuchet MS'>(mm/dd/yyyy)</font>
							<font size='1' face='trebuchet MS'><u>Phone No:</u>&nbsp;<input style='font-size: 10px; height: 20px; width: 80px;' type='text' maxlength='12' name='PhoneNo' value='<%=PN%>'></font>
							<font size='1' face='trebuchet MS'><u>Cell No:</u>&nbsp;<input style='font-size: 10px; height: 20px; width: 80px;' type='text' maxlength='12' name='CellNo' value='<%=CN%>'></font>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td colspan='12'>
							<font size='1' face='trebuchet MS'><u>eMail:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 120px;' maxlength='50' name='eMail' value="<%=Ema%>"></font>
							<font size='1' face='trebuchet MS'><u>Method of Communication:</u>&nbsp;
							<select style='font-size: 10px; height: 20px; width: 150px;' name='selpref'></font>
							<option value='0'>&nbsp;</option>
							<option value='1' <%=prefmail%>>Mail</option>
							<option value='2' <%=prefemail%>>Email</option>
							<option value='3' <%=prefFon%>>Phone</option>
							<option value='4' <%=preftxt%>>Text</option>
						</select>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td colspan='2'>
							<font size='1' face='trebuchet MS'><u>Date of Hire:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 80px;' maxlength='10' name='DateHired' value='<%=DH%>' onKeyUp="javascript:return maskMe(this.value,this,'2,5','/');" onBlur="javascript:return maskMe(this.value,this,'2,5','/');"></font>
							<font size='1' face='trebuchet MS'>(mm/dd/yyyy)</font>
							<font size='1' face='trebuchet MS'><u>Termination Date:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 80px;' name='TermD8' maxlength='10' value="<%=TD%>" onKeyUp="javascript:return maskMe(this.value,this,'2,5','/');" onBlur="javascript:return maskMe(this.value,this,'2,5','/');"></font>
							<font size='1' face='trebuchet MS'>(mm/dd/yyyy)</font>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td colspan='2'>
							<font size='1' face='trebuchet MS'><u>Employee Manual Received:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 80px;' name='DateManual' maxlength='10' value="<%=tmpMan%>" onKeyUp="javascript:return maskMe(this.value,this,'2,5','/');" onBlur="javascript:return maskMe(this.value,this,'2,5','/');"></font>
							<font size='1' face='trebuchet MS'>(mm/dd/yyyy)</font>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td>
							<font size='1' face='trebuchet MS'><u>Status:</u>&nbsp;
							<select style='font-size: 10px; height: 20px; width: 120px;' name='Stat' onchange='JavaScript:Sep();'></font>
								<option <%=A%> value='1'>Active</option>
								<option <%=I%> value='2'>Inactive</option>
								<option <%=P%> value='3'>Potential Applicant</option>
							</select>
							<font size='1' face='trebuchet MS'><u>Separation Code:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 80px;' maxlength='4' name='SepCode' value='<%=SC%>'></font>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td colspan='2'>
							<font size='1' face='trebuchet MS'><u>Salary:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 80px;' name='txtSal' maxlength='8' value="<%=Sal%>"></font>
							<font size='1' face='trebuchet MS'><u>File Number:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 80px;' name='txtFN' maxlength='6' value="<%=FNum%>" <%=myFin%>></font>
							<input type='checkbox' name='chkPP' <%=pp%> ><font size='1' face='trebuchet MS'><u>Private Pay Eligible Worker:</u></font>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td colspan='12'>
							<font size='1' face='trebuchet MS'><u>Comments:</u>&nbsp;
							<textarea rows='2' name='Pcomments' cols="20" ><%=cmt%></textarea>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td>
							<font size='1' face='trebuchet MS'><u>Misdemeanor on File:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 250px;' name='misd' maxlength='250' value="<%=misd%>"></font>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td>
							<font size='1' face='trebuchet MS'><u>Written Warning:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 250px;' name='warn' maxlength='250' value="<%=warn%>"></font>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td colspan='2'>
							<font size='1' face='trebuchet MS'><u>RIHCC 1:</u>&nbsp;
							<select style='font-size: 10px; height: 20px; width: 150px;' name='selpm1'></font>
								<option value='0'>&nbsp;</option>
								<%=strPM%>
							</select>
							<font size='1' face='trebuchet MS'><u>RIHCC 2:</u>&nbsp;
							<select style='font-size: 10px; height: 20px; width: 150px;' name='selpm2'></font>
								<option value='0'>&nbsp;</option>
								<%=strPM2%>
							</select>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td align='left'>
							<% If CDv <> "" Then chk = "checked" %>
							<input type='checkbox' name='chkDrive' onclick='JavaScript:Driver();' <%=chk%>>
							<font size='1' face='trebuchet MS'><u>Driver:</u>&nbsp;
						</td>
					</tr>
					<tr>
						<td align='left'>
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							<% If CF <> "" Then chk2 = "checked" %>
							<input type='checkbox' name='chkonFile' <%=chk2%>>
							<font size='1' face='trebuchet MS'><u>Copy of License on File:</u>&nbsp;
						</td>
					</tr>
					<tr>
						<td align='left'>
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							<font size='1' face='trebuchet MS'><u>License #:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 100px;' maxlength='17' name='LisNo' value='<%=LsN%>'></font>
						</td>
					</tr>
					<tr>
						<td align='left'>
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							<font size='1' face='trebuchet MS'><u>License Expiration Date:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 80px;' maxlength='10' name='LisExpDte' value='<%=LED%>' onKeyUp="javascript:return maskMe(this.value,this,'2,5','/');" onBlur="javascript:return maskMe(this.value,this,'2,5','/');"></font>
							<font size='1' face='trebuchet MS'>(mm/dd/yyyy)</font>
						</td>
					</tr>
					<tr>
						<td align='left'>
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							<% If CI <> "" Then chk3 = "checked" %>
							<input type='checkbox' name='chkIns' <%=chk3%>>
							<font size='1' face='trebuchet MS'><u>Copy of Auto Insurance:</u>&nbsp;
						</td>
					</tr>
					<tr>
						<td align='left'>
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							<font size='1' face='trebuchet MS'><u>Insurance Expiration Date:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 80px;' maxlength='10' name='Insdate' value='<%=ID%>' onKeyUp="javascript:return maskMe(this.value,this,'2,5','/');" onBlur="javascript:return maskMe(this.value,this,'2,5','/');"></font>
							<font size='1' face='trebuchet MS'>(mm/dd/yyyy)</font>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td align='left'>
							<% If CT <> "" Then chk4 = "checked" %>
							<input type='checkbox' name='chkTown' onclick='JavaScript:InterTown();' <%=chk4%>>
							<font size='1' face='trebuchet MS'><u>Interested in Working with More Consumer:</u>&nbsp;
						</td>
					</tr>
					<tr>
						<td>
							<font size='1' face='trebuchet MS'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<u>Towns:</u></font>
							<a href='' style='text-decoration:none'><font size='1' face='trebuchet MS'>[Delete]</font></a>	
							<font size='1' face='trebuchet MS'>*enter one Town at a time.</font>
						</td>
					</tr>
					&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
					<%=strTowns%>
					<tr>
						<td>
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							<input type='text' maxlength='20' name='Towns' style='font-size: 10px; height: 20px; width: 80px;' value='<%=T%>'></font>
							<input type='hidden' name='ctr' value='<%=ctr%>'>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td align='left'>
							<input type='checkbox' name='essentials' <%=essentials%>>
							<font size='1' face='trebuchet MS'><u>Essentials Training Required:</u>&nbsp;
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
				</table>
				<table border='0'>
					<tr>
						<td align='center' colspan='2'>
							<input type='button' value='Save' style='width: 110px;' class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='JavaScript:WD_Edit();'>
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
	Session("WFiles") = ""
%>