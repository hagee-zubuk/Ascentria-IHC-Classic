<%Language=VBScript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_security.asp" -->
<%
	'response.write Session("CFiles")
	If Session("CFiles") <> "" Then 
		tmpCFiles = Split(Z_DoDecrypt(Session("CFiles")),"|")
		MN = tmpCFiles(0)
		SSN = tmpCFiles(1)
		LN = tmpCFiles(2)
		FN = tmpCFiles(3)
		ADr = tmpCFiles(4)
		PN = tmpCFiles(5)
		DOB = tmpCFiles(6)
		G = tmpCFiles(7)
		D = tmpCFiles(8)
		RD = tmpCFiles(9)
		SD = tmpCFiles(10)
		CDv = tmpCFiles(11)
		CR = tmpCFiles(12)
		MH = tmpCFiles(13)
		C = tmpCFiles(14)
		cty = tmpCFiles(15) 
		ste = tmpCFiles(16)
		zcode = tmpCFiles(17)
		AmS = tmpCFiles(18)
		AR = tmpCFiles(19)
		PM = tmpCFiles(20)
		TmD = tmpCFiles(21)
		ED = tmpCFiles(22)
		CareDte = tmpCFiles(23)
		EfD =  tmpCFiles(24)
		milecap = tmpCFiles(25)
		email = tmpCFiles(26)
		mAdR = tmpCFiles(27)
		mcty = tmpCFiles(28) 
		mste = tmpCFiles(29)
		mzcode = tmpCFiles(30)
		chkMail = ""
		If tmpCFiles(31) <> "" Then chkMail = "checked"
		code = tmpCFiles(32)
		mcode = ""
		pcode = ""
		ccode = ""
		acode = ""
		vcode = ""
		If code = "M" Then mcode = "selected"
		If code = "P" Then pcode = "selected"
		If code = "C" Then ccode = "selected"
		If code = "A" Then acode = "selected"
		If code = "V" Then vcode = "selected"
		contract = tmpCFiles(33)
		tmprate = tmpCFiles(34)
	End If
	Set rsPM = Server.CreateObject("ADODB.RecordSet")
	sqlPM = "SELECT * FROM Proj_Man_T ORDER BY Lname, Fname"
	rsPM.Open sqlPM, g_strCONN, 1, 3
	Do Until rsPM.EOF
		PMname = rsPM("Lname") & ", " & rsPM("Fname")
		SelPM = ""
		If rsPM("ID") = PM Then SelPM = "SELECTED"
		strPM = StrPM & "<option " & SelPM & " value='" & rsPM("ID") & "' >" & PMname & "</option>" 
		rsPM.MoveNext
	Loop
	rsPM.Close
	Set rsPM = Nothing
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
		'''''''''''''''MCC
		Set rsPM = Server.CreateObject("ADODB.RecordSet")
		sqlPM = "SELECT * FROM ManagedCare_T ORDER BY MCCname"
		rsPM.Open sqlPM, g_strCONN, 1, 3
		Do Until rsPM.EOF
			MCCname = rsPM("MCCname")
			Selmmc = ""
			If rsPM("MCCid") = cmc Then Selmmc = "SELECTED"
			strmcc = strmcc & "<option " & Selmmc & " value='" & rsPM("MCCid") & "' >" & MCCname & "</option>" 
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
		<title>LSS - In-Home Care - Consumer Details - NEW</title>
		<script language='JavaScript'>
			function CDN_Edit()
			{
				document.frmNewConDet.action = "A_C_Action.asp?page=1&new=1";
				document.frmNewConDet.submit();
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
			function myRate()	{
				document.frmNewConDet.chkVAHM.disabled = true;
				document.frmNewConDet.hrshm.disabled = true;
				document.frmNewConDet.chkVAHA.disabled = true;
				document.frmNewConDet.hrsha.disabled = true;
				document.frmNewConDet.chkVAHM.checked = false;
				document.frmNewConDet.hrshm.value = "";
				document.frmNewConDet.chkVAHA.checked = false;
				document.frmNewConDet.hrsha.value = "";
				if (document.frmNewConDet.selcode.value == "M")
					{document.frmNewConDet.txtrate.disabled = true}
				else if (document.frmNewConDet.selcode.value == "V") {
					document.frmNewConDet.chkVAHM.disabled = false;
					document.frmNewConDet.hrshm.disabled = false;
					document.frmNewConDet.chkVAHA.disabled = false;
					document.frmNewConDet.hrsha.disabled = false;
				}
				else
					{document.frmNewConDet.txtrate.disabled = false};
			}
			function hrschk(xxx) {
				if (xxx == 1) {
					if (document.frmNewConDet.chkVAHM.checked == false) {
						document.frmNewConDet.hrshm.value = "";
					}
				}
				else if(xxx == 2){
					if (document.frmNewConDet.chkVAHA.checked == false) {
						document.frmNewConDet.hrsha.value = "";
					}
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
	<body bgcolor='white' LEFTMARGIN='0' TOPMARGIN='0' onload='JavaScript:myRate();'>
		<form method='post' name='frmNewConDet' >
			<!-- #include file="_boxup.asp" -->
			<!-- #include file="_NavHeader.asp" -->
			<br>
			<center>
				<table border='0'>
					<tr>
						<td colspan='4' align='center' width='550px'>
							<font size='2' face='trebuchet MS'><b><u>Consumer Details</u></b></font>
							<font size='1' face='trebuchet MS'>[Details]</font>
							<font size='1' face='trebuchet MS'>[Status]</font>
							<font size='1' face='trebuchet MS'>[Health]</font>
							<font size='1' face='trebuchet MS'>[Files]</font>
							<font size='1' face='trebuchet MS'>[Log]</font>
						</td>
					</tr>
					<tr><td colspan='2' align='center'><font color='red' face='trebuchet MS' size='1'>&nbsp;<%=Session("MSG")%>&nbsp;</font></td></tr>
					<tr>
						<td>
							<font size='1' face='trebuchet MS'><u>Medicaid Number:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 90px;'  maxlength='14' name='MCnum' value="<%=MCNum%>"></font>
							<font size='1' face='trebuchet MS'>*Required
						&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
						&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
						&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							<font size='1' face='trebuchet MS'><u>SSN:</u>&nbsp;
							<input style='font-size: 10px; height: 20px; width: 90px;' type='text' maxlength='11' name='SSN' value='<%=SSN%>'></td>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td>
							<font size='1' face='trebuchet MS'><u>Name:</u>&nbsp;<input maxlength='20' style='font-size: 10px; height: 20px; width: 90px;' type='text' name='lname' value='<%=LN%>'>,&nbsp; 
							<input maxlength='20' style='font-size: 10px; height: 20px; width: 90px;' type='text' name='fname' value='<%=FN%>'>
							<font size='1' face='trebuchet MS'>(Last Name, First Name)</font>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td>
							<font size='1' face='trebuchet MS'><u>Address:</u>&nbsp;
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
							<input type='text' style='font-size: 10px; height: 20px; width: 100px;' name='zcode' maxlength='10' value="<%=zcode%>">
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
							<input type='text' style='font-size: 10px; height: 20px; width: 100px;' name='mzcode' maxlength='10' value="<%=mzcode%>">
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
							<font size='1' face='trebuchet MS'><u>Phone No:</u>&nbsp;<input style='font-size: 10px; height: 20px; width: 80px;' type='text' maxlength='12' name='PhoneNo' value="<%=PN%>"></font>
							&nbsp;&nbsp;&nbsp;&nbsp;
							<font size='1' face='trebuchet MS'><u>eMail:</u>&nbsp;<input style='font-size: 10px; height: 20px; width: 125px;' type='text' maxlength='50' name='email' value="<%=email%>"></font>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
						<tr>
						<td>
							<font size='1' face='trebuchet MS'><u>Secondary Phone No:</u>&nbsp;<input style='font-size: 10px; height: 20px; width: 80px;' type='text' maxlength='12' name='PhoneNo2' value="<%=FonNum2%>" onKeyUp="javascript:return maskMe(this.value,this,'3,7','-');" onBlur="javascript:return maskMe(this.value,this,'3,7','-');"></font>
							&nbsp;&nbsp;&nbsp;&nbsp;
							<font size='1' face='trebuchet MS'><u>Mobile Number:</u>&nbsp;<input style='font-size: 10px; height: 20px; width: 80px;' type='text' maxlength='12' name='mobilenum' value="<%=celNum%>" onKeyUp="javascript:return maskMe(this.value,this,'3,7','-');" onBlur="javascript:return maskMe(this.value,this,'3,7','-');"></font>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
							<td valign='center'><font size='1' face='trebuchet MS'><u>Emergency Contact Info:</u></font>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 250px;' name='emerinfo' maxlength='49' value="<%=emerinfo%>"</td>
						</td>
					</tr>
					<tr>
						<td>
								&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							<font size='1' face='trebuchet MS'><u>Relationship:</u>&nbsp;<input style='font-size: 10px; height: 20px; width: 80px;' type='text' maxlength='49' name='emerrel' value="<%=emerrel%>" ></font>
							&nbsp;&nbsp;&nbsp;&nbsp;
							<font size='1' face='trebuchet MS'><u>Contact Number:</u>&nbsp;<input style='font-size: 10px; height: 20px; width: 80px;' type='text' maxlength='12' name='emerphone' value="<%=emerphone%>" onKeyUp="javascript:return maskMe(this.value,this,'3,7','-');" onBlur="javascript:return maskMe(this.value,this,'3,7','-');"></font>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td>
							<font size='1' face='trebuchet MS'><u>DOB:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 80px;' maxlength='10' name='DOB' value='<%=DOB%>' onKeyUp="javascript:return maskMe(this.value,this,'2,5','/');" onBlur="javascript:return maskMe(this.value,this,'2,5','/');"></font>
							<font size='1' face='trebuchet MS'>(mm/dd/yyyy)</font>
							<font size='1' face='trebuchet MS'><u>Gender:</u>&nbsp;
							<select style='font-size: 10px; height: 20px; width: 80px;' name='Gen'></font>
								<option>Male</option>
								<option>Female</option>
							</select>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td>
							<font size='1' face='trebuchet MS'><u>Directions:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 400px;' maxlength='100' name='Direct' value='<%=D%>'></font>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td>
							<font size='1' face='trebuchet MS'><u>Referral Date:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 80px;' name='RefDate' maxlength='10' value="<%=RD%>" onKeyUp="javascript:return maskMe(this.value,this,'2,5','/');" onBlur="javascript:return maskMe(this.value,this,'2,5','/');"></font>
							<font size='1' face='trebuchet MS'>(mm/dd/yyyy)</font>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td>
							<font size='1' face='trebuchet MS'><u>Consumer Start Date:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 80px;' name='StrtDte' maxlength='10' value="<%=SD%>" onKeyUp="javascript:return maskMe(this.value,this,'2,5','/');" onBlur="javascript:return maskMe(this.value,this,'2,5','/');"></font>
							<font size='1' face='trebuchet MS'>(mm/dd/yyyy)</font>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td>
							<font size='1' face='trebuchet MS'><u>Amendment Effective Date:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 80px;' name='EffDte' maxlength='10' value="<%=EfD%>" onKeyUp="javascript:return maskMe(this.value,this,'2,5','/');" onBlur="javascript:return maskMe(this.value,this,'2,5','/');"></font>
							<font size='1' face='trebuchet MS'>(mm/dd/yyyy)</font>
							<font size='1' face='trebuchet MS'><u>Amendment Expiration Date:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 80px;' name='EndDte' maxlength='10' value="<%=ED%>" onKeyUp="javascript:return maskMe(this.value,this,'2,5','/');" onBlur="javascript:return maskMe(this.value,this,'2,5','/');"></font>
							<font size='1' face='trebuchet MS'>(mm/dd/yyyy)</font>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td>
							<font size='1' face='trebuchet MS'><u>Amendment Signed:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 80px;' name='ASDte' maxlength='10' value="<%=AmS%>" onKeyUp="javascript:return maskMe(this.value,this,'2,5','/');" onBlur="javascript:return maskMe(this.value,this,'2,5','/');"></font>
							<font size='1' face='trebuchet MS'>(mm/dd/yyyy)</font>
							<font size='1' face='trebuchet MS'><u>Amendment Received:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 80px;' name='ARDte' maxlength='10' value="<%=AR%>" onKeyUp="javascript:return maskMe(this.value,this,'2,5','/');" onBlur="javascript:return maskMe(this.value,this,'2,5','/');"></font>
							<font size='1' face='trebuchet MS'>(mm/dd/yyyy)</font>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td>
							<font size='1' face='trebuchet MS'><u>Consumer End Date:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 80px;' name='TermDte' maxlength='10' value="<%=Tmd%>" onKeyUp="javascript:return maskMe(this.value,this,'2,5','/');" onBlur="javascript:return maskMe(this.value,this,'2,5','/');"></font>
							<font size='1' face='trebuchet MS'>(mm/dd/yyyy)</font>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td>
							<font size='1' face='trebuchet MS'><u>Current Care Plan Date:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 80px;' name='CareDte' maxlength='10' value="<%=CareDte%>" onKeyUp="javascript:return maskMe(this.value,this,'2,5','/');" onBlur="javascript:return maskMe(this.value,this,'2,5','/');"></font>
							<font size='1' face='trebuchet MS'>(mm/dd/yyyy)</font>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td>
							<font size='1' face='trebuchet MS'><u>Code:</u>&nbsp;
								<select style='font-size: 10px; height: 20px; width: 40px;' name='selcode' onchange='JavaScript:myRate();'></font>
								<option value='M' <%=mcode%> >M</option>
								<option value='P' <%=pcode%> >P</option>
								<option value='C' <%=ccode%> >C</option>
								<option value='A' <%=acode%> >A</option>
								<option value='V' <%=vcode%> >V</option>
							</select>
							<font size='1' face='trebuchet MS'><u>Rate:</u>&nbsp;
								<input type='text' style='font-size: 10px; height: 20px; width: 80px;' name='txtrate' maxlength='10' value="<%=tmprate%>" >
							</font>
						</td>
					</tr>
					<tr>
						<td align='left' colspan='2'>
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							<input type='checkbox' name='chkVAHM' value='1' <%=vahm%> onclick='hrschk(1);'>
							<font size='1' face='trebuchet MS'><u>VA-HM</u>&nbsp;
						    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
						    <font size='1' face='trebuchet MS'><u>Hours:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 80px;' name='hrshm' maxlength='6' value="<%=hrshm%>"></font>
						</td>
					</tr>
					<tr>
						<td align='left' colspan='2'>
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							<input type='checkbox' name='chkVAHA' value='1' <%=vaha%> onclick='hrschk(2);'>
							<font size='1' face='trebuchet MS'><u>VA-HA</u>&nbsp;
						    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
						    <font size='1' face='trebuchet MS'><u>Hours:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 80px;' name='hrsha' maxlength='6' value="<%=hrsha%>"></font>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td>
							<font size='1' face='trebuchet MS'><u>Private Pay Contract Hours:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 80px;' name='txtcon' maxlength='10' value="<%=contract%>"></font>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
						<tr>
						<td align='left'>
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							<% If CDv <> "" Then chk = "checked" %>
							<input type='checkbox' name='chkDrive' <%=chk%>>
							<font size='1' face='trebuchet MS'><u>Driving Part of Care Plan:</u>&nbsp;
						&nbsp;&nbsp;&nbsp;&nbsp;
							<font size='1' face='trebuchet MS'><u>Max Hours:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 80px;' maxlength='6' name='maxhrs' value="<%=MH%>"></font>
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
						    <font size='1' face='trebuchet MS'><u>Mileage Cap:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 80px;' name='txtmile' maxlength='6' value="<%=milecap%>"></font>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td>
							<font size='1' face='trebuchet MS'><u>Managed Care Company:</u></font>
							<select style='font-size: 10px; height: 20px; width: 150px;' name='selmmc'>
								<option value='0'></option>
								<%=strmcc%>
							</select>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
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
					<tr>
						<td valign='center'>
							<font size='1' face='trebuchet MS'><u>RIHCC:</u></font>
							<select style='font-size: 10px; height: 20px; width: 150px;'name='PMsel'>
								<option value='0'></option>
								<%=strPM%>
							</select>
							&nbsp;&nbsp;&nbsp;&nbsp;
							<font size='1' face='trebuchet MS'><u>Comments:</u>&nbsp;
							<textarea rows='2' name='Pcomments' cols="20" ><%=C%></textarea></td
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
				</table>
				<table border='0'>
					<tr>
						<td align='center' colspan='2'>
							<input type='button' value='Save' style='width: 110px;' class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'"  onclick='JavaScript:CDN_Edit();'>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
				</table>
				<!-- #include file="_boxdown.asp" -->
			</center>	
			
		</form>
	</body>
</html>
<%	Session("MSG") = "" 
	Session("CFiles") = ""
%>