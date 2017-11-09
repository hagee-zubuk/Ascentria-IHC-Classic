<%Language=VBScript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_security.asp" -->
<%
	'If Request("Worker") <> "" Then
	'	tmpMCNum = split(Request("worker"), " - ") 
	'	ID = tmpMCNum(0)
	'	Session("WID") = Request("Worker")
	'Else
	'	If Request("WID") <> "" Then Session("WID") = Request("WID")
	'End If
	If Request("Worker") <> "" Then
		WID = Request("Worker")
	ElseIf Request("WID") <> "" Then
		WID = Request("WID") 
	Else
		Session("MSG") = "Sorry. To maintain integrity of database, please Sign-in again."
		Response.Redirect "default.asp"
	End If
	FinOnly = "ReadOnly"
	If Session("lngType") = 1 Or Session("lngType") = 2 Then
		FinOnly = ""
	End If
	myFin = ""
	Lang = ""
	'langsel = Z_ListLanguages(34)
	'If UCase(Session("lngType")) = "0" Then myFin = "READONLY"
	'response.write "WID : " & request("wid")
	If WID <> "" Then
		Set tblWorker = Server.CreateObject("ADODB.Recordset")
		sqlWorker = "SELECT * FROM Worker_t WHERE Social_Security_Number = '" & WID & "' "
		tblWorker.Open sqlWorker, g_strCONN, 1, 3
			If Not tblWorker.EOF Then
				Index = tblWorker("Index")
				Session("Widx") = Index
				SSN = tblWorker("Social_Security_Number")
				Lname = tblWorker("Lname")
				Fname = tblWorker("Fname")
				Session("Wname") = lname & ", " & fname
				DOB = tblWorker("DOB")
				If tblWorker("Gender")= "Male" Then
					F = ""
					M = "selected"
				ElseIf tblWorker("Gender")= "Female" Then
					F = "selected"
					M = ""
				End If
				Addr = tblWorker("Address")
				cty = tblWorker("City")
				ste = tblWorker("State")
				zcode = tblWorker("Zip")
				MAddr = tblWorker("mAddress")
				Mcty = tblWorker("mCity")
				Mste = tblWorker("mState")
				Mzcode = tblWorker("mZip")
				FonNum = tblWorker("PhoneNo")
				CelNum = tblWorker("CellNo")
				Ema = tblWorker("eMail")
				DateHired = tblWorker("Date_Hired")
				TD = tblWorker("term_date")
				Act = ""
				InAct = ""
				Pap = ""
				If tblWorker("Status") = "Active" Then Act = "selected"
				If tblWorker("Status") = "InActive" Then InAct = "selected"
				If tblWorker("Status") = "Potential Applicant" Then Pap = "selected"
				'Stat = tblWorker("Status")
				SepCod = tblWorker("Sep_code")
				If tblWorker("Driver") = True Then Drive = "checked"
				If tblWorker("License_File") = True Then LF = "checked"
				'If tblWorker("More_Consumer") = True Then MCon = "checked"
				LN = tblWorker("LicenseNo")
				LED = tblWorker("LicenseExpDate")
				If tblWorker("AutoInsur") = True Then AI = "checked"
				IED = tblWorker("InsuranceExpdate")
				If tblWorker("More_Towns") = True Then MT = "checked" 
				Sal = Z_FormatNumber(tblWorker("Salary"), 2)
				tmpMan = tblWorker("manual")
				pm1 = tblWorker("pm1")
				pm2 = tblWorker("pm2")
				FNum = tblWorker("FileNum")
				If tblWorker("privatepay") = True Then pp = "checked"
				badge = tblWorker("badge")
				ubadge = tblWorker("ubadge")
				umid = tblWorker("umid")
				flt = ""
				if tblWorker("flt") then flt = "checked"
				cmt = tblWorker("comment")
				misd = tblWorker("misdemeanor")
				warn = tblWorker("warning")
				If tblWorker("prefcom") = 1 Then prefmail = "selected"
				If tblWorker("prefcom") = 2 Then prefemail = "selected"
				If tblWorker("prefcom") = 3 Then prefFon = "selected"
				If tblWorker("prefcom") = 4 Then preftxt = "selected"
						concounty = tblWorker("countid")
				if tblWorker("essentials") Then essentials = "checked"
				langsel = Z_ListLanguages(tblWorker("langid"))
			Else
				Session("MSG") = "Session has expired. Please sign in again."
				Response.Redirect "Default.asp"
			End If
		tblWorker.Close
		Set tblWorker = Nothing	
		Set tblTowns = Server.CreateObject("ADODB.RecordSet")
		sqlTowns = "SELECT * FROM W_Towns_t WHERE SSN = '" & WID & "' "
		tblTowns.Open sqlTowns, g_strCONN, 1, 3
		If Not tblTowns.EOF Then
			ctr = 0 
			Do Until tblTowns.EOF
				if Z_IsOdd(ctr) = true then 
					kulay = "#FFFAF0" 
				else 
					kulay = "#FFFFFF"
				end if
					strTowns = strTowns & "<tr bgcolor='" & kulay & "'><td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & _
						"<input type='checkbox' name='chk" & ctr & "' value='" & tblTowns("Index") & "'>" & _
						"<input type='text' style='font-size: 10px; height: 20px; width: 80px;' " & _
						"name='Town" & ctr & "' value='" & tblTowns("Towns") & "'></font></td></tr>"
					tblTowns.MoveNext
					ctr = ctr + 1
				Loop
		End If
		tblTowns.Close
		Set tblTowns = Nothing
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
			selcount = ""
			if concounty = rsPM("uid") then selcount = "selected" 
			strcount = strcount & "<option " &  selCount & " value='" & rsPM("uid") & "' >" & rsPM("county") & "</option>" 
			rsPM.MoveNext
		Loop
		rsPM.Close
		Set rsPM = Nothing
	Else
		Session("MSG") = "Sorry. To maintain integrity of database, please Sign-in again."
		Response.Redirect "default.asp"
	End If
%>
<html>
	<head>
		<title>LSS - In-Home Care - PCSP WORKER Details</title>
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
				document.frmWorDet.action = "A_W_Action.asp?page=1";
				document.frmWorDet.submit();
			}
			function WD_Del()
			{
				var ans = window.confirm("Delete Worker? Click Cancel to stop.");
				if (ans){
				document.frmWorDet.action = "A_Del.asp?act=5";
				document.frmWorDet.submit();
				}
			}
			
			function Sep()
			{
				if (document.frmWorDet.Stat.value == 2)
					{document.frmWorDet.SepCode.disabled = false;}
				else
					{document.frmWorDet.SepCode.disabled = true;}
			}
			function Driver()
			{
				if (document.frmWorDet.chkDrive.checked == true)
					{document.frmWorDet.chkonFile.disabled = false;
					document.frmWorDet.LisNo.disabled = false;
					document.frmWorDet.chkIns.disabled = false;
					document.frmWorDet.LisExpDte.disabled = false;
					document.frmWorDet.Insdate.disabled = false;}
				else
					{document.frmWorDet.chkonFile.disabled = true;
					document.frmWorDet.LisNo.disabled = true;
					document.frmWorDet.chkIns.disabled = true;
					document.frmWorDet.LisExpDte.disabled = true;
					document.frmWorDet.Insdate.disabled = true;}
			}
			function InterTown()
			{
				if (document.frmWorDet.chkTown.checked == true)
					{document.frmWorDet.Towns.disabled = false;}
				else
					{document.frmWorDet.Towns.disabled = true;}
					
			}
			function WD_Del_Towns()
			{
				document.frmWorDet.action = "A_W_Del.asp?page=0";
				document.frmWorDet.submit();
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
	<body bgcolor='white' LEFTMARGIN='0' TOPMARGIN='0' onload='JavaScript: Sep(); Driver(); InterTown();'>
		<form method='post' name='frmWorDet'>
			<!-- #include file="_boxup.asp" -->
			<!-- #include file="_NavHeader.asp" -->
			<table border='0' align='center'>
					<tr>
						<td colspan='4' align='center' width='500px'>
							<font size='2' face='trebuchet MS'><b><u>PCSP Worker Details</u></b></font>
							<font size='2' face='trebuchet MS'>[Details]</font>
							<a href='A_W_Files.asp?WID=<%=WID%>' style='text-decoration:none'><font size='1'>[Files]</font></a>
							<a href="A_W_Skills.asp?WID=<%=WID%>" style='text-decoration: none;'>
								<font color='blue' size='1' face='trebuchet MS'>[Skills]</font>
							</a>
							<a href="WorkCon.asp?WID=<%=WID%>" style='text-decoration: none;'>
								<font color='blue' size='1' face='trebuchet MS'>[List]</font>
							</a>
							<a href="A_W_log.asp?WID=<%=WID%>" style='text-decoration: none;'>
								<font color='blue' size='1' face='trebuchet MS'>[Log]</font>
							</a>
							<a href="A_W_misc.asp?WID=<%=WID%>" style='text-decoration: none;'>
								<font color='blue' size='1' face='trebuchet MS'>[Violations]</font>
							</a>
							<a href="wimport.asp?WID=<%=WID%>" style='text-decoration: none;'>
								<font color='blue' size='1' face='trebuchet MS'>[Uploads]</font>
							</a>
						</td>
					</tr>
					<tr><td colspan='2' align='center'><font face='trebuchet MS' color='red' size='1'>&nbsp;<%=Session("MSG")%>&nbsp;</font></td></tr>
						<% If Session("lngType") = 1 Or Session("lngType") = 2 Or session("UserID") = 893 Then %>
					<tr>
						<td>
							<font size='1' face='trebuchet MS'><u>SSN:</u>&nbsp;
							<input style='font-size: 10px; height: 20px; width: 80px;' type='text' readonly name='SSN' value="<%=SSN%>"></td>
							<input type='hidden' name='WID' value="<%=SSN%>">
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<% Else %>
						<input type='hidden' name='SSN' value="<%=SSN%>">
						<input type='hidden' name='WID' value="<%=SSN%>">
					<% End If %>
					<tr>
						<td>
							<font size='1' face='trebuchet MS'><u>Name:</u>&nbsp;<input style='font-size: 10px; height: 20px; width: 90px;' type='text' name='lname' maxlength='20' value="<%=Lname%>">,&nbsp; 
							<input style='font-size: 10px; height: 20px; width: 90px;' type='text' name='fname' maxlength='20' value="<%=Fname%>">
							<font size='1' face='trebuchet MS'>(Last Name, First Name)</font>
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							<input type='checkbox' name='chkflt' <%=flt%> value='1'><u>Float Worker</u>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td>
							<font size='1' face='trebuchet MS'><u>Residence Address:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 250px;' name='Addr' maxlength='49' value="<%=Addr%>"></font>
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
							<input type='text' style='font-size: 10px; height: 20px; width: 250px;' name='MAddr' maxlength='49' value="<%=MAddr%>"></font>
							<input type='checkbox' name='chkMail'><font size='1' face='trebuchet MS'><u>Same as Residence:</u></font>
						</td>
					</tr>
					<tr>
						<td colspan='2'>
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							<font size='1' face='trebuchet MS'><u>City:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 100px;' name='Mcty' maxlength='49' value="<%=Mcty%>">
							<font size='1' face='trebuchet MS'><u>State:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 50px;' name='Mste' maxlength='2' value="<%=Mste%>">
							<font size='1' face='trebuchet MS'><u>Zip Code:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 100px;' name='Mzcode' maxlength='10' value="<%=Mzcode%>" onKeyUp="javascript:return maskMe(this.value,this,'5','-');" onBlur="javascript:return maskMe(this.value,this,'5','-');">
						</td>
					</tr>
						<tr><td>&nbsp;</td></tr>
					<tr>
						<td>
							<font size='1' face='trebuchet MS'><u>County:</u>&nbsp;
								<select style='font-size: 10px; height: 20px; width: 100px;' name='selcount'></font>
								<option value="0">&nbsp;</option>
								<%=strcount%>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
						<tr><td colspan='2'>
							<font size='1' face='trebuchet MS'><u>Gender:</u>&nbsp;
							<select style='font-size: 10px; height: 20px; width: 80px;' name='Gen' value="<%=Gender%>"></font>
								<option <%=M%>>Male</option>
								<option <%=F%>>Female</option>
							</select>
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							<font size='1' face='trebuchet MS'><u>Language:</u>&nbsp;</font>
							<select style='font-size: 10px; height: 20px; width: 200px;' name='langid' id='langid'>
								<%=langsel%>
							</select>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td colspan='12'>
							<font size='1' face='trebuchet MS'><u>DOB:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 80px;' maxlength='10' name='DOB' value="<%=DOB%>" onKeyUp="javascript:return maskMe(this.value,this,'2,5','/');" onBlur="javascript:return maskMe(this.value,this,'2,5','/');"></font>
							<font size='1' face='trebuchet MS'>(mm/dd/yyyy)</font>
							<font size='1' face='trebuchet MS'><u>Phone No:</u>&nbsp;<input style='font-size: 10px; height: 20px; width: 80px;' type='text' maxlength='12' name='PhoneNo' value="<%=FonNum%>" onKeyUp="javascript:return maskMe(this.value,this,'3,7','-');" onBlur="javascript:return maskMe(this.value,this,'3,7','-');"></font>
							<font size='1' face='trebuchet MS'><u>Cell No:</u>&nbsp;<input style='font-size: 10px; height: 20px; width: 80px;' type='text' name='CellNo' maxlength='12' value="<%=CelNum%>" onKeyUp="javascript:return maskMe(this.value,this,'3,7','-');" onBlur="javascript:return maskMe(this.value,this,'3,7','-');"></font>
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
							<input type='text' style='font-size: 10px; height: 20px; width: 80px;' name='DateHired' maxlength='10' value="<%=DateHired%>" onKeyUp="javascript:return maskMe(this.value,this,'2,5','/');" onBlur="javascript:return maskMe(this.value,this,'2,5','/');"></font>
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
								<option <%=Act%> value='1'>Active</option>
								<option <%=InAct%> value='2'>Inactive</option>
								<option <%=Pap%> value='3'>Potential Applicant</option>
							</select>
							<font size='1' face='trebuchet MS'><u>Separation Code:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 80px;' maxlength='4' name='SepCode' value="<%=SepCod%>"></font>
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
							<font size='1' face='trebuchet MS'><u>Badge Number:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 80px;' maxlength='6' name='txtBadge' value="<%=badge%>"></font>
							&nbsp;&nbsp;&nbsp;&nbsp;
							<font size='1' face='trebuchet MS'><u>Comments:</u>&nbsp;
							<textarea rows='2' name='Pcomments' cols="20" ><%=cmt%></textarea>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td colspan='12'>
							<font size='1' face='trebuchet MS'><u>UtliPro Badge ID:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 80px;' maxlength='6' name='txtuBadge' value="<%=ubadge%>"></font>
							&nbsp;&nbsp;&nbsp;&nbsp;
							<font size='1' face='trebuchet MS'><u>UltiPro Manager ID:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 80px;' maxlength='6' name='txtumid' value="<%=umid%>"></font>
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
							<input type='checkbox' name='chkDrive' <%=Drive%> onclick='JavaScript:Driver();'>
							<font size='1' face='trebuchet MS'><u>Driver:</u>&nbsp;
						</td>
					</tr>
					<tr>
						<td align='left'>
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							<input type='checkbox' name='chkonFile' <%=LF%>>
							<font size='1' face='trebuchet MS'><u>Copy of License on File:</u>&nbsp;
						</td>
					</tr>
					<tr>
						<td align='left'>
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							<font size='1' face='trebuchet MS'><u>License #:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 100px;' maxlength='17' name='LisNo' value="<%=LN%>"></font>
						</td>
					</tr>
					<tr>
						<td align='left'>
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							<font size='1' face='trebuchet MS'><u>License Expiration Date:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 80px;' maxlength='10' name='LisExpDte' value="<%=LED%>" onKeyUp="javascript:return maskMe(this.value,this,'2,5','/');" onBlur="javascript:return maskMe(this.value,this,'2,5','/');"></font>
							<font size='1' face='trebuchet MS'>(mm/dd/yyyy)</font>
						</td>
					</tr>
					<tr>
						<td align='left'>
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							<input type='checkbox' name='chkIns' <%=AI%>>
							<font size='1' face='trebuchet MS'><u>Copy of Auto Insurance:</u>&nbsp;
						</td>
					</tr>
					<tr>
						<td align='left'>
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							<font size='1' face='trebuchet MS'><u>Insurance Expiration Date:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 80px;' name='Insdate' maxlength='10' value="<%=IED%>" onKeyUp="javascript:return maskMe(this.value,this,'2,5','/');" onBlur="javascript:return maskMe(this.value,this,'2,5','/');"></font>
							<font size='1' face='trebuchet MS'>(mm/dd/yyyy)</font>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td align='left'>
							<input type='checkbox' name='chkTown' <%=MT%> onclick='JavaScript:InterTown();'>
							<font size='1' face='trebuchet MS'><u>Interested in Working with More Consumer:</u>&nbsp;
						</td>
					</tr>
					<tr>
						<td>
							<font size='1' face='trebuchet MS'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<u>Towns:</u></font>
							<a href='JavaScript:WD_Del_Towns();' style='text-decoration:none'><font size='1' face='trebuchet MS'>[Delete]</font></a>	
							<font size='1' face='trebuchet MS'>*enter one Town at a time.</font>
						</td>
					</tr>
					&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
					<%=strTowns%>
					<tr>
						<td>
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							<input type='text' name='Towns' disabled maxlength='40' style='font-size: 10px; height: 20px; width: 80px;'></font>
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
					<tr>
					<td align='center' colspan='2'>
							<input type='button' value='Save' style='width: 110px;' class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='JavaScript:WD_Edit();'>
							<% If UCase(Session("lngType")) = "2" Then %>
								<input type='button' value='Delete' style='width: 110px;' class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='JavaScript:WD_Del();'>
							<% End If %>
						</td>
					</tr>
				</table>
				<!-- #include file="_boxdown.asp" -->
			</center>	
		</form>
	</body>
</html>
<% Session("MSG") = "" 
	Session("WFiles") = ""
%>