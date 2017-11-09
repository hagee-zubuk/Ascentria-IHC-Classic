<%Language=VBScript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_security.asp" -->
<%
	If Request("MNum") <> "" Then
		'Get Age
		Set tblAge = Server.CreateObject("ADODB.RecordSet")
		sqlAge = "SELECT DOB FROM [Consumer_t] WHERE [Medicaid_Number] = '" & Request("MNum") & "' "
		tblAge.Open sqlAge, g_strCONN, 1, 3
		If Not tblAge.EOF Then
			If tblAge("DOB") <> empty Then
				Age = DateDiff("yyyy", tblAge("DOB"), Now)
			Else
				Age = 0
			End If
		Else
			Age = 0
		End If
		tblAge.Close
		Set tblAge = Nothing
		'Rating
		Set tblHealth = Server.CreateObject("ADODB.Recordset")
		sqlHealth = "SELECT Top 1 * FROM C_Health_t WHERE Medicaid_Number = '" & Request("MNum") & "' ORDER BY datestamp DESC"
		tblHealth.Open sqlHealth, g_strCONN, 3, 1
			If Not tblHealth.EOF Then
				MCNum = tblHealth("Medicaid_Number")
				Rating = tblHealth("Rating")
				'Ambulation
				If tblHealth("Indept") = True Then 
					I = "checked"
				ElseIf tblHealth("Cane") = True Then 
					C = "checked"
				ElseIF tblHealth("Walker") = True Then
					W = "checked"
				ElseIF tblHealth("Walk") = True Then
					WW = "checked"
				ElseIF tblHealth("WheelChair") = True Then
					WC = "checked"
				End If
				'ADL
				If tblHealth("ADL_Indep") = True Then 
					AI = "checked"
				ElseIf tblHealth("Monitor") = True Then 
					M = "checked"
				ElseIF tblHealth("MinAss") = True Then
					MA = "checked"
				ElseIF tblHealth("Ass") = True Then
					A = "checked"
				ElseIF tblHealth("Complete") = True Then
					CC = "checked"
				End If
				'new
				If tblHealth("oxy") Then oxy = "checked"
				If tblHealth("gait") Then gait = "checked"
				If tblHealth("blance") Then blance = "checked"
				If tblHealth("fear") Then fear = "checked"
				If tblHealth("furni") Then furni = "checked"
				If tblHealth("noncompl") Then noncompl = "checked"
				If tblHealth("substance") Then substance = "checked"
				If tblHealth("chew") Then chew = "checked"
				If tblHealth("teeth") Then teeth = "checked"
				If tblHealth("specdiet") Then specdiet = "checked"
				whatdiet = tblHealth("whatdiet")
				If tblHealth("allergy") Then allergy = "checked"
				whatallergy = tblHealth("whatallergy")
				If tblHealth("diabetes") Then diabetes = "checked"
				If tblHealth("hyper") Then hyper = "checked"
				If tblHealth("hypo") Then hypo = "checked"
				If tblHealth("heartdis") Then heartdis = "checked"
				If tblHealth("heartfail") Then heartfail = "checked"
				If tblHealth("deepvain") Then deepvain = "checked"
				If tblHealth("hyperten") Then hyperten = "checked"
				If tblHealth("hypoten") Then hypoten = "checked"
				If tblHealth("neuro") Then neuro = "checked"
				If tblHealth("vasc") Then vasc = "checked"
				If tblHealth("gerd") Then gerd = "checked"
				If tblHealth("ulcers") Then ulcers = "checked"
				If tblHealth("arth") Then arth = "checked"
				If tblHealth("hip") Then hip = "checked"
				If tblHealth("limb") Then limb = "checked"
				If tblHealth("osteo") Then osteo = "checked"
				If tblHealth("bone") Then bone = "checked"
				If tblHealth("als") Then als = "checked"
				If tblHealth("cereb") Then cereb = "checked"
				If tblHealth("stroke") Then stroke = "checked"
				If tblHealth("dementia") Then dementia = "checked"
				If tblHealth("hunting") Then hunting = "checked"
				If tblHealth("sclerosis") Then sclerosis = "checked"
				If tblHealth("parapal") Then parapal = "checked"
				If tblHealth("park") Then park = "checked"
				If tblHealth("quadri") Then quadri = "checked"
				If tblHealth("seize") Then seize = "checked"
				If tblHealth("TIA") Then TIA = "checked"
				If tblHealth("Trauma") Then Trauma = "checked"
				If tblHealth("anx") Then anx = "checked"
				If tblHealth("depress") Then depress = "checked"
				If tblHealth("bipolar") Then bipolar = "checked"
				If tblHealth("schiz") Then schiz = "checked"
				If tblHealth("abuse") Then abuse = "checked"
				If tblHealth("otherpsy") Then otherpsy = "checked"
				If tblHealth("asthma") Then asthma = "checked"
				If tblHealth("copd") Then copd = "checked"
				If tblHealth("TB") Then TB = "checked"
				If tblHealth("cat") Then cat = "checked"
				If tblHealth("retin") Then retin = "checked"
				If tblHealth("glaucoma") Then glaucoma = "checked"
				If tblHealth("macular") Then macular = "checked"
				If tblHealth("hear") Then hear = "checked"
				If tblHealth("allergies") Then allergies = "checked"
				whatallergies = tblHealth("whatallergies")
				If tblHealth("anemia") Then anemia = "checked"
				If tblHealth("cancer") Then cancer = "checked"
				If tblHealth("devdis") Then devdis = "checked"
				If tblHealth("morbid") Then morbid = "checked"
				If tblHealth("renal") Then renal = "checked"
				If tblHealth("otherdiag") Then otherdiag = "checked"
				whatdiag = tblHealth("whatdiag")
				If tblHealth("housekeep") Then housekeep = "checked"
				If tblHealth("laundry") Then laundry = "checked"
				If tblHealth("meal") Then meal = "checked"
				If tblHealth("grocery") Then grocery = "checked"
				If tblHealth("dress") Then dress = "checked"
				If tblHealth("eat") Then eat = "checked"
				If tblHealth("asstwalk") Then asstwalk = "checked"
				If tblHealth("asstwheel") Then asstwheel = "checked"
				If tblHealth("asstmotor") Then asstmotor = "checked"
				If tblHealth("commeal") Then commeal = "checked"
				If tblHealth("medical") Then medical = "checked"
				If tblHealth("shower") Then shower = "checked"
				If tblHealth("tub") Then tub = "checked"
				If tblHealth("oral") Then oral = "checked"
				If tblHealth("commode") Then commode = "checked"
				If tblHealth("sit") Then sit = "checked"
				If tblHealth("medication") Then medication = "checked"
				If tblHealth("undress") Then undress = "checked"
				If tblHealth("shampoosink") Then shampoosink = "checked"
				If tblHealth("oralcare") Then oralcare = "checked"
				If tblHealth("massage") Then massage = "checked"
				If tblHealth("shampoobed") Then shampoobed = "checked"
				If tblHealth("shave") Then shave = "checked"
				If tblHealth("bedbath") Then bedbath = "checked"
				If tblHealth("bedpan") Then bedpan = "checked"
				If tblHealth("ptexer") Then ptexer = "checked"
				If tblHealth("hoyer") Then hoyer = "checked"
				If tblHealth("eye") Then eye = "checked"
				If tblHealth("transferbelt") Then transferbelt = "checked"
				If tblHealth("alz") Then alz = "checked"
				If tblHealth("Incontinence") Then Incontinence = "checked"
				If tblHealth("Hospice") Then Hospice = "checked"
				If tblHealth("lotion") Then lotion = "checked"
			Else
				I = "Checked"
				AI = "Checked"
				'tblHealth.AddNew
				'tblHealth("Medicaid_Number") = Request("MNum")
				'tblHealth.Update
			End If
		tblHealth.Close
		Set tblHealth = Nothing	
		Set rsDiag = Server.CreateObject("ADODB.RecordSet")
		sqlDiag = "SELECT * FROM C_Diagnosis_t WHERE Medicaid_Number ='" & Request("MNum") & "' "
		rsDiag.Open sqlDiag, g_strCONN, 1, 3
		ctr = 0
		Do Until rsDiag.EOF
			If Z_IsOdd(ctr) = True Then 
				kulay = "#FFFAF0" 
			Else 
				kulay = "#FFFFFF"
			End If
			strDiag = strDiag & "<tr bgcolor='" & kulay & "'><td align='center'><input type='checkbox' name='chkDiag" & ctr & _
				"' value='" & rsDiag("Index") & "'></td><td><input type='text' style='font-size: 10px; height: 20px; " & _
				"width: 150px;' readonly value='" & rsDiag("Diagnosis") & "'></font></td></tr>" & vbCrLf
			rsDiag.MoveNext
			ctr = ctr + 1
		Loop
		rsDiag.Close
		Set rsDiag = Nothing
	Else
		Session("MSG") = "Sorry. To maintain integrity of database, please Sign-in again."
		Response.Redirect "default.asp"
	End If
%>
<html>
	<head>
		<title>LSS - In-Home Care - Consumer Details - Health</title>
		<script language='JavaScript'>
		function DiagDel()
		{
			document.frmConDetHel.action = "A_C_Action.asp?page=5";
			document.frmConDetHel.submit();
		} 	
		function SaveMe() {
			strmsg = "";
			if (document.frmConDetHel.specdiet.checked == true && document.frmConDetHel.whatdiet.value == '') {
				strmsg = "\nPlease Input Special Diet.";
			}
			if (document.frmConDetHel.allergy.checked == true && document.frmConDetHel.whatallergy.value == '') {
				strmsg = strmsg + "\nPlease Input Food allergies.";
			}
			if (document.frmConDetHel.allergies.checked == true && document.frmConDetHel.whatallergies.value == '') {
				strmsg = strmsg + "\nPlease Input Allergies.";
			}
			if (strmsg == '') {
				document.frmConDetHel.action = "A_C_Action.asp?page=3";
				document.frmConDetHel.submit();
			}
			else {
				alert("ERROR:" + strmsg);
				return;
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
		<form method='post' name='frmConDetHel' action='A_C_Action.asp?page=3'>
			<!-- #include file="_boxup.asp" -->
			<!-- #include file="_NavHeader.asp" -->
			<br>
			<center>
				<table border='0'>
					<tr>
						<td colspan='4' align='center' width='500px'>
							<font size='2' face='trebuchet MS'><b><u>Consumer Details - Health</u></b></font>
							<a href='A_Consumer.asp?MNum=<%=Request("MNum")%>' style='text-decoration:none'><font size='1' face='trebuchet MS'>[Details]</font></a>
							<a href='A_C_Status.asp?MNum=<%=Request("MNum")%>' style='text-decoration:none'><font size='1' face='trebuchet MS'>[Status]</font></a>
							<font size='2' face='trebuchet MS'>[Health]</font>
							<a href='A_C_Files.asp?MNum=<%=Request("MNum")%>' style='text-decoration:none'><font size='1' face='trebuchet MS'>[Files]</font></a>
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
					<!--<tr>
						<td>
							<font size='1' face='trebuchet MS'><u>Rating:</u>&nbsp;
							<input type='text' readonly style='font-size: 10px; height: 20px; width: 80px;' name='Rate' value="<%=Rating%>"></font>
							<font size='1' face='trebuchet MS'>* Save entries to compute for rating. </font>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>//-->
					<tr>
						<td>
							<font size='1' face='trebuchet MS'><u>Age:</u>&nbsp;
							<input type='text' readonly style='font-size: 10px; height: 20px; width: 80px;' name='Age' value="<%=Age%>"></font>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td width='410px'>
							<table border='0'>
								<tr>
									<td>
										<font size='1' face='trebuchet MS'><u>Ambulation:</u>&nbsp;
									</td>
								</tr>
								<tr>
									<td><font size='1' face='trebuchet MS'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<u>Independent:</u></td>
									<td align='center'>
										<input type='radio' <%=I%> name='Amb' value='1'>
									</td>
								</tr>
								<tr>
									<td><font size='1' face='trebuchet MS'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<u>Cane:</u></td>
									<td align='center'>
										<input type='radio' <%=C%> name='Amb' value='2'>
									</td>
								</tr>
								<tr>
									<td><font size='1' face='trebuchet MS'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<u>Walker:</u></td>
									<td align='center'>
										<input type='radio' <%=W%> name='Amb' value='3'>
									</td>
								</tr>
								<tr>
									<td><font size='1' face='trebuchet MS'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<u>Walk/Wheel:</u></td>
									<td align='center'>
										<input type='radio' <%=WW%> name='Amb' value='4'>
									</td>
								</tr>	
								<tr>
									<td><font size='1' face='trebuchet MS'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<u>WheelChair:</u></td>
									<td align='center'>
										<input type='radio' <%=WC%> name='Amb' value='5'>
									</td>
								</tr>
							</td>
							</tr>
						</table>
					</td>
					<td colspan='2'>
							<table border='0'>
								<tr>
									<td colspan='2'>
										<font size='1' face='trebuchet MS'><u>ADL:</u>&nbsp;
									</td>
								</tr>
								<tr>
									<td><font size='1' face='trebuchet MS'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<u>Independent:</u></td>
									<td align='center'>
										<input type='radio' name='ADL' <%=AI%> value='1'>
									</td>
								</tr>
								<tr>
									<td><font size='1' face='trebuchet MS'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<u>Monitor:</u></td>
									<td align='center'>
										<input type='radio' <%=M%> name='ADL' value='2'>
									</td>
								</tr>
								<tr>
									<td><font size='1' face='trebuchet MS'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<u>Min. Assistance:</u></td>
									<td align='center'>
										<input type='radio' <%=MA%> name='ADL' value='3'>
									</td>
								</tr>
								<tr>
									<td><font size='1' face='trebuchet MS'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<u>Assistance:</u></td>
									<td align='center'>
										<input type='radio' <%=A%> name='ADL' value='4'>
									</td>
								</tr>	
								<tr>
									<td><font size='1' face='trebuchet MS'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<u>Complete Care:</u></td>
									<td align='center'>
										<input type='radio' <%=CC%> name='ADL' value='5'>
									</td>
								</tr>
							</td>
							</tr>
						</table>
					</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td colspan='4'>
							<table border='0'>
								<tr>
									<td colspan='2'>
										<font size='1' face='trebuchet MS'><u>Other Qualifying Conditions:</u>&nbsp;
									</td>
								</tr>
								<tr>
									<td><font size='1' face='trebuchet MS'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<u>Oxygen Use:</u></td>
									<td align='center'>
										<input type='checkbox' value='1' <%=oxy%> name='oxy'>
									</td>
								</tr>
								<tr>
									<td><font size='1' face='trebuchet MS'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<u>Fall Risk:</u></td>
									<td align='center'>
										&nbsp;
									</td>
								</tr>
								<tr>
									<td><font size='1' face='trebuchet MS'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<u>Unsteady Gait:</u></td>
									<td align='center'>
										<input type='checkbox' value='1' <%=gait%> name='gait'>
									</td>
									<td><font size='1' face='trebuchet MS'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<u>Furniture Walking:</u></td>
									<td align='center'>
										<input type='checkbox' value='1' <%=furni%> name='furni'>
									</td>
								</tr>
								<tr>
									<td><font size='1' face='trebuchet MS'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<u>Balance problems when standing:</u></td>
									<td align='center'>
										<input type='checkbox' value='1' <%=blance%> name='blance'>
									</td>
									<td><font size='1' face='trebuchet MS'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<u>Non-compliant with Assistive Devices:</u></td>
									<td align='center'>
										<input type='checkbox' value='1' <%=noncompl%> name='noncompl'>
									</td>
								</tr>
								<tr>
									<td><font size='1' face='trebuchet MS'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<u>Limits activities because fearful of falling:</u></td>
									<td align='center'>
										<input type='checkbox' value='1' <%=fear%> name='fear'>
									</td>
									<td><font size='1' face='trebuchet MS'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<u>Substances or drug use as a contributing factor:</u></td>
									<td align='center'>
										<input type='checkbox' value='1' <%=substance%> name='substance'>
									</td>
								</tr>
								<tr>
									<td><font size='1' face='trebuchet MS'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<u>Nutritional Problems/Approaches:</u></td>
									<td align='center'>
										&nbsp;
									</td>
								</tr>
								<tr>
									<td><font size='1' face='trebuchet MS'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<u>Chewing or swallowing problem:</u></td>
									<td align='center'>
										<input type='checkbox' value='1' <%=chew%> name='chew'>
									</td>
									<td><font size='1' face='trebuchet MS'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<u>Missing teeth or dentures:</u></td>
									<td align='center'>
										<input type='checkbox' value='1' <%=teeth%> name='teeth'>
									</td>
								</tr>
								<tr>
									<td><font size='1' face='trebuchet MS'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<u>Special Diet:</u>
										<input type='text' style='font-size: 10px; height: 20px; width: 100px;' maxlength='200' value='<%=whatdiet%>' name='whatdiet'>
									</td>
									<td align='center'>
										<input type='checkbox' value='1' <%=specdiet%> name='specdiet'>
									</td>
									<td><font size='1' face='trebuchet MS'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<u>Food allergies:</u>
										<input type='text' style='font-size: 10px; height: 20px; width: 100px;' maxlength='200' value='<%=whatallergy%>' name='whatallergy'>
										</td>
									<td align='center'>
										<input type='checkbox' value='1' <%=allergy%> name='allergy'>
									</td>
								</tr>
								<tr><td>&nbsp;</td></tr>
								<tr>
									<td colspan='2'>
										<font size='1' face='trebuchet MS'><u>Diagnosis:</u>&nbsp;
									</td>
								</tr>
								<tr>
									<td><font size='1' face='trebuchet MS'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<u>Endocrine/Metabolic:</u></td>
									<td align='center'>
										&nbsp;
									</td>
								</tr>
								<tr>
									<td><font size='1' face='trebuchet MS'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<u>Diabetes:</u></td>
									<td align='center'>
										<input type='checkbox' value='1' <%=diabetes%> name='diabetes'>
									</td>
									<td><font size='1' face='trebuchet MS'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<u>Hyperthyroidism:</u></td>
									<td align='center'>
										<input type='checkbox' value='1' <%=hyper%> name='hyper'>
									</td>
								</tr>
								<tr>
									<td><font size='1' face='trebuchet MS'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<u>Hypothyroidism:</u></td>
									<td align='center'>
										<input type='checkbox' value='1' <%=hypo%> name='hypo'>
									</td>
								</tr>
								<tr>
									<td><font size='1' face='trebuchet MS'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<u>Heart/Circulation:</u></td>
									<td align='center'>
										&nbsp;
									</td>
								</tr>
								<tr>
									<td><font size='1' face='trebuchet MS'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<u>Arteriosclerotic heart disease:</u></td>
									<td align='center'>
										<input type='checkbox' value='1' <%=heartdis%> name='heartdis'>
									</td>
									<td><font size='1' face='trebuchet MS'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<u>Congestive Heart Failure:</u></td>
									<td align='center'>
										<input type='checkbox' value='1' <%=heartfail%> name='heartfail'>
									</td>
								</tr>
								<tr>
									<td><font size='1' face='trebuchet MS'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<u>Deep Vein Thrombosis:</u></td>
									<td align='center'>
										<input type='checkbox' value='1' <%=deepvain%> name='deepvain'>
									</td>
									<td><font size='1' face='trebuchet MS'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<u>Hypertension:</u></td>
									<td align='center'>
										<input type='checkbox' value='1' <%=hyperten%> name='hyperten'>
									</td>
								</tr>
								<tr>
									<td><font size='1' face='trebuchet MS'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<u>Hypotension:</u></td>
									<td align='center'>
										<input type='checkbox' value='1' <%=hypoten%> name='hypoten'>
									</td>
									<td><font size='1' face='trebuchet MS'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<u>Neuropathy:</u></td>
									<td align='center'>
										<input type='checkbox' value='1' <%=neuro%> name='neuro'>
									</td>
								</tr>
								<tr>
									<td><font size='1' face='trebuchet MS'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<u>Peripheral Vascular Disease:</u></td>
									<td align='center'>
										<input type='checkbox' value='1' <%=vasc%> name='vasc'>
									</td>
								</tr>
								<tr>
									<td><font size='1' face='trebuchet MS'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<u>Gastric:</u></td>
									<td align='center'>
										&nbsp;
									</td>
								</tr>
								<tr>
									<td><font size='1' face='trebuchet MS'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<u>Gerd:</u></td>
									<td align='center'>
										<input type='checkbox' value='1' <%=Gerd%> name='Gerd'>
									</td>
									<td><font size='1' face='trebuchet MS'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<u>Ulcers:</u></td>
									<td align='center'>
										<input type='checkbox' value='1' <%=Ulcers%> name='Ulcers'>
									</td>
								</tr>
								<tr>
									<td><font size='1' face='trebuchet MS'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<u>Musculoskeletal:</u></td>
									<td align='center'>
										&nbsp;
									</td>
								</tr>
								<tr>
									<td><font size='1' face='trebuchet MS'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<u>Arthritis:</u></td>
									<td align='center'>
										<input type='checkbox' value='1' <%=arth%> name='arth'>
									</td>
									<td><font size='1' face='trebuchet MS'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<u>Hip fracture:</u></td>
									<td align='center'>
										<input type='checkbox' value='1' <%=hip%> name='hip'>
									</td>
								</tr>
								<tr>
									<td><font size='1' face='trebuchet MS'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<u>Missing Limb (e.g. amputation):</u></td>
									<td align='center'>
										<input type='checkbox' value='1' <%=limb%> name='limb'>
									</td>
									<td><font size='1' face='trebuchet MS'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<u>Osteoporosis:</u></td>
									<td align='center'>
										<input type='checkbox' value='1' <%=osteo%> name='osteo'>
									</td>
								</tr>
								<tr>
									<td><font size='1' face='trebuchet MS'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<u>Pathological bone fracture:</u></td>
									<td align='center'>
										<input type='checkbox' value='1' <%=bone%> name='bone'>
									</td>
								</tr>
								<tr>
									<td><font size='1' face='trebuchet MS'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<u>Neurological:</u></td>
									<td align='center'>
										&nbsp;
									</td>
								</tr>
								<tr>
									<td><font size='1' face='trebuchet MS'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<u>ALS:</u></td>
									<td align='center'>
										<input type='checkbox' value='1' <%=als%> name='als'>
									</td>
									<td><font size='1' face='trebuchet MS'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<u>Cerebral Palsy:</u></td>
									<td align='center'>
										<input type='checkbox' value='1' <%=cereb%> name='cereb'>
									</td>
								</tr>
								<tr>
									<td><font size='1' face='trebuchet MS'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<u>Stroke:</u></td>
									<td align='center'>
										<input type='checkbox' value='1' <%=stroke%> name='stroke'>
									</td>
									<td><font size='1' face='trebuchet MS'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<u>Dementia/Alzheimer’s:</u></td>
									<td align='center'>
										<input type='checkbox' value='1' <%=dementia%> name='dementia'>
									</td>
								</tr>
								<tr>
									<td><font size='1' face='trebuchet MS'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<u>Huntington's:</u></td>
									<td align='center'>
										<input type='checkbox' value='1' <%=hunting%> name='hunting'>
									</td>
									<td><font size='1' face='trebuchet MS'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<u>Multiple Sclerosis:</u></td>
									<td align='center'>
										<input type='checkbox' value='1' <%=sclerosis%> name='sclerosis'>
									</td>
								</tr>
								<tr>
									<td><font size='1' face='trebuchet MS'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<u>Paraplegia:</u></td>
									<td align='center'>
										<input type='checkbox' value='1' <%=parapal%> name='parapal'>
									</td>
									<td><font size='1' face='trebuchet MS'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<u>Parkinson’s Disease:</u></td>
									<td align='center'>
										<input type='checkbox' value='1' <%=park%> name='park'>
									</td>
								</tr>
								<tr>
									<td><font size='1' face='trebuchet MS'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<u>Quadriplegia:</u></td>
									<td align='center'>
										<input type='checkbox' value='1' <%=quadri%> name='quadri'>
									</td>
									<td><font size='1' face='trebuchet MS'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<u>Seizure Disorder:</u></td>
									<td align='center'>
										<input type='checkbox' value='1' <%=seize%> name='seize'>
									</td>
								</tr>
								<tr>
									<td><font size='1' face='trebuchet MS'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<u>Transient Ischemic Attack (TIA):</u></td>
									<td align='center'>
										<input type='checkbox' value='1' <%=TIA%> name='TIA'>
									</td>
									<td><font size='1' face='trebuchet MS'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<u>Traumatic Brain Injury:</u></td>
									<td align='center'>
										<input type='checkbox' value='1' <%=Trauma%> name='Trauma'>
									</td>
								</tr>
								<tr>
									<td><font size='1' face='trebuchet MS'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<u>Psychiatric/Mood:</u></td>
									<td align='center'>
										&nbsp;
									</td>
								</tr>
								<tr>
									<td><font size='1' face='trebuchet MS'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<u>Anxiety Disorder:</u></td>
									<td align='center'>
										<input type='checkbox' value='1' <%=anx%> name='anx'>
									</td>
									<td><font size='1' face='trebuchet MS'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<u>Depression:</u></td>
									<td align='center'>
										<input type='checkbox' value='1' <%=depress%> name='depress'>
									</td>
								</tr>
								<tr>
									<td><font size='1' face='trebuchet MS'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<u>Bipolar Disorder:</u></td>
									<td align='center'>
										<input type='checkbox' value='1' <%=bipolar%> name='bipolar'>
									</td>
									<td><font size='1' face='trebuchet MS'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<u>Schizophrenia:</u></td>
									<td align='center'>
										<input type='checkbox' value='1' <%=schiz%> name='schiz'>
									</td>
								</tr>
								<tr>
									<td><font size='1' face='trebuchet MS'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<u>Substance Abuse (alcohol or drug):</u></td>
									<td align='center'>
										<input type='checkbox' value='1' <%=abuse%> name='abuse'>
									</td>
									<td><font size='1' face='trebuchet MS'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<u>Other psychiatric diagnosis:</u></td>
									<td align='center'>
										<input type='checkbox' value='1' <%=otherpsy%> name='otherpsy'>
									</td>
								</tr>
								<tr>
									<td><font size='1' face='trebuchet MS'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<u>Pulmonary:</u></td>
									<td align='center'>
										&nbsp;
									</td>
								</tr>
								<tr>
									<td><font size='1' face='trebuchet MS'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<u>Asthma:</u></td>
									<td align='center'>
										<input type='checkbox' value='1' <%=asthma%> name='asthma'>
									</td>
									<td><font size='1' face='trebuchet MS'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<u>Emphysema/COPD:</u></td>
									<td align='center'>
										<input type='checkbox' value='1' <%=copd%> name='copd'>
									</td>
								</tr>
								<tr>
									<td><font size='1' face='trebuchet MS'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<u>Tuberculosis-TB:</u></td>
									<td align='center'>
										<input type='checkbox' value='1' <%=TB%> name='TB'>
									</td>
								</tr>
								<tr>
									<td><font size='1' face='trebuchet MS'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<u>Sensory:</u></td>
									<td align='center'>
										&nbsp;
									</td>
								</tr>
								<tr>
									<td><font size='1' face='trebuchet MS'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<u>Cataracts:</u></td>
									<td align='center'>
										<input type='checkbox' value='1' <%=cat%> name='cat'>
									</td>
									<td><font size='1' face='trebuchet MS'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<u>Diabetic retinopathy:</u></td>
									<td align='center'>
										<input type='checkbox' value='1' <%=retin%> name='retin'>
									</td>
								</tr>
								<tr>
									<td><font size='1' face='trebuchet MS'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<u>Glaucoma:</u></td>
									<td align='center'>
										<input type='checkbox' value='1' <%=Glaucoma%> name='Glaucoma'>
									</td>
									<td><font size='1' face='trebuchet MS'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<u>Macular Degeneration:</u></td>
									<td align='center'>
										<input type='checkbox' value='1' <%=macular%> name='macular'>
									</td>
								</tr>
								<tr>
									<td><font size='1' face='trebuchet MS'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<u>Hearing Impairment:</u></td>
									<td align='center'>
										<input type='checkbox' value='1' <%=hear%> name='hear'>
									</td>
								</tr>
								<tr>
									<td><font size='1' face='trebuchet MS'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<u>Other:</u></td>
									<td align='center'>
										&nbsp;
									</td>
								</tr>
								<tr>
									<td><font size='1' face='trebuchet MS'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<u>Allergies, specify:</u>
										<input type='text' style='font-size: 10px; height: 20px; width: 100px;' maxlength='200' value='<%=whatallergies%>' name='whatallergies'>
									</td>
									<td align='center'>
										<input type='checkbox' value='1' <%=allergies%> name='allergies'>
									</td>
									<td><font size='1' face='trebuchet MS'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<u>Anemia:</u></td>
									<td align='center'>
										<input type='checkbox' value='1' <%=anemia%> name='anemia'>
									</td>
								</tr>
								<tr>
									<td><font size='1' face='trebuchet MS'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<u>Cancer:</u></td>
									<td align='center'>
										<input type='checkbox' value='1' <%=Cancer%> name='Cancer'>
									</td>
									<td><font size='1' face='trebuchet MS'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<u>Developmental Disability:</u></td>
									<td align='center'>
										<input type='checkbox' value='1' <%=devdis%> name='devdis'>
									</td>
								</tr>
								<tr>
									<td><font size='1' face='trebuchet MS'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<u>Morbid Obesity:</u></td>
									<td align='center'>
										<input type='checkbox' value='1' <%=morbid%> name='morbid'>
									</td>
									<td><font size='1' face='trebuchet MS'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<u>Renal Failure:</u></td>
									<td align='center'>
										<input type='checkbox' value='1' <%=renal%> name='renal'>
									</td>
								</tr>
								<tr><td>&nbsp;</td></tr>
								<tr>
									<td><font size='1' face='trebuchet MS'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<u>Other Diagnosis:</u>
										<input type='text' style='font-size: 10px; height: 20px; width: 125px;' maxlength='200' value='<%=whatdiag%>' name='whatdiag'>
									</td>
									<td align='center'>
										<input type='checkbox' value='1' <%=otherdiag%> name='otherdiag'>
									</td>
								</tr>
							</table>
						</td>
						<td>&nbsp;</td>
					</tr>
					<!--<tr>
						<td colspan='2' valign='top'>
							<table border='0'>
								<tr>
									<td colspan='2'>
										<font size='1' face='trebuchet MS'><u>Other Diagnosis:</u></font>
										<a href='JavaScript: DiagDel();' style='text-decoration: none;'><font size='1' face='trebuchet MS'>[Delete]</font></a>
										<br><font size='1' face='trebuchet MS'>*enter one diagnosis at a time.</font>
									</td>
								</tr>
								<tr>
									<%=strDiag%>
									<td>&nbsp;&nbsp;&nbsp;</td>
									<td><input type='text' style='font-size: 10px; height: 20px; width: 150px;' maxlength='50' name='Diag'>
								</tr>
							</table>
						</td>
					</tr>//-->
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td colspan='4'>
							<table border='0'>
								<tr>
									<td colspan='2'>
										<font size='1' face='trebuchet MS'><u>Skills Criteria:</u>&nbsp;
									</td>
								</tr>
								<tr>
									<td colspan='4' align='left'><font size='1' face='trebuchet MS'><u>40 Bathing</u></font></td>
								</tr>
								<tr>
									<td align='center'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<input type='checkbox' value='1' <%=shower%> name='shower'>
									</td>
									<td><font size='1' face='trebuchet MS'>
										<u>Needs assistance with shower:</u>
									</td>
									<td align='center'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<input type='checkbox' value='1' <%=tub%> name='tub'>
									</td>
									<td><font size='1' face='trebuchet MS'>
										<u>Needs assistance with tub bath:</u>
									</td>
								</tr>
								<tr>
									<td align='center'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<input type='checkbox' value='1' <%=shampoosink%> name='shampoosink'>
									</td>
									<td><font size='1' face='trebuchet MS'>
										<u>Shampoo in sink:</u>
									</td>
									<td align='center'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<input type='checkbox' value='1' <%=shampoobed%> name='shampoobed'>
									</td>
									<td><font size='1' face='trebuchet MS'>
										<u>Shampoo in bed needed:</u>
									</td>
								</tr>
								<tr>
									<td align='center'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<input type='checkbox' value='1' <%=bedbath%> name='bedbath'>
									</td>
									<td><font size='1' face='trebuchet MS'>
										<u>Bed bath needed:</u>
									</td>
								</tr>
								<tr>
									<td colspan='4' align='left'><font size='1' face='trebuchet MS'><u>41 Dressing</u></font></td>
								</tr>
								<tr>
									<td align='center'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<input type='checkbox' value='1' <%=dress%> name='dress'>
									</td>
									<td><font size='1' face='trebuchet MS'>
										<u>Needs assistance with dressing or undressing:</u>
									</td>
									<td align='center'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<input type='checkbox' value='1' <%=undress%> name='undress'>
									</td>
									<td><font size='1' face='trebuchet MS'>
										<u>Dressing or undressing:</u>
									</td>
								</tr>
								<tr>
									<td colspan='4' align='left'><font size='1' face='trebuchet MS'><u>42 Transfer / Mobility</u></font></td>
								</tr>
								<tr>
									<td align='center'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<input type='checkbox' value='1' <%=asstwalk%> name='asstwalk'>
									</td>
									<td><font size='1' face='trebuchet MS'>
										<u>Needs assistance with walker:</u>
									</td>
									<td align='center'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<input type='checkbox' value='1' <%=asstwheel%> name='asstwheel'>
									</td>
									<td><font size='1' face='trebuchet MS'>
										<u>Needs assistance with wheelchair:</u>
									</td>
								</tr>
								<tr>
									<td align='center'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<input type='checkbox' value='1' <%=asstmotor%> name='asstmotor'>
									</td>
									<td><font size='1' face='trebuchet MS'>
										<u>Needs assistance with motorized scooter:</u>
									</td>
									<td align='center'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<input type='checkbox' value='1' <%=sit%> name='sit'>
									</td>
									<td><font size='1' face='trebuchet MS'>
										<u>Assist with standard sit and transfer needed:</u>
									</td>
								</tr>
								<tr>
									<td align='center'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<input type='checkbox' value='1' <%=transferbelt%> name='transferbelt'>
									</td>
									<td><font size='1' face='trebuchet MS'>
										<u>Use of transfer belt needed:</u>
									</td>
									<td align='center'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<input type='checkbox' value='1' <%=hoyer%> name='hoyer'>
									</td>
									<td><font size='1' face='trebuchet MS'>
										<u>Hoyer Lift transfers required:</u>
									</td>
								</tr>
								<tr>
									<td colspan='4' align='left'><font size='1' face='trebuchet MS'><u>43 Toileting / Hygiene</u></font></td>
								</tr>
								<tr>
									<td align='center'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<input type='checkbox' value='1' <%=oral%> name='oral'>
									</td>
									<td><font size='1' face='trebuchet MS'>
										<u>Needs assistance with oral hygiene:</u>
									</td>
									<td align='center'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<input type='checkbox' value='1' <%=commode%> name='commode'>
									</td>
									<td><font size='1' face='trebuchet MS'>
										<u>Needs assistance with commode:</u>
									</td>
								</tr>
								<tr>
									<td align='center'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<input type='checkbox' value='1' <%=oralcare%> name='oralcare'>
									</td>
									<td><font size='1' face='trebuchet MS'>
										<u>Oral hygiene care needed:</u>
									</td>
									<td align='center'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<input type='checkbox' value='1' <%=massage%> name='massage'>
									</td>
									<td><font size='1' face='trebuchet MS'>
										<u>Massage, back rub, foot rub needed:</u>
									</td>
								</tr>
								<tr>
									<td align='center'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<input type='checkbox' value='1' <%=shave%> name='shave'>
									</td>
									<td><font size='1' face='trebuchet MS'>
										<u>Shave consumer:</u>
									</td>
									<td align='center'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<input type='checkbox' value='1' <%=bedpan%> name='bedpan'>
									</td>
									<td><font size='1' face='trebuchet MS'>
										<u>Assistance with bed pan:</u>
									</td>
								</tr>
								<tr>
									
									<td align='center'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<input type='checkbox' value='1' <%=eye%> name='eye'>
									</td>
									<td><font size='1' face='trebuchet MS'>
										<u>Instill eye drops or ointments:</u>
									</td>
									<td align='center'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<input type='checkbox' value='1' <%=Incontinence%> name='Incontinence'>
									</td>
									<td><font size='1' face='trebuchet MS'>
										<u>Incontinence care:</u>
									</td>
								</tr>
									<tr>
									
									<td align='center'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<input type='checkbox' value='1' <%=lotion%> name='lotion'>
									</td>
									<td><font size='1' face='trebuchet MS'>
										<u>Apply lotion or ointments:</u>
									</td>
								</tr>
								<tr>
									<td colspan='4' align='left'><font size='1' face='trebuchet MS'><u>44 Meal Preparation</u></font></td>
								</tr>
								<tr>
									<td align='center'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<input type='checkbox' value='1' <%=meal%> name='meal'>
									</td>
									<td><font size='1' face='trebuchet MS'>
										<u>Needs assistance with meal preparation:</u>
									</td>
									<td align='center'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<input type='checkbox' value='1' <%=commeal%> name='commeal'>
									</td>
									<td><font size='1' face='trebuchet MS'>
										<u>Complete full meal preparation needed:</u>
									</td>
								</tr>
								<tr>
									<td colspan='4' align='left'><font size='1' face='trebuchet MS'><u>45 Medication</u></font></td>
								</tr>
								<tr>
									<td align='center'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<input type='checkbox' value='1' <%=medication%> name='medication'>
									</td>
									<td><font size='1' face='trebuchet MS'>
										<u>Needs medication reminders:</u>
									</td>
								</tr>
								<tr>
									<td colspan='4' align='left'><font size='1' face='trebuchet MS'><u>46 Housekeeping / Cleaning</u></font></td>
								</tr>
								<tr>
									<td align='center'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<input type='checkbox' value='1' <%=housekeep%> name='housekeep'>
									</td>
									<td><font size='1' face='trebuchet MS'>
										<u>Needs assistance with housekeeping cleaning:</u>
									</td>
								</tr>
								<tr>
									<td colspan='4' align='left'><font size='1' face='trebuchet MS'><u>47 Laundry</u></font></td>
								</tr>
								<tr>
									<td align='center'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<input type='checkbox' value='1' <%=laundry%> name='laundry'>
									</td>
									<td><font size='1' face='trebuchet MS'>
										<u>Needs assistance with laundry:</u>
									</td>
								</tr>
								<tr>
									<td colspan='4' align='left'><font size='1' face='trebuchet MS'><u>48 Range of Motion</u></font></td>
								</tr>
								<tr>
									<td align='center'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<input type='checkbox' value='1' <%=ptexer%> name='ptexer'>
									</td>
									<td><font size='1' face='trebuchet MS'>
										<u>Assistance with physical therapy exercises:</u>
									</td>
								</tr>
								<tr>
									<td colspan='4' align='left'><font size='1' face='trebuchet MS'><u>51 Accompaniment</u></font></td>
								</tr>
								<tr>
									<td align='center'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<input type='checkbox' value='1' <%=medical%> name='medical'>
									</td>
									<td><font size='1' face='trebuchet MS'>
										<u>Accompaniment to medical, dental, prescriptions needed:</u>
									</td>
								</tr>
								<tr>
									<td colspan='4' align='left'><font size='1' face='trebuchet MS'><u>52 Accompaniment (non-medical)</u></font></td>
								</tr>
								<tr>
									<td align='center'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<input type='checkbox' value='1' <%=grocery%> name='grocery'>
									</td>
									<td><font size='1' face='trebuchet MS'>
										<u>Needs assistance with grocery shopping or essential community services:</u>
									</td>
								</tr>
								<tr>
									<td colspan='4' align='left'><font size='1' face='trebuchet MS'><u>53 Assistance with Eating</u></font></td>
								</tr>
								<tr>
									<td align='center'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<input type='checkbox' value='1' <%=eat%> name='eat'>
									</td>
									<td><font size='1' face='trebuchet MS'>
										<u>Needs assistance with eating:</u>
									</td>
								</tr>
								<tr>
									<td colspan='4' align='left'><font size='1' face='trebuchet MS'><u>Advanced Skills</u></font></td>
								</tr>
								<tr>
									<td align='center'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<input type='checkbox' value='1' <%=alz%> name='alz'>
									</td>
									<td><font size='1' face='trebuchet MS'>
										<u>Experience with Alzheimer's and/or Dementia needed:</u>
									</td>
										
									<td align='center'>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<input type='checkbox' value='1' <%=Hospice%> name='Hospice'>
									</td>
									<td><font size='1' face='trebuchet MS'>
										<u>Hospice level patient:</u>
									</td>
								</tr>
							</table>
						</td>
					</tr>
				</table>
				<br>
				<table border='0'>
					<tr>
						<td align='center' colspan='2'>
							<input type='button' value='Save' class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" style='width: 110px;' onclick='SaveMe();'>
							<input type='hidden' name='ctr' value='<%=ctr%>'>
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
<%Session("MSG") = ""%>