<%language=vbscript%>
<!-- #include file="_Files.asp" -->

<!-- #include file="_Utils.asp" -->
<!-- #include file="_security.asp" -->
<%
	If Request("WID") <> "" Then
		Driver = False
		Set rsDrive = Server.CreateObject("ADODB.RecordSet")
		rsDrive.Open "SELECT Driver FROM worker_T Where Social_Security_number = '" & Request("WID") & "' ", g_strCONN, 3, 1
		If Not rsDrive.EOF Then
			If rsDrive("Driver") Then Driver = True
		End If
		rsDrive.Close
		Set rsDrive = Nothing
		Set tblFiles = Server.CreateObject("ADODB.Recordset")
		sqlFiles = "SELECT * FROM w_skills_T WHERE wid = '" & Request("WID") & "' "
		tblFiles.Open sqlFiles, g_strCONN, 1, 3
			If Not tblFiles.EOF Then
				If tblFiles("housekeep") Then housekeep = "checked"
				If tblFiles("laundry") Then laundry = "checked"
				If tblFiles("meal") Then meal = "checked"
				If tblFiles("grocery") Then grocery = "checked"
				If tblFiles("dress") Then dress = "checked"
				If tblFiles("eat") Then eat = "checked"
				If tblFiles("asstwalk") Then asstwalk = "checked"
				If tblFiles("asstwheel") Then asstwheel = "checked"
				If tblFiles("asstmotor") Then asstmotor = "checked"
				If tblFiles("commeal") Then commeal = "checked"
				If tblFiles("medical") Then medical = "checked"
				If tblFiles("shower") Then shower = "checked"
				If tblFiles("tub") Then tub = "checked"
				If tblFiles("oral") Then oral = "checked"
				If tblFiles("commode") Then commode = "checked"
				If tblFiles("sit") Then sit = "checked"
				If tblFiles("medication") Then medication = "checked"
				If tblFiles("undress") Then undress = "checked"
				If tblFiles("shampoosink") Then shampoosink = "checked"
				If tblFiles("oralcare") Then oralcare = "checked"
				If tblFiles("massage") Then massage = "checked"
				If tblFiles("shampoobed") Then shampoobed = "checked"
				If tblFiles("shave") Then shave = "checked"
				If tblFiles("bedbath") Then bedbath = "checked"
				If tblFiles("bedpan") Then bedpan = "checked"
				If tblFiles("ptexer") Then ptexer = "checked"
				If tblFiles("hoyer") Then hoyer = "checked"
				If tblFiles("eye") Then eye = "checked"
				If tblFiles("transferbelt") Then transferbelt = "checked"
				If tblFiles("alz") Then alz = "checked"
				If tblFiles("Incontinence") Then Incontinence = "checked"
				If tblFiles("Hospice") Then Hospice = "checked"
				If tblFiles("lotion") Then lotion = "checked"
			Else
				tblFiles.AddNew
				tblFiles("wid") = Request("WID")
				tblFiles.Update
			End If
		tblFiles.Close
		Set tblFiles = Nothing
		Driverme = ""
		If Not Driver Then Driverme = "disabled"
		
	Else
		Session("MSG") = "Sorry. To maintain integrity of database, please Sign-in again."
		Response.Redirect "default.asp"
	End If
%>
<html>
	<head>
		<title>LSS - In-Home Care - PCSP Worker Details - Skills</title>
		<script language='JavaScript'>
			function WDF_Edit()
			{
				document.frmWorDetFil.action = "A_W_Action.asp?page=4";
				document.frmWorDetFil.submit();
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
	<body bgcolor='white' LEFTMARGIN='0' TOPMARGIN='0' onload=''>
		<form method='post' name='frmWorDetFil'>
			<!-- #include file="_boxup.asp" -->
			<!-- #include file="_NavHeader.asp" -->
			<br>
			<center>
				<table border='0'>
					<tr>
						<td colspan='4' align='center' width='500px'>
							<font size='2' face='trebuchet MS'><b><u>PCSP Worker Details - Files</u></b></font>
							<a href='A_Worker.asp?WID=<%=Request("WID")%>' style='text-decoration:none'><font size='1' face='trebuchet MS'>[Details]</font></a>
							<a href='A_w_Files.asp?WID=<%=Request("WID")%>' style='text-decoration:none'><font size='1' face='trebuchet MS'>[Files]</font></a>
							
								<font  size='2' face='trebuchet MS'>[Skills]</font>
			
							<a href="WorkCon.asp?WID=<%=Request("WID")%>" style='text-decoration: none;'>
								<font color='blue' size='1' face='trebuchet MS'>[List]</font>
							</a>
							<a href="A_W_log.asp?WID=<%=Request("WID")%>" style='text-decoration: none;'>
								<font color='blue' size='1' face='trebuchet MS'>[Log]</font>
							</a>
							<a href="A_W_misc.asp?WID=<%=Request("WID")%>" style='text-decoration: none;'>
								<font color='blue' size='1' face='trebuchet MS'>[Violations]</font>
							</a>
							<a href="wimport.asp?WID=<%=Request("WID")%>" style='text-decoration: none;'>
								<font color='blue' size='1' face='trebuchet MS'>[Uploads]</font>
							</a>
						</td>
					</tr>
					<tr><td colspan='2' align='center'><font color='red' face='trebuchet MS' size='1'>&nbsp;<%=Session("MSG")%>&nbsp;</font></td></tr>
					<tr>
						<td>
							<font size='1' face='trebuchet MS'><u>Name:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 120px;' readonly name='MCnum' value="<%=Session("Wname")%>"></font>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td align='left'>
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
										<input type='checkbox' value='1' <%=Driverme%> <%=medical%> name='medical'>
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
										<input type='checkbox' value='1' <%=Driverme%> <%=grocery%> name='grocery'>
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
					<tr><td>&nbsp;</td></tr>
				</table>
				<table border='0'>
					<tr>
						<td align='center' colspan='2'>
							<input type='button' value='Save' style='width: 110px;' class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='JavaScript:WDF_Edit();'>
							<input type='hidden' name='ctr' value='<%=ctr%>'>
							<input type='hidden' name='WID' value="<%=Request("WID")%>">
						</td>
					</tr>
					
				</table>
			</center>	
			<!-- #include file="_boxdown.asp" -->
		</form>
	</body>
</html>
<%Session("MSG") = "" %>