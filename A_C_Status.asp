<%Language=VBScript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_security.asp" -->
<%
	If Request("MNum") <> "" Then
		Set tblStatus = Server.CreateObject("ADODB.Recordset")
		sqlStatus = "SELECT * FROM C_Status_t WHERE Medicaid_Number = '" & Request("MNum") & "' "
		tblStatus.Open sqlStatus, g_strCONN, 1, 3
			If Not tblStatus.EOF Then
				'THrs = tblStatus("Hours_Approved")
				'Active?
				'If tblStatus("Active") = True Then Act = "checked"
				Act = ""
				Nurs = ""
				Direct = ""
				Work = ""
				Death = ""
				'If tblStatus("Status") = "Inactive" Then Act = "Checked"
				If tblStatus("Active") = False Then Act = "checked"
				IDate = tblStatus("Inactive_Date")
				If tblStatus("Enter_Nursing_Home") = True Then Nurs = "checked"
				If tblStatus("Unable_Self_Direct") = True Then Direct = "checked"
				If tblStatus("Unable_Suitable_Worker") = True Then Work = "checked"
				If tblStatus("Death") = True Then Death = "checked"
				A_Other = tblStatus("A_Other")
				'On Hold?
				'If tblStatus("On_Hold") = True Then Hold = "checked"
				'Hold = ""
				'Hosp = ""
				'NWork = ""
				'If tblStatus("onHold") = True Then Hold = "checked"
				'If tblStatus("In_Hospital") = True Then Hosp = "checked"
				'If tblStatus("New_Worker") = True Then NWork = "checked"
				'H_Other = tblStatus("H_Other")
				'ToDate = tblStatus("H_To_Date")
				'FromDate = tblStatus("H_From_Date")
			Else
				tblStatus.AddNew
				tblStatus("Medicaid_Number") = Request("MNum")
				tblStatus.Update
			End If
		tblStatus.Close
		Set tblStatus = Nothing	
			
		'on hold new
		'Set rsHold = Server.CreateObject("ADODB.RecordSet")
		'rsHold.Open "SELECT TOP 1 * FROM C_OnHold_T WHERE Cid = '" & Request("MNum") & "' ORDER BY [datestamp] DESC", g_strCONN, 3, 1
		'If Not rsHold.EOF Then
		'	Hold = ""
		'	Hosp = ""
		'	NWork = ""
		'	If rsHold("on_Hold") = True Then Hold = "checked"
		'	If rsHold("In_Hospital") = True Then Hosp = "checked"
		'	If rsHold("New_Worker") = True Then NWork = "checked"
		'	H_Other = rsHold("H_Other")
		'	ToDate = rsHold("H_To_Date")
		'	FromDate = rsHold("H_From_Date")
		'End If
		'rsHold.Close
		'Set rsHold = Nothing
		ScriptName = Request.ServerVariables("SCRIPT_NAME")
		NumPerPage = 5
		If Request.QueryString("LogPage") = "" Then
			CurrPage = 1
		Else
			CurrPage = CInt(Request.QueryString("LogPage"))
		End if
		Set rsLog = Server.CreateObject("ADODB.RecordSet")
		sqlLog = "SELECT * FROM [C_OnHold_T] WHERE [Cid] = '" & Request("MNum") & "' ORDER BY [datestamp] DESC"
		rsLog.Open sqlLog, g_strCONN, 1, 3
		TotalPages = 1
		If Not rsLog.EOF Then
			rsLog.MoveFirst
			rsLog.PageSize = NumPerPage
			TotalPages = rsLog.PageCount
			rsLog.AbsolutePage = CurrPage
			
			Hold = ""
			Hosp = ""
			NWork = ""
			If rsLog("on_Hold") = True Then Hold = "checked"
			If rsLog("In_Hospital") = True Then Hosp = "checked"
			If rsLog("New_Worker") = True Then NWork = "checked"
			H_Other = rsLog("H_Other")
			ToDate = rsLog("H_To_Date")
			FromDate = rsLog("H_From_Date")
			rslog.movenext
		End If
		
		ctr = 0
		reason = ""
		Holdlog = ""
		Do While Not rsLog.EOF And ctr < rsLog.PageSize
			'If rsLog("train") <> "" Then
				if Z_IsOdd(ctr) = false then 
						kulay = "#FFFAF0" 
					else 
						kulay = "#FFFFFF"
					end if
				dtestamp = rsLog("datestamp")
				If rsLog("on_Hold") = True Then Holdlog = "True"
				If rsLog("In_Hospital") = True Then reason = "In Hospital/Rehab/Nursing Home "
				If rsLog("New_Worker") = True Then reason = reason & " Needs New Worker "
				If rsLog("H_Other") <> "" Then reason = reason & rsLog("H_Other")
				strLog = strLog & "<tr bgcolor='" & kulay & "'>" & _
					"<td align='center'><font size='1' face='trebuchet MS'>" & dtestamp & _
					"</font></td><td align='center'><font size='1' face='trebuchet MS'>" & Holdlog & _
					"</font></td>" & _
					"<td align='center'><font size='1' face='trebuchet MS'>" & rsLog("h_from_date") & _
					"</font></td><td align='center'><font size='1' face='trebuchet MS'>" & rsLog("h_to_date") & _
					"</font></td><td align='center'><font size='1' face='trebuchet MS'>" & reason & _
					"</font></td>" & _
					"</tr>"			
	
				'rsLog.MoveNext
				ctr = ctr + 1
			'End If
			rsLog.MoveNext
			
		Loop
		rsLog.Close
		'''''''hosp
		NumPerPage4 = 5
		
		If Request.QueryString("LogPage4") = "" Then
			CurrPage4 = 1
		Else
			CurrPage4 = CInt(Request.QueryString("LogPage4"))
		End if
		Set rsLog4 = Server.CreateObject("ADODB.RecordSet")
		sqlLog4 = "SELECT * FROM [C_Site_Visit_Dates_t] WHERE [Medicaid_Number] = '" & Request("Mnum") & "' AND not [hospdate] is null ORDER BY [hospdate] DESC"
		rsLog4.Open sqlLog4, g_strCONN, 1, 3
		TotalPages4 = 1
		If Not rsLog4.EOF Then
			rsLog4.MoveFirst
			rsLog4.PageSize = NumPerPage4
			TotalPages4 = rsLog4.PageCount
			rsLog4.AbsolutePage = CurrPage4
		Else
			strLog4 = "<tr><td colspan='4' align='center'><font size='1'>N/A</font></td></tr>"
		End If
		ctr4 = 0
		Do While Not rsLog4.EOF AND ctr4 < rsLog4.PageSize 
				If rsLog4("hospdate") <> "" Then
					if Z_IsOdd(ctr4) = true then 
							kulay = "#FFFAF0" 
						else 
							kulay = "#FFFFFF"
						end if
					strLog4 = strLog4 & "<tr bgcolor='" & kulay & "'><td align='center'><input type='checkbox' name='chkhosp" & ctr4 & _
						"' value='" & rsLog4("index") & "'></td><td align='center'><input type='text' " & _
						"style='font-size: 10px; height: 20px;' size='9'  maxlength='10' name='txthospdate" & ctr4 & "' value='" & _
						rsLog4("hospdate") & "'></td><td colspan='2'><textarea cols='30' name='hospcom" & ctr4 & "'>" & rsLog4("hospcom") & "</textarea></td></tr>"
					'rsLog2.MoveNext
					ctr4 = ctr4+ 1
				End If
			rsLog4.MoveNext
		Loop
		rsLog4.Close
		Set rsLog4 = Nothing
		'''''''sup
		NumPerPage5 = 5
		
		If Request.QueryString("LogPage5") = "" Then
			CurrPage5 = 1
		Else
			CurrPage5 = CInt(Request.QueryString("LogPage5"))
		End if
		Set rsLog5 = Server.CreateObject("ADODB.RecordSet")
		sqlLog5 = "SELECT * FROM [C_Site_Visit_Dates_t] WHERE [Medicaid_Number] = '" & Request("Mnum") & "' AND not [supdate] is null ORDER BY [supdate] DESC"
		rsLog5.Open sqlLog5, g_strCONN, 1, 3
		TotalPages5 = 1
		If Not rsLog5.EOF Then
			rsLog5.MoveFirst
			rsLog5.PageSize = NumPerPage5
			TotalPages5 = rsLog5.PageCount
			rsLog5.AbsolutePage = CurrPage5
		Else
			strLog5 = "<tr><td colspan='4' align='center'><font size='1'>N/A</font></td></tr>"
		End If
		ctr5 = 0
		Do While Not rsLog5.EOF AND ctr5 < rsLog5.PageSize 
				If rsLog5("supdate") <> "" Then
					if Z_IsOdd(ctr5) = true then 
							kulay = "#FFFAF0" 
						else 
							kulay = "#FFFFFF"
						end if
					strLog5 = strLog5 & "<tr bgcolor='" & kulay & "'><td align='center'><input type='checkbox' name='chksup" & ctr5 & _
						"' value='" & rsLog5("index") & "'></td><td align='center'><input type='text' " & _
						"style='font-size: 10px; height: 20px;' size='9'  maxlength='10' name='txtsupdate" & ctr5 & "' value='" & _
						rsLog5("supdate") & "'></td><td colspan='2'><textarea cols='30' name='supcom" & ctr5 & "'>" & rsLog5("supnotes") & "</textarea></td></tr>"
					'rsLog2.MoveNext
					ctr5 = ctr5 + 1
				End If
			rsLog5.MoveNext
		Loop
		rsLog5.Close
		Set rsLog5 = Nothing
	Else
		Session("MSG") = "Sorry. To maintain integrity of database, please Sign-in again."
		Response.Redirect "default.asp"
	End If
%>
<html>
	<head>
		<title>LSS - In-Home Care - Consumer Details - Status</title>
		<script language='JavaScript'>
			function active()
			{
				if (document.frmConDetStat.chkActive.checked == true)
					{document.frmConDetStat.Inactive.disabled = false
					document.frmConDetStat.chkDeath.disabled = false
					document.frmConDetStat.chkWork.disabled = false
					document.frmConDetStat.chkNurse.disabled = false	
					document.frmConDetStat.chkDir.disabled = false
					document.frmConDetStat.Others.disabled = false}
				else
					{document.frmConDetStat.Inactive.disabled = true
					document.frmConDetStat.chkDeath.disabled = true
					document.frmConDetStat.chkWork.disabled = true
					document.frmConDetStat.chkNurse.disabled = true	
					document.frmConDetStat.chkDir.disabled = true
					document.frmConDetStat.Others.disabled = true};
			}
			function onhold()
			{
				if (document.frmConDetStat.chkOnHold.checked == true)
					{document.frmConDetStat.chkHos.disabled = false
					document.frmConDetStat.chkNew.disabled = false
					document.frmConDetStat.frmDate.disabled = false
					document.frmConDetStat.toDate.disabled = false
					document.frmConDetStat.Ot.disabled = false}
				else
					{document.frmConDetStat.chkHos.disabled = true
					document.frmConDetStat.chkNew.disabled = true
					document.frmConDetStat.frmDate.disabled = true
					document.frmConDetStat.toDate.disabled = true
					document.frmConDetStat.Ot.disabled = true
					document.frmConDetStat.chkHos.checked = false
					document.frmConDetStat.chkNew.checked = false
					document.frmConDetStat.frmDate.value = ''
					document.frmConDetStat.toDate.value = ''
					document.frmConDetStat.Ot.value = ''
					};
			}
			function CDS_Edit()
			{
				if (document.frmConDetStat.chkActive.checked == true && document.frmConDetStat.Inactive.value == "")
				{
					alert("ERROR: Inactive Date is required.")
					return;
				}
				document.frmConDetStat.action = "A_C_Action.asp?page=2";
				document.frmConDetStat.submit();
			}
			function CD_Del_Towns()
			{
				document.frmConDetStat.action = "A_C_Del.asp?req=1";
				document.frmConDetStat.submit();
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
			function DelLog()
			{
				document.frmConDetStat.action = "A_C_Del.asp?req=4";
				document.frmConDetStat.submit();
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
	<body bgcolor='white' onload='JavaScript:active();onhold()' LEFTMARGIN='0' TOPMARGIN='0'>
		<form method='post' name='frmConDetStat'>
			<!-- #include file="_boxup.asp" -->
			<!-- #include file="_NavHeader.asp" -->
			<br>
			<center>
				<table border='0' width='500px'>
					<tr>
						<td colspan='4' align='center' >
							<font size='2' face='trebuchet MS'><b><u>Consumer Details - Status</u></b></font>
							<a href='A_Consumer.asp?MNum=<%=Request("MNum")%>' style='text-decoration:none'><font size='1' face='trebuchet MS'>[Details]</font></a>
							<font size='2' face='trebuchet MS'>[Status]</a>
							<a href='A_C_health.asp?MNum=<%=Request("MNum")%>' style='text-decoration:none'><font size='1' face='trebuchet MS'>[Health]</font></a>
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
					
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td align='left'>
							<font size='1' face='trebuchet MS'><u>Inactive:</u>&nbsp;
							<input type='checkbox' name='chkActive' <%=Act%> value='1' onclick='Javascript:active();'>
						</td>
					</tr>
					<tr>
						<td align='left'>
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							<font size='1' face='trebuchet MS'><u>Inactive Date:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 80px;' name='Inactive' maxlength='10' value="<%=IDate%>" onKeyUp="javascript:return maskMe(this.value,this,'2,5','/');" onBlur="javascript:return maskMe(this.value,this,'2,5','/');"></font>
							<font size='1' face='trebuchet MS'>(mm,dd,yyyy)</font>
						</td>
					</tr>
					<tr>
						<td align='left'>
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							<input type='checkbox' name='chkNurse' disabled <%=Nurs%>>
							<font size='1' face='trebuchet MS'><u>Enter Nursing Home or Other Setting:</u>&nbsp;
						&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							<input type='checkbox' name='chkDir' disabled <%=Direct%>>
							<font size='1' face='trebuchet MS'><u>Unable to Self-Direct:</u>&nbsp;
						</td>
					</tr>
					<tr>
						<td align='left'>
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							<input type='checkbox' name='chkWork' disabled <%=Work%>>
							<font size='1' face='trebuchet MS'><u>Unable to Find Suitable Worker:</u>&nbsp;
						&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
						&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							<input type='checkbox' name='chkDeath' disabled <%=Death%>>
							<font size='1' face='trebuchet MS'><u>Death:</u>&nbsp;
						</td>
					</tr>
					<tr>
						<td align='left'>
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							<font size='1' face='trebuchet MS'><u>Other:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 80px;' maxlength='49' name='Others' disabled value="<%=A_Other%>"></font>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td align='left'>
							<font size='1' face='trebuchet MS'><u>Services Temporarily on Hold:</u>&nbsp;
							<input type='checkbox' name='chkOnHold' <%=Hold%> value='1' onclick='Javascript:onhold();'>
						</td>
					</tr>
					<tr>
						<td align='left'>
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							<font size='1' face='trebuchet MS'><u>From:</u>&nbsp;
							<input type='text' disabled maxlength='10' style='font-size: 10px; height: 20px; width: 80px;' name='frmDate' value="<%=FromDate%>" onKeyUp="javascript:return maskMe(this.value,this,'2,5','/');" onBlur="javascript:return maskMe(this.value,this,'2,5','/');">
							<font size='1' face='trebuchet MS'><u>To:</u>&nbsp;
							<input type='text' disabled maxlength='10' style='font-size: 10px; height: 20px; width: 80px;' name='toDate' value="<%=ToDate%>" onKeyUp="javascript:return maskMe(this.value,this,'2,5','/');" onBlur="javascript:return maskMe(this.value,this,'2,5','/');"> 
							<font size='1' face='trebuchet MS'>(mm,dd,yyyy)</font>
						</td>
					</tr>
					<tr>
						<td align='left'>
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							<input type='checkbox' name='chkHos' disabled <%=Hosp%> value=1>
							<font size='1' face='trebuchet MS'><u>In Hospital/Rehab/Nursing Home</u>&nbsp;
						&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
						&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							<input type='checkbox' name='chkNew' disabled <%=NWork%> value=1>
							<font size='1' face='trebuchet MS'><u>Needs New Worker:</u>&nbsp;
						</td>
					</tr>
					<tr>
						<td align='left'>
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							<font size='1' face='trebuchet MS'><u>Other:</u>&nbsp;
							<input type='text' disabled style='font-size: 10px; height: 20px; width: 200px;' name='Ot' maxlength='49' value="<%=H_Other%>"></font>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td valign='top' colspan='2' align='center'>
						<table border='1' width='50%'>
							<tr bgcolor='#040C8B'><td colspan='6' align='center' style="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=0, startColorstr=#040C8B, endColorstr=#c7c9e5);">
								<font size='2' face='trebuchet MS' color='white'>On Hold Log</font>
								</td></tr>
							<tr><td colspan='4'  align='center'>
								<%  'Site
									If Not CurrPage = 1 Then 
										Response.Write "<a href='" & ScriptName & "?LogPage=" & CurrPage - 1 &   "'><font size='1' face='trebuchet MS'>Prev</a> | "
										Session("page") = CurrPage 
									Else
										Response.Write "<font size='1' face='trebuchet MS'>Prev | "
									End If
									If Not CurrPage = TotalPages Then
										Response.Write "<a href='" & ScriptName & "?LogPage=" & CurrPage + 1 &  "'>Next</font></a>"
										Session("page") = CurrPage 
									Else
										Response.Write "Next</font>"
									End If
									
								%>
								</td>
							
								<td align='right'><font size='1' face='trebuchet MS'>On Hold <%=CurrPage%> of <%=TotalPages%></font></td>
							</tr>
							<tr>
								<td align='center'><font size='1' face='trebuchet MS'>Datestamp</font></td>
								<td align='center'><font size='1' face='trebuchet MS'>On Hold</font></td>
								<td align='center'><font size='1' face='trebuchet MS'>From</font></td>
								<td align='center'><font size='1' face='trebuchet MS'>To</font></td>
								<td align='center'><font size='1' face='trebuchet MS'>Reason</font></td>
		
								</tr>
							<%=strLog%>
						</table>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td valign='top'>
						<table border='1'>
							<tr bgcolor='#040C8B'>
								<td colspan='5' align='center' style="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=0, startColorstr=#040C8B, endColorstr=#c7c9e5);">
									<font size='2' face='trebuchet MS' color='white'>Hospitalizations <a href='JavaScript: DelLog();' style='text-decoration: none'><font size='1' face='trebuchet ms' color='white'>[Delete]</font></a></font>
									</td></tr>
							<tr><td colspan='2' width='100%' align='center'>
								<%	'misc
									If Not CurrPage5 = 1 Then 
										Response.Write "<a href='" & ScriptName & "?LogPage=" & CurrPage & "&LogPage4=" & CurrPage4 - 1 & "&Mnum=" & Request("Mnum") & "'><font size='1' face='trebuchet MS'>Prev</a> | "
										Session("page") = CurrPage5
									Else
										Response.Write "<font size='1' face='trebuchet MS'>Prev | "
									End If
									If Not CurrPage5 = TotalPages5 Then
										Response.Write "<a href='" & ScriptName & "?LogPage=" & CurrPage & "&LogPage4=" & CurrPage4 + 1 & "&Mnum=" & Request("Mnum") & "'>Next</font></a>"
										Session("page") = CurrPage4
									Else
										Response.Write "Next</font>"
									End If
									
								%>
								</td>
								<td colspan='3' align='right'><font size='1' face='trebuchet MS'>Hospitalizations <%=CurrPage4%> of <%=TotalPages4%></font></td>
							</tr>
							
							<tr><td colspan='2' align='center'><font size='1' face='trebuchet MS'>Date</font></td>
									<td colspan='2' align='center'><font size='1' face='trebuchet MS'>Comment</font></td>
								</tr>
							<%=strLog4%>
							<tr bgcolor='#040C8B'>
								<td align='center' colspan='7' style="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=0, startColorstr=#040C8B, endColorstr=#c7c9e5);">
									<font size='2' face='trebuchet ms' color='white'>New Hospitalizations Entries</font></td></tr>
							<tr>
								<tr>
									<td  width='25px'>&nbsp;</td>
									<td valign='top' align='center'>
										<input style='font-size: 10px; height: 20px;' size='9' maxlength='10' name='txtHospdate'>
									</td>
									<td>
										<textarea cols='30' name='hospcom'></textarea>
									</td>
								</tr>
						</table>
						</td>
						<td valign='top'>
						<table border='1'>
							<tr bgcolor='#040C8B'>
								<td colspan='5' align='center' style="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=0, startColorstr=#040C8B, endColorstr=#c7c9e5);">
									<font size='2' face='trebuchet MS' color='white'>Supervisory Notes <a href='JavaScript: DelLog();' style='text-decoration: none'><font size='1' face='trebuchet ms' color='white'>[Delete]</font></a></font>
									</td></tr>
							<tr><td colspan='2' width='100%' align='center'>
								<%	'misc
									If Not CurrPage4 = 1 Then 
										Response.Write "<a href='" & ScriptName & "?LogPage=" & CurrPage & "&LogPage4=" & CurrPage4 - 1 & "&Mnum=" & Request("Mnum") & "'><font size='1' face='trebuchet MS'>Prev</a> | "
										Session("page") = CurrPage4
									Else
										Response.Write "<font size='1' face='trebuchet MS'>Prev | "
									End If
									If Not CurrPage4 = TotalPages4 Then
										Response.Write "<a href='" & ScriptName & "?LogPage=" & CurrPage & "&LogPage4=" & CurrPage4 + 1 & "&Mnum=" & Request("Mnum") & "'>Next</font></a>"
										Session("page") = CurrPage4
									Else
										Response.Write "Next</font>"
									End If
									
								%>
								</td>
								<td colspan='3' align='right'><font size='1' face='trebuchet MS'>Supervisory Notes <%=CurrPage4%> of <%=TotalPages4%></font></td>
							</tr>
							
							<tr><td colspan='2' align='center'><font size='1' face='trebuchet MS'>Date</font></td>
									<td colspan='2' align='center'><font size='1' face='trebuchet MS'>Comment</font></td>
								</tr>
							<%=strLog5%>
							<tr bgcolor='#040C8B'>
								<td align='center' colspan='7' style="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=0, startColorstr=#040C8B, endColorstr=#c7c9e5);">
									<font size='2' face='trebuchet ms' color='white'>New Supervisory Notes Entries</font></td></tr>
							<tr>
								<tr>
									<td  width='25px'>&nbsp;</td>
									<td valign='top' align='center'>
										<input style='font-size: 10px; height: 20px;' size='9' maxlength='10' name='txtsupdate'>
									</td>
									<td>
										<textarea cols='30' name='supcom'></textarea>
									</td>
								</tr>
						</table>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
				</td>
					<tr>
						<td align='center' colspan='2'>
							<input type='button' value='Save' style='width: 110px;' class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='JavaScript:CDS_Edit();'>
							<input type='hidden' name='Mnum' value="<%=Request("MNum")%>">
							<input type='hidden' name='ctr4' value='<%=ctr4%>'>
							<input type='hidden' name='ctr' value='<%=ctr%>'>
							<input type='hidden' name='ctr5' value='<%=ctr5%>'>
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