<%language=vbscript%>
<!-- #include file="_Files.asp" -->

<!-- #include file="_Utils.asp" -->
<!-- #include file="_security.asp" -->
<%
	If Request("Mnum") <> "" Then
		ScriptName = Request.ServerVariables("SCRIPT_NAME")
		''''''Site
		NumPerPage = 10
		
		If Request.QueryString("LogPage") = "" Then
			CurrPage = 1
		Else
			CurrPage = CInt(Request.QueryString("LogPage"))
		End if
		Set rsLog = Server.CreateObject("ADODB.RecordSet")
		sqlLog = "SELECT * FROM [C_Site_Visit_Dates_t] WHERE [Medicaid_Number] = '" & Request("Mnum") & "' ORDER BY [Site_V_Date] DESC"
		'response.write sqllog
		rsLog.Open sqlLog, g_strCONN, 1, 3
		TotalPages = 10
		If Not rsLog.EOF Then
			rsLog.MoveFirst
			rsLog.PageSize = NumPerPage
			TotalPages = rsLog.PageCount
			rsLog.AbsolutePage = CurrPage
		Else
			strLog = "<tr><td colspan='4' align='center'><font size='1'>N/A</font></td></tr>"
		End If
		ctr = 0
		Do While Not rsLog.EOF And ctr < rsLog.PageSize
			If rsLog("Site_V_Date") <> "" Then
				if Z_IsOdd(ctr) = true then 
						kulay = "#FFFAF0" 
					else 
						kulay = "#FFFFFF"
					end if
				strLog = strLog & "<tr bgcolor='" & kulay & "'><td align='center'><input type='checkbox' name='chkSV" & ctr & _
					"' value='" & rsLog("index") & "'></td><td align='center'><input type='text' " & _
					"style='font-size: 10px; height: 20px;' size='9' maxlength='10' name='txtSVD" & ctr & "' value='" & _
					rsLog("Site_V_Date") & "'></td><td colspan='2'><textarea cols='18' name='Vcom" & ctr & "'>" & rsLog("Comments") & "</textarea></td></tr>"
				'rsLog.MoveNext
				ctr = ctr + 1
			End If
			rsLog.MoveNext
			
		Loop
		rsLog.Close
		
		Set rsLog = Nothing
		'''''''phone
		NumPerPage2 = 10
		
		If Request.QueryString("LogPage2") = "" Then
			CurrPage2 = 1
		Else
			CurrPage2 = CInt(Request.QueryString("LogPage2"))
		End if
		Set rsLog2 = Server.CreateObject("ADODB.RecordSet")
		sqlLog2 = "SELECT * FROM [C_Site_Visit_Dates_t] WHERE [Medicaid_Number] = '" & Request("Mnum") & "' ORDER BY [PhoneCon_last] DESC"
		response.write "<!--SQL2: " & sqlLog2 & "-->"
		rsLog2.Open sqlLog2, g_strCONN, 1, 3
		TotalPages2 = 1
		If Not rsLog2.EOF Then
			rsLog2.MoveFirst
			rsLog2.PageSize = NumPerPage2
			TotalPages2 = rsLog2.PageCount
			rsLog2.AbsolutePage = CurrPage2
		Else
			strLog2 = "<tr><td colspan='4' align='center'><font size='1'>N/A</font></td></tr>"
		End If
		ctr2 = 0
		Do While Not rsLog2.EOF AND ctr2 < rsLog2.PageSize 
				If rsLog2("PhoneCon_last") <> "" Then
					if Z_IsOdd(ctr2) = true then 
							kulay = "#FFFAF0" 
						else 
							kulay = "#FFFFFF"
						end if
					strLog2 = strLog2 & "<tr bgcolor='" & kulay & "'><td align='center'><input type='checkbox' name='chkPC" & ctr2 & _
						"' value='" & rsLog2("index") & "'></td><td align='center'><input type='text' " & _
						"style='font-size: 10px; height: 20px;' size='9'  maxlength='10' name='txtPCD" & ctr2 & "' value='" & _
						rsLog2("PhoneCon_last") & "'></td><td colspan='2'><textarea cols='18' name='Pcom" & ctr2 & "'>" & rsLog2("PCom") & "</textarea></td></tr>"
					'rsLog2.MoveNext
					ctr2 = ctr2 + 1
				End If
			rsLog2.MoveNext
		Loop
		rsLog2.Close
		Set rsLog2 = Nothing
		'''''''misc
		NumPerPage3 = 10
		
		If Request.QueryString("LogPage3") = "" Then
			CurrPage3 = 1
		Else
			CurrPage3 = CInt(Request.QueryString("LogPage3"))
		End if
		Set rsLog3 = Server.CreateObject("ADODB.RecordSet")
		sqlLog3 = "SELECT * FROM [C_Site_Visit_Dates_t] WHERE [Medicaid_Number] = '" & Request("Mnum") & "' ORDER BY [miscCon] DESC"
		response.write "<!--SQL3: " & sqlLog3 & "-->"
		rsLog3.Open sqlLog3, g_strCONN, 1, 3
		TotalPages3 = 1
		If Not rsLog3.EOF Then
			rsLog3.MoveFirst
			rsLog3.PageSize = NumPerPage3
			TotalPages3 = rsLog3.PageCount
			rsLog3.AbsolutePage = CurrPage3
		Else
			strLog3 = "<tr><td colspan='4' align='center'><font size='1'>N/A</font></td></tr>"
		End If
		ctr3 = 0
		Do While Not rsLog3.EOF AND ctr3 < rsLog3.PageSize 
				If rsLog3("miscCon") <> "" Then
					if Z_IsOdd(ctr3) = true then 
							kulay = "#FFFAF0" 
						else 
							kulay = "#FFFFFF"
						end if
					strLog3 = strLog3 & "<tr bgcolor='" & kulay & "'><td align='center'><input type='checkbox' name='chkMC" & ctr3 & _
						"' value='" & rsLog3("index") & "'></td><td align='center'><input type='text' " & _
						"style='font-size: 10px; height: 20px;' size='9'  maxlength='10' name='txtMCD" & ctr3 & "' value='" & _
						rsLog3("miscCon") & "'></td><td colspan='2'><textarea cols='18' name='MCon" & ctr3 & "'>" & rsLog3("Misc") & "</textarea></td></tr>"
					'rsLog2.MoveNext
					ctr3 = ctr3+ 1
				End If
			rsLog3.MoveNext
		Loop
		rsLog3.Close
		Set rsLog3 = Nothing
		'''''''hosp
		'NumPerPage4 = 10
		'
		'If Request.QueryString("LogPage4") = "" Then
		'	CurrPage4 = 1
		'Else
		'	CurrPage4 = CInt(Request.QueryString("LogPage4"))
		'End if
		'Set rsLog4 = Server.CreateObject("ADODB.RecordSet")
		'sqlLog4 = "SELECT * FROM [C_Site_Visit_Dates_t] WHERE [Medicaid_Number] = '" & Request("Mnum") & "' ORDER BY [hospdate] DESC"
		'response.write "<!--SQL3: " & sqlLog3 & "-->"
		'rsLog4.Open sqlLog4, g_strCONN, 1, 3
		'TotalPages4 = 1
		'If Not rsLog4.EOF Then
		'	rsLog4.MoveFirst
		'	rsLog4.PageSize = NumPerPage4
		'	TotalPages4 = rsLog4.PageCount
		'	rsLog4.AbsolutePage = CurrPage4
		'Else
		'	strLog4 = "<tr><td colspan='4' align='center'><font size='1'>N/A</font></td></tr>"
		'End If
		'ctr4 = 0
		'Do While Not rsLog4.EOF AND ctr4 < rsLog4.PageSize 
		'		If rsLog4("hospdate") <> "" Then
		'			if Z_IsOdd(ctr4) = true then 
		'					kulay = "#FFFAF0" 
		'				else 
		'					kulay = "#FFFFFF"
		'				end if
		'			strLog4 = strLog4 & "<tr bgcolor='" & kulay & "'><td align='center'><input type='checkbox' name='chkhosp" & ctr4 & _
		'				"' value='" & rsLog4("index") & "'></td><td align='center'><input type='text' " & _
		'				"style='font-size: 10px; height: 20px;' size='9'  maxlength='10' name='txthospdate" & ctr4 & "' value='" & _
		'				rsLog4("hospdate") & "'></td><td colspan='2'><textarea cols='18' name='hospcom" & ctr4 & "'>" & rsLog4("hospcom") & "</textarea></td></tr>"
		'			'rsLog2.MoveNext
		'			ctr4 = ctr4+ 1
		'		End If
		'	rsLog4.MoveNext
		'Loop
		'rsLog4.Close
		'Set rsLog4 = Nothing
	Else
		Session("MSG") = "Sorry. To maintain integrity of database, please Sign-in again."
		Response.Redirect "default.asp"
	End If
%>
<html>
	<head>
		<title>LSS - In-Home Care - Consumer Details - Log</title>
		<script language='JavaScript'>
			function DelLog()
			{
				document.frmConDetLog.action = "A_C_Del.asp?req=3";
				document.frmConDetLog.submit();
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
		<!-- #include file="_boxup.asp" -->
		<!-- #include file="_NavHeader.asp" -->
		<form method='post' name='frmConDetLog' action='NewLog.asp'>
		
			<br>
		
				<table border='0' align='center'>
					<tr>
						<td colspan='4' align='center' width='500px'>
							<font size='2' face='trebuchet MS'><b><u>Consumer Details - Log</u></b></font>
							<a href='A_Consumer.asp?MNum=<%=Request("MNum")%>' style='text-decoration:none'><font size='1' face='trebuchet MS'>[Details]</font></a>
							<a href='A_C_Status.asp?MNum=<%=Request("MNum")%>' style='text-decoration:none'><font size='1' face='trebuchet MS'>[Status]</font></a>
							<a href='A_C_Health.asp?MNum=<%=Request("MNum")%>' style='text-decoration:none'><font size='1' face='trebuchet MS'>[Health]</font></a>
							<a href='A_C_Files.asp?MNum=<%=Request("MNum")%>' style='text-decoration:none'><font size='1' face='trebuchet MS'>[Files]</font></a>
							<font size='2' face='trebuchet MS'>[Log]</font>
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
				</table>
				<br>
				<div style='OVERFLOW: AUTO; ' align='center'>
				<table cellspacing='0' cellpadding='0' border='0' align='center'>
					<tr>
						<td valign='top'>
						<table border='1'>
							<tr bgcolor='#040C8B'>
								<td colspan='5' align='center' style="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=0, startColorstr=#040C8B, endColorstr=#c7c9e5);">
									<font size='2' face='trebuchet MS' color='white'>Site Visit</font>
									</td></tr>
								<tr><td colspan='2' width='100px' align='center'>
									<%  'Site
										If Not CurrPage = 1 Then 
											Response.Write "<a href='" & ScriptName & "?LogPage=" & CurrPage - 1 &  "&LogPage2=" & CurrPage2 & "&LogPage3=" & CurrPage3 & "&LogPage4=" & CurrPage4 &  "&Mnum=" & Request("Mnum") & "')'><font size='1' face='trebuchet MS'>Prev</a> | "
											Session("page") = CurrPage 
										Else
											Response.Write "<font size='1' face='trebuchet MS'>Prev | "
										End If
										If Not CurrPage = TotalPages Then
											Response.Write "<a href='" & ScriptName & "?LogPage=" & CurrPage + 1 & "&LogPage2=" & CurrPage2 & "&LogPage3=" & CurrPage3 & "&LogPage4=" & CurrPage4 &  "&Mnum=" & Request("Mnum") & "'>Next</font></a>"
											Session("page") = CurrPage 
										Else
											Response.Write "Next</font>"
										End If
										
									%>
								</td>
								<td colspan='3' align='right'><font size='1' face='trebuchet MS'>Site Visit <%=CurrPage%> of <%=TotalPages%></font></td>
							</tr>
							<tr><td colspan='2' align='center'><font size='1' face='trebuchet MS'>Date</font></td>
									<td colspan='2' align='center'><font size='1' face='trebuchet MS'>Comment</font></td>
								</tr>
							<%=strLog%>
						</table>
						</td>
						<td>&nbsp;</td>
						<td valign='top'>
						<table border='1'>
							<tr bgcolor='#040C8B'>
								<td colspan='5' align='center' style="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=0, startColorstr=#040C8B, endColorstr=#c7c9e5);">
									<font size='2' face='trebuchet MS' color='white'>Phone Call</font>
									</td></tr>
							<tr><td colspan='2' width='100px' align='center'>
								<%	'phone
									If Not CurrPage2 = 1 Then 
										Response.Write "<a href='" & ScriptName & "?LogPage=" & CurrPage &  "&LogPage2=" & CurrPage2 - 1 & "&LogPage3=" & CurrPage3  & "&LogPage4=" & CurrPage4 &  "&Mnum=" & Request("Mnum") & "'><font size='1' face='trebuchet MS'>Prev</a> | "
										Session("page") = CurrPage2
									Else
										Response.Write "<font size='1' face='trebuchet MS'>Prev | "
									End If
									If Not CurrPage2 = TotalPages2 Then
										Response.Write "<a href='" & ScriptName & "?LogPage=" & CurrPage & "&LogPage2=" & CurrPage2 + 1 & "&LogPage3=" & CurrPage3 & "&LogPage4=" & CurrPage4 &  "&Mnum=" & Request("Mnum") & "'>Next</font></a>"
										Session("page") = CurrPage2
									Else
										Response.Write "Next</font>"
									End If
									
								%>
								</td>
								<td colspan='3' align='right'><font size='1' face='trebuchet MS'>Phone calls <%=CurrPage2%> of <%=TotalPages2%></font></td>
							</tr>
							
							<tr><td colspan='2' align='center'><font size='1' face='trebuchet MS'>Date</font></td>
									<td colspan='2' align='center'><font size='1' face='trebuchet MS'>Comment</font></td>
								</tr>
							<%=strLog2%>
						</table>
						</td>
						<td>&nbsp;</td>
						<td valign='top'>
						<table border='1'>
							<tr bgcolor='#040C8B'>
								<td colspan='5' align='center' style="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=0, startColorstr=#040C8B, endColorstr=#c7c9e5);">
									<font size='2' face='trebuchet MS' color='white'>Misc. Contacts</font>
									</td></tr>
							<tr><td colspan='2' width='100px' align='center'>
								<%	'misc
									If Not CurrPage3 = 1 Then 
										Response.Write "<a href='" & ScriptName & "?LogPage=" & CurrPage &  "&LogPage2=" & CurrPage2  &  "&LogPage3=" & CurrPage3 - 1 & "&LogPage4=" & CurrPage4 & "&Mnum=" & Request("Mnum") & "'><font size='1' face='trebuchet MS'>Prev</a> | "
										Session("page") = CurrPage3
									Else
										Response.Write "<font size='1' face='trebuchet MS'>Prev | "
									End If
									If Not CurrPage3 = TotalPages3 Then
										Response.Write "<a href='" & ScriptName & "?LogPage=" & CurrPage & "&LogPage2=" & CurrPage2 &  "&LogPage3=" & CurrPage3 + 1 & "&LogPage4=" & CurrPage4 & "&Mnum=" & Request("Mnum") & "'>Next</font></a>"
										Session("page") = CurrPage3
									Else
										Response.Write "Next</font>"
									End If
									
								%>
								</td>
								<td colspan='3' align='right'><font size='1' face='trebuchet MS'>Misc. Contacts <%=CurrPage3%> of <%=TotalPages3%></font></td>
							</tr>
							
							<tr><td colspan='2' align='center'><font size='1' face='trebuchet MS'>Date</font></td>
									<td colspan='2' align='center'><font size='1' face='trebuchet MS'>Comment</font></td>
								</tr>
							<%=strLog3%>
						</table>
						</td>
							<td>&nbsp;</td>
						
					</tr>
					<tr bgcolor='#040C8B'>
						<td align='center' colspan='7' style="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=0, startColorstr=#040C8B, endColorstr=#c7c9e5);">
							<font size='2' face='trebuchet ms' color='white'>New Log Entries</font></td></tr>
					<tr>
						<td colspan='7'>
							<table border='0' width='100%'>
								<tr>
									<td width='25px'>&nbsp;</td>
									<td valign='top' align='center'>
										<input style='font-size: 10px; height: 20px;' size='9'  maxlength='10' name='txtSVD'>
									</td>
									<td>
										<textarea cols='18' name='txtSVC'></textarea>
									</td>
									<td  width='25px'>&nbsp;</td>
									<td valign='top' align='center'>
										<input style='font-size: 10px; height: 20px;' size='9' maxlength='10' name='txtPCD'>
									</td>
									<td>
										<textarea cols='18' name='txtPCC'></textarea>
									</td>
									<td  width='25px'>&nbsp;</td>
									<td valign='top' align='center'>
										<input style='font-size: 10px; height: 20px;' size='9' maxlength='10' name='txtMCD'>
									</td>
									<td>
										<textarea cols='18' name='MCon'></textarea>
									</td>
								
								</tr>
							</table>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td colspan='7' align='center'>
							<input type='button' style='width: 200px;' class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'"  value='Save Log Entry' onclick='document.frmConDetLog.submit();'>
							<input type='button' style='width: 200px;' value='Delete Checked Log Entry' class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='JavaScript: DelLog();'>
						</td>
						<input type='hidden' name='ctr' value='<%=ctr%>'>
						<input type='hidden' name='ctr2' value='<%=ctr2%>'>
						<input type='hidden' name='ctr3' value='<%=ctr3%>'>
						<input type='hidden' name='ctr4' value='<%=ctr4%>'>
						<input type='hidden' name='Mnum' value="<%=Request("MNum")%>">
					</tr>
				</td>
				</table>
			
			</div>
		</form>
		<!-- #include file="_boxdown.asp" -->
	</body>
</html>
<% Session("MSG") = "" %>