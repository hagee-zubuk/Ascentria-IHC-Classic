<%Language=VBScript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_security.asp" -->
<%
	If Request("WID") <> "" Then
		ScriptName = Request.ServerVariables("SCRIPT_NAME")
		''''''Site
		NumPerPage = 10
		
		If Request.QueryString("LogPage") = "" Then
			CurrPage = 1
		Else
			CurrPage = CInt(Request.QueryString("LogPage"))
		End if
		Set rsLog = Server.CreateObject("ADODB.RecordSet")
		sqlLog = "SELECT * FROM [W_log_t] WHERE [ssn] = '" & Request("WID") & "' ORDER BY [siteV] DESC"
		'response.write "<!--SQL: " & sqlLog & "-->"
		rsLog.Open sqlLog, g_strCONN, 1, 3
		TotalPages = 1
		If Not rsLog.EOF Then
			rsLog.MoveFirst
			rsLog.PageSize = NumPerPage
			TotalPages = rsLog.PageCount
			response.write "<!--page1: " & TotalPages & " -->"
			rsLog.AbsolutePage = CurrPage
		Else
			strLog = "<tr><td colspan='4' align='center'><font size='1'>N/A</font></td></tr>"
		End If
		ctr = 0
		Do While Not rsLog.EOF And ctr < rsLog.PageSize
			If rsLog("siteV") <> "" Then
				if Z_IsOdd(ctr) = true then 
						kulay = "#FFFAF0" 
					else 
						kulay = "#FFFFFF"
					end if
				strLog = strLog & "<tr bgcolor='" & kulay & "'><td align='center'><input type='checkbox' name='chkSV" & ctr & _
					"' value='" & rsLog("index") & "'></td><td align='center'><input type='text' " & _
					"style='font-size: 10px; height: 20px;' size='9' maxlength='10' name='txtSVD" & ctr & "' value='" & _
					rsLog("sitev") & "'></td><td colspan='2'><textarea cols='18' name='Vcom" & ctr & "'>" & rsLog("scom") & "</textarea></td></tr>"
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
		sqlLog2 = "SELECT * FROM [w_Log_t] WHERE [ssn] = '" & Request("WID") & "' ORDER BY [phoneC] DESC"
		'response.write "<!--SQL2: " & sqlLog2 & "-->"
		rsLog2.Open sqlLog2, g_strCONN, 1, 3
		TotalPages2 = 1
		If Not rsLog2.EOF Then
			rsLog2.MoveFirst
			rsLog2.PageSize = NumPerPage2
			TotalPages2 = rsLog2.PageCount
			response.write "<!--page2: " & rsLog2.PageCount & " -->"
			rsLog2.AbsolutePage = CurrPage2
		Else
			strLog2 = "<tr><td colspan='4' align='center'><font size='1'>N/A</font></td></tr>"
		End If
		ctr2 = 0
		Do While Not rsLog2.EOF AND ctr2 < rsLog2.PageSize 
				If rsLog2("phoneC") <> "" Then
					if Z_IsOdd(ctr2) = true then 
							kulay = "#FFFAF0" 
						else 
							kulay = "#FFFFFF"
						end if
				strLog2 = strLog2 & "<tr bgcolor='" & kulay & "'><td align='center'><input type='checkbox' name='chkPC" & ctr2 & _
					"' value='" & rsLog2("index") & "'></td><td align='center'><input type='text' " & _
					"style='font-size: 10px; height: 20px;' size='9' maxlength='10' name='txtPCD" & ctr2 & "' value='" & _
					rsLog2("phonec") & "'></td><td colspan='2'><textarea cols='18' name='Pcom" & ctr2 & "'>" & rsLog2("PCom") & "</textarea></td></tr>"
					'rsLog2.MoveNext
					ctr2 = ctr2 + 1
				End If
			rsLog2.MoveNext
		Loop
		rsLog2.Close
		Set rsLog2 = Nothing
		
		'''misc
		NumPerPage3 = 10
		
		If Request.QueryString("LogPage3") = "" Then
			CurrPage3 = 1
		Else
			CurrPage3 = CInt(Request.QueryString("LogPage3"))
		End if
		Set rsLog3 = Server.CreateObject("ADODB.RecordSet")
		sqlLog3 = "SELECT * FROM [w_log_T] WHERE [ssn] = '" & Request("WID") & "' ORDER BY [misc] DESC"
		'response.write "<!--SQL3: " & sqlLog3 & "-->"
		rsLog3.Open sqlLog3, g_strCONN, 1, 3
		TotalPages3 = 1
		If Not rsLog3.EOF Then
			rsLog3.MoveFirst
			rsLog3.PageSize = NumPerPage3
			TotalPages3 = rsLog3.PageCount
			response.write "<!--page3: " & rsLog3.PageCount & " -->"
			rsLog3.AbsolutePage = CurrPage3
		Else
			strLog3 = "<tr><td colspan='4' align='center'><font size='1'>N/A</font></td></tr>"
		End If
		ctr3 = 0
		Do While Not rsLog3.EOF AND ctr3 < rsLog3.PageSize 
				If rsLog3("misc") <> "" Then
					if Z_IsOdd(ctr3) = true then 
							kulay = "#FFFAF0" 
						else 
							kulay = "#FFFFFF"
						end if
					strLog3 = strLog3 & "<tr bgcolor='" & kulay & "'><td align='center'><input type='checkbox' name='chkMC" & ctr3 & _
						"' value='" & rsLog3("index") & "'></td><td align='center'><input type='text' " & _
						"style='font-size: 10px; height: 20px;' size='9'  maxlength='10' name='txtMCD" & ctr3 & "' value='" & _
						rsLog3("misc") & "'></td><td colspan='2'><textarea cols='18' name='MCon" & ctr3 & "'>" & rsLog3("mcom") & "</textarea></td></tr>"
					'rsLog2.MoveNext
					ctr3 = ctr3+ 1
				End If
			rsLog3.MoveNext
		Loop
		rsLog3.Close
		Set rsLog3 = Nothing
	Else
		Session("MSG") = "Sorry. To maintain integrity of database, please Sign-in again."
		Response.Redirect "default.asp"
	End If
%>
<html>
	<head>
		<title>LSS - In-Home Care - PCSP Worker Details - Log</title>
		<script language='JavaScript'>
			function DelLog()
			{
				document.frmWorDetLog.action = "A_W_Del.asp?page=2";
				document.frmWorDetLog.submit();
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
		<form method='post' name='frmWorDetLog' action='NewLog.asp?add=1'>
		
			<br>
			<center>
				<table border='0'>
					<tr>
						<td colspan='4' align='center' width='500px'>
							<font size='2' face='trebuchet MS'><b><u>PCSP Worker Details - Log</u></b></font>
							<a href='A_Worker.asp?WID=<%=Request("WID")%>' style='text-decoration:none'><font size='1' face='trebuchet MS'>[Details]</font></a>
							
							<a href='A_w_Files.asp?WID=<%=Request("WID")%>' style='text-decoration:none'><font size='1' face='trebuchet MS'>[Files]</font></a>
							<a href="A_W_Skills.asp?WID=<%=Request("WID")%>" style='text-decoration: none;'>
								<font color='blue' size='1' face='trebuchet MS'>[Skills]</font>
							</a>
							<a href='workCon.asp?WID=<%=Request("WID")%>' style='text-decoration:none'><font size='1' face='trebuchet MS'>[List]</font></a>
							<font size='2' face='trebuchet MS'>[Log]</font>
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
							<input type='text' style='font-size: 10px; height: 20px; width: 120px;' readonly name='Cname' value="<%=Session("wname")%>"></font>
						</td>
					</tr>
				</table>
				<br>
				<div style='OVERFLOW: AUTO; ' align='center'>
				<table cellspacing='0' cellpadding='0' border='0' align='center'>
					<tr>
						<td valign='top'>
						<table border='1'>
							<tr bgcolor='#040C8B'><td colspan='4' align='center' style="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=0, startColorstr=#040C8B, endColorstr=#c7c9e5);">
								<font size='2' face='trebuchet MS' color='white'>Site Visit</font></td></tr>
							<tr><td colspan='2' width='100px' align='center'>
								<%  'Site
									If Not CurrPage = 1 Then 
										Response.Write "<a href='" & ScriptName & "?LogPage=" & CurrPage - 1 &  "&LogPage2=" & CurrPage2 & "&LogPage3=" & CurrPage3 & "&WID=" & Request("WID") & "'><font size='1' face='trebuchet MS'>Prev</a> | "
										Session("page") = CurrPage 
									Else
										Response.Write "<font size='1' face='trebuchet MS'>Prev | "
									End If
									If Not CurrPage = TotalPages Then
										Response.Write "<a href='" & ScriptName & "?LogPage=" & CurrPage + 1 & "&LogPage2=" & CurrPage2 & "&LogPage3=" & CurrPage3 & "&WID=" & Request("WID") & "'>Next</font></a>"
										Session("page") = CurrPage 
									Else
										Response.Write "Next</font>"
									End If
									
								%>
								</td>
								<td>&nbsp;</td>
								<td align='right'><font size='1' face='trebuchet MS'>Site Visit <%=CurrPage%> of <%=TotalPages%></font></td>
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
							<tr bgcolor='#040C8B'><td colspan='4' align='center' style="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=0, startColorstr=#040C8B, endColorstr=#c7c9e5);">
								<font size='2' face='trebuchet MS' color='white'>Phone Call</font></td></tr>
							<tr><td colspan='2' width='100px' align='center'>
								<%	'phone
									If Not CurrPage2 = 1 Then 
										Response.Write "<a href='" & ScriptName & "?LogPage=" & CurrPage &  "&LogPage2=" & CurrPage2 - 1 & "&LogPage3=" & CurrPage3  & "&WID=" & Request("WID") & "'><font size='1' face='trebuchet MS'>Prev</a> | "
										Session("page") = CurrPage2
									Else
										Response.Write "<font size='1' face='trebuchet MS'>Prev | "
									End If
									If Not CurrPage2 = TotalPages2 Then
										Response.Write "<a href='" & ScriptName & "?LogPage=" & CurrPage & "&LogPage2=" & CurrPage2 + 1 & "&LogPage3=" & CurrPage3 & "&WID=" & Request("WID") & "'>Next</font></a>"
										Session("page") = CurrPage2
									Else
										Response.Write "Next</font>"
									End If
									
								%>
								</td>
								<td>&nbsp;</td>
								<td align='right'><font size='1' face='trebuchet MS'>Phone calls <%=CurrPage2%> of <%=TotalPages2%></font></td>
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
										Response.Write "<a href='" & ScriptName & "?LogPage=" & CurrPage &  "&LogPage2=" & CurrPage2  &  "&LogPage3=" & CurrPage3 - 1 & "&WID=" & Request("WID") & "'><font size='1' face='trebuchet MS'>Prev</a> | "
										Session("page") = CurrPage3
									Else
										Response.Write "<font size='1' face='trebuchet MS'>Prev | "
									End If
									If Not CurrPage3 = TotalPages3 Then
										Response.Write "<a href='" & ScriptName & "?LogPage=" & CurrPage & "&LogPage2=" & CurrPage2 &  "&LogPage3=" & CurrPage3 + 1 & "&WID=" & Request("WID") & "'>Next</font></a>"
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
					</tr>
					<tr bgcolor='#040C8B'>
						<td align='center' colspan='5' style="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=0, startColorstr=#040C8B, endColorstr=#c7c9e5);">
							<font size='2' face='trebuchet ms' color='white'>New Log Entries</font></td></tr>
					<tr>
						<td colspan='5'>
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
						<td colspan='5' align='center'>
							<input type='button' style='width: 200px;' class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" value='Save Log Entry' onclick='document.frmWorDetLog.submit();'>
							<input type='button' style='width: 200px;' class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" value='Delete Checked Log Entry' onclick='JavaScript: DelLog();'>
						</td>
						<input type='hidden' name='WID' value="<%=Request("WID")%>">
						<input type='hidden' name='ctr' value='<%=ctr%>'>
						<input type='hidden' name='ctr2' value='<%=ctr2%>'>
						<input type='hidden' name='ctr3' value='<%=ctr3%>'>
					</tr>
				</td>
				
				</table>
			</div>
		</form>
		<!-- #include file="_boxdown.asp" -->
	</body>
</html>
<% Session("MSG") = "" %>