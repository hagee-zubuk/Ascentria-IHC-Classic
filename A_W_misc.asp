<%Language=VBScript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_security.asp" -->
<%
ScriptName = Request.ServerVariables("SCRIPT_NAME")
		''''''Site
		NumPerPage = 10
		
		If Request.QueryString("LogPage") = "" Then
			CurrPage = 1
		Else
			CurrPage = CInt(Request.QueryString("LogPage"))
		End if
		Set rsLog = Server.CreateObject("ADODB.RecordSet")
		sqlLog = "SELECT * FROM [W_vio_t] WHERE [ssn] = '" & Request("wID") & "' ORDER BY [viodate] DESC"
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
			strLog = "<tr><td colspan='3' align='center'><font size='1'>N/A</font></td></tr>"
		End If
		ctr = 0
		Do While Not rsLog.EOF And ctr < rsLog.PageSize
			If rsLog("viodate") <> "" Then
				if Z_IsOdd(ctr) = true then 
						kulay = "#FFFAF0" 
					else 
						kulay = "#FFFFFF"
					end if
				verbalwarn = ""	
				writewarn = ""
				finalwarn = ""
				strLog = strLog & "<tr bgcolor='" & kulay & "'><td align='center'><input type='checkbox' name='chkwarn" & ctr & _
					"' value='" & rsLog("index") & "'></td><td align='center'><input type='text' " & _
					"style='font-size: 10px; height: 20px;' size='9' maxlength='10' name='txtwarn" & ctr & "' value='" & _
					rsLog("viodate") & "'></td><td align='center'><textarea cols='18' name='vionote" & ctr & "'>" & rsLog("notes") & "</textarea></td></tr>"			
	
				'rsLog.MoveNext
				ctr = ctr + 1
			End If
			rsLog.MoveNext
			
		Loop
		rsLog.Close
%>
<html>
	<head>
		<title>LSS - In-Home Care - PCSP Worker Details - Violations</title>
		<script language='JavaScript'>
			function DelWarn()
			{
				document.frmWorDetVio.action = "A_W_Del.asp?page=3";
				document.frmWorDetVio.submit();
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
		<form method='post' name='frmWorDetVio' action='NewVio.asp?add=1'>
			<br>
			<center>
				<table border='0'>
					<tr>
						<td colspan='4' align='center' width='500px'>
							<font size='2' face='trebuchet MS'><b><u>PCSP Worker Details - Violations</u></b></font>
							<a href='A_Worker.asp?WID=<%=Request("WID")%>' style='text-decoration:none'><font size='1' face='trebuchet MS'>[Details]</font></a>
							
							<a href='A_w_Files.asp?WID=<%=Request("WID")%>' style='text-decoration:none'><font size='1' face='trebuchet MS'>[Files]</font></a>
							<a href="A_W_Skills.asp?WID=<%=Request("WID")%>" style='text-decoration: none;'>
								<font color='blue' size='1' face='trebuchet MS'>[Skills]</font>
							</a>
							<a href='workCon.asp?WID=<%=Request("WID")%>' style='text-decoration:none'><font size='1' face='trebuchet MS'>[List]</font></a>
							<a href="A_W_log.asp?WID=<%=Request("WID")%>" style='text-decoration: none;'>
								<font color='blue' size='1' face='trebuchet MS'>[Log]</font>
							</a>
							<font size='2' face='trebuchet MS'>[Violations]</font>
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
				
				<table cellspacing='0' cellpadding='0' border='0' align='center'>
					<tr>
						<td valign='top'>
						<table border='1' width='100%'>
							<tr bgcolor='#040C8B'><td colspan='3' align='center' style="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=0, startColorstr=#040C8B, endColorstr=#c7c9e5);">
								<font size='2' face='trebuchet MS' color='white'>Violations</font></td></tr>
							<tr><td colspan='2'  align='center'>
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
							
								<td align='right'><font size='1' face='trebuchet MS'>Violations <%=CurrPage%> of <%=TotalPages%></font></td>
							</tr>
							<tr>
								<td>&nbsp;</td>
								<td align='center'><font size='1' face='trebuchet MS'>Warning Date</font></td>
									<td align='center'><font size='1' face='trebuchet MS'>Notes</font></td>
								</tr>
							<%=strLog%>
						</table>
						</td>
						
					<tr bgcolor='#040C8B'>
						<td align='center' colspan='1' style="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=0, startColorstr=#040C8B, endColorstr=#c7c9e5);">
							<font size='2' face='trebuchet ms' color='white'>New Violation Entries</font></td></tr>
					<tr>
						<td colspan='1'>
							<table border='0' align='left' width='100%'>
								<tr>
									<td width='32px'>&nbsp;</td>
									<td valign='top' align='center'>
										<input style='font-size: 10px; height: 20px;' size='9'  maxlength='10' name='txtwarn'>
									</td>
									<td align='center'>
										<textarea cols='18' name='vionote'></textarea>
									</td>
									
								</tr>
							</table>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td colspan='5' align='center'>
							<input type='button' style='width: 200px;' class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" value='Save' onclick='document.frmWorDetVio.submit();'>
							<input type='button' style='width: 200px;' class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" value='Delete'onclick='JavaScript: DelWarn();'>
						</td>
						<input type='hidden' name='ctr' value='<%=ctr%>'>
						<input type='hidden' name='WID' value="<%=Request("WID")%>">
					</tr>
				</td>
				
				</table>
			
		</form>
		<!-- #include file="_boxdown.asp" -->
	</body>
</html>
<% Session("MSG") = "" %>