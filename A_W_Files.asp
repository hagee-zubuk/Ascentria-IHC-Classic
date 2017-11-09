<%language=vbscript%>
<!-- #include file="_Files.asp" -->

<!-- #include file="_Utils.asp" -->
<!-- #include file="_security.asp" -->
<%
	If Request("WID") <> "" Then
		Set tblFiles = Server.CreateObject("ADODB.Recordset")
		sqlFiles = "SELECT * FROM W_Files_t WHERE SSN = '" & Request("WID") & "' "
		tblFiles.Open sqlFiles, g_strCONN, 1, 3
			If Not tblFiles.EOF Then
				MCNum = tblFiles("SSN")
				If tblFiles("t1") = True Then t1 = "checked"
				If tblFiles("t2") = True Then t2 = "checked"
				If tblFiles("t3") = True Then t3 = "checked"
				If tblFiles("tb") = True Then tb = "checked"
				If tblFiles("tb2") = True Then tb2 = "checked"
				If tblFiles("phy") = True Then phy = "checked"
				If tblFiles("orient") = True Then chkorient = "checked"
				If tblFiles("pptrain") = True Then chkpp = "checked"
				If tblFiles("lnaactive") = True Then chklnaa = "checked"
				If tblFiles("lnainactive") = True Then chklnai = "checked"
				tbdate = tblFiles("tbdate")
				tbdate2 = tblFiles("tb2date")
				orientdate = tblFiles("orientdate")
				if tblFiles("essentials") Then essentials = "checked"
				essentialsdate = tblFiles("essentialsdate")
			Else
				tblFiles.AddNew
				tblFiles("SSN") = Request("WID")
				tblFiles.Update
			End If
		tblFiles.Close
		Set tblFiles = Nothing
	
		'training
		ScriptName = Request.ServerVariables("SCRIPT_NAME")
		''''''Site
		NumPerPage = 10
		
		If Request.QueryString("LogPage") = "" Then
			CurrPage = 1
		Else
			CurrPage = CInt(Request.QueryString("LogPage"))
		End if
		Set rsLog = Server.CreateObject("ADODB.RecordSet")
		sqlLog = "SELECT * FROM [W_log_t] WHERE [ssn] = '" & Request("WID") & "' ORDER BY [train] DESC"
		'response.write "<!--SQL: " & sqlLog & "-->"
		rsLog.Open sqlLog, g_strCONN, 1, 3
		TotalPages = 1
		If Not rsLog.EOF Then
			rsLog.MoveFirst
			rsLog.PageSize = NumPerPage
			TotalPages = rsLog.PageCount
			rsLog.AbsolutePage = CurrPage
		Else
			strLog = "<tr><td colspan='3' align='center'><font size='1'>N/A</font></td></tr>"
		End If
		ctr = 0
		Do While Not rsLog.EOF And ctr < rsLog.PageSize
			If rsLog("train") <> "" Then
				if Z_IsOdd(ctr) = true then 
						kulay = "#FFFAF0" 
					else 
						kulay = "#FFFFFF"
					end if
				strLog = strLog & "<tr bgcolor='" & kulay & "'><td align='center'><input type='checkbox' name='chktrain" & ctr & _
					"' value='" & rsLog("index") & "'></td><td align='center'><input type='text' " & _
					"style='font-size: 10px; height: 20px;' size='25' maxlength='50' name='txttrain" & ctr & "' value='" & _
					rsLog("train") & "'></td><td align='center'><input type='text' " & _
					"style='font-size: 10px; height: 20px;' size='7' maxlength='8' name='txthrs" & ctr & "' value='" & _
					rsLog("thrs") & "'></td><td align='center'><textarea cols='18' name='trainnote" & ctr & "'>" & rsLog("tcom") & "</textarea></td></tr>"			
	
				'rsLog.MoveNext
				ctr = ctr + 1
			End If
			rsLog.MoveNext
			
		Loop
		rsLog.Close
	Else
		Session("MSG") = "Sorry. To maintain integrity of database, please Sign-in again."
		Response.Redirect "default.asp"
	End If
%>
<html>
	<head>
		<title>LSS - In-Home Care - PCSP Worker Details - Files</title>
		<script language='JavaScript'>
			function WDF_Edit()
			{
				document.frmWorDetFil.action = "A_W_Action.asp?page=2&del=0";
				document.frmWorDetFil.submit();
			}
			function DelList()
			{
				document.frmWorDetFil.action = "A_W_Action.asp?page=2&del=1";
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
							<font size='2' face='trebuchet MS'>[Files]</font>
							<a href='A_W_Skills.asp?WID=<%=Request("WID")%>' style='text-decoration: none;'>
								<font color='blue' size='1' face='trebuchet MS'>[Skills]</font>
							</a>
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
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							<input type='checkbox' name='chkorient' <%=chkorient%>>
							<font size='1' face='trebuchet MS'><u>Orientation</u>&nbsp;
							<font size='1' face='trebuchet MS'><u>Datestamp:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 80px;' name='orientdate' maxlength='10' value="<%=orientdate%>" onKeyUp="javascript:return maskMe(this.value,this,'2,5','/');" onBlur="javascript:return maskMe(this.value,this,'2,5','/');"></font>
							<font size='1' face='trebuchet MS'>(mm/dd/yyyy)</font>
						</td>
						
					</tr>
					<tr>
						<td align='left'>
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							<input type='checkbox' name='essentials' <%=essentials%>>
							<font size='1' face='trebuchet MS'><u>Essentials Training</u>&nbsp;
							<font size='1' face='trebuchet MS'><u>Datestamp:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 80px;' name='essentialsdate' maxlength='10' value="<%=essentialsdate%>" onKeyUp="javascript:return maskMe(this.value,this,'2,5','/');" onBlur="javascript:return maskMe(this.value,this,'2,5','/');"></font>
							<font size='1' face='trebuchet MS'>(mm/dd/yyyy)</font>
						</td>
						
					</tr>
					<tr>
						<td align='left'>
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							<input type='checkbox' name='chkpp' <%=chkpp%>>
							<font size='1' face='trebuchet MS'><u>PP Training</u>&nbsp;
						</td>
						
					</tr>
					<tr>
						<td align='left'>
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							<input type='checkbox' name='chktb' <%=tb%>>
							<font size='1' face='trebuchet MS'><u>TB Test 1</u>&nbsp;
							<font size='1' face='trebuchet MS'><u>Datestamp:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 80px;' name='tbdate' maxlength='10' value="<%=tbdate%>" onKeyUp="javascript:return maskMe(this.value,this,'2,5','/');" onBlur="javascript:return maskMe(this.value,this,'2,5','/');"></font>
							<font size='1' face='trebuchet MS'>(mm/dd/yyyy)</font>
						</td>
						
					</tr>
					<tr>
						<td align='left'>
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							<input type='checkbox' name='chktb2' <%=tb2%>>
							<font size='1' face='trebuchet MS'><u>TB Test 2</u>&nbsp;
							<font size='1' face='trebuchet MS'><u>Datestamp:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 80px;' name='tbdate2' maxlength='10' value="<%=tbdate2%>" onKeyUp="javascript:return maskMe(this.value,this,'2,5','/');" onBlur="javascript:return maskMe(this.value,this,'2,5','/');"></font>
							<font size='1' face='trebuchet MS'>(mm/dd/yyyy)</font>
						</td>
						
					</tr>
					<tr>
						<td align='left'>
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							<input type='checkbox' name='chklnaa' <%=chklnaa%>>
							<font size='1' face='trebuchet MS'><u>LNA Active</u>&nbsp;
						</td>
						
					</tr>
					<tr>
						<td align='left'>
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							<input type='checkbox' name='chklnai' <%=chklnai%>>
							<font size='1' face='trebuchet MS'><u>LNA Inactive</u>&nbsp;
						</td>
						
					</tr>
					<tr>
						<td align='left'>
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							<input type='checkbox' name='chkphy' <%=phy%>>
							<font size='1' face='trebuchet MS'><u>Physical</u>&nbsp;
						</td>
						
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td valign='top'>
						<table border='1' width='100%'>
							<tr bgcolor='#040C8B'><td colspan='4' align='center' style="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=0, startColorstr=#040C8B, endColorstr=#c7c9e5);">
								<font size='2' face='trebuchet MS' color='white'>Training</font>
								<a href='JavaScript: DelList();' style='text-decoration: none'><font size='1' face='trebuchet ms' color='white'>[Delete]</font></a>
								</td></tr>
							<tr><td colspan='3'  align='center'>
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
							
								<td align='right'><font size='1' face='trebuchet MS'>Training <%=CurrPage%> of <%=TotalPages%></font></td>
							</tr>
							<tr>
								<td>&nbsp;</td>
								<td align='center'><font size='1' face='trebuchet MS'>Training</font></td>
								<td align='center'><font size='1' face='trebuchet MS'>Hours</font></td>
									<td align='center'><font size='1' face='trebuchet MS'>Notes</font></td>
								</tr>
							<%=strLog%>
						</table>
						</td>
						
					<tr bgcolor='#040C8B'>
						<td align='center' colspan='1' style="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=0, startColorstr=#040C8B, endColorstr=#c7c9e5);">
							<font size='2' face='trebuchet ms' color='white'>New Training Entries</font></td></tr>
					<tr>
						<td colspan='1'>
							<table border='0' align='left' width='100%'>
								<tr>
									<td width='32px'>&nbsp;</td>
									<td valign='top' align='center'>
										<input style='font-size: 10px; height: 20px;' size='25'  maxlength='50' name='txttrain'>
									</td>
									<td valign='top' align='center'>
										<input style='font-size: 10px; height: 20px;' size='7'  maxlength='6' name='txthrs'>
									</td>
									<td align='center'>
										<textarea cols='18' name='trainnote'></textarea>
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