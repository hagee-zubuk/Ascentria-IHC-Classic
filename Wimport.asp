<%@Language=VBScript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_security.asp" -->
<%
wid = Request("wid")
If Z_fixNull(request("fname")) <> "" Then 
		viewpath = uploadFilePath & request("fname")
		strMSG = Request("fn")
End If
If wid <> "" Then
	Set tblStatus = Server.CreateObject("ADODB.Recordset")
	sqlStatus = "SELECT * FROM uploads_T WHERE wid = '" & wid & "' ORDER BY [datestamp] DESC"
	tblStatus.Open sqlStatus, g_strCONN, 1, 3
	If Not tblStatus.EOF Then
		Do Until tblStatus.EOF
			strVform = strVform & "<tr>" & _
				"<td align='center' style='border: 1px solid;'><font size='1' face='trebuchet MS'>" & tblStatus("datestamp") & "</font></td>" & _
				"<td align='center' style='border: 1px solid;'><font size='1' face='trebuchet MS'>" & tblStatus("ofilename") & "</font></td>" & _
				"<td align='center' style='border: 1px solid;'><a style='text-decoration: none;' href='wimport.asp?wid=" & wid & "&fname=" & tblStatus("filename") & "&fn=" & tblStatus("ofilename")  & "'><img src='images/zoom.gif' title='view file'></a>" & _
				"</tr>"
			tblStatus.MoveNext
		Loop
	Else
		
	End If
	tblStatus.Close
	Set tblStatus = Nothing	

Else
	Session("MSG") = "Sorry. To maintain integrity of database, please Sign-in again."
	Response.Redirect "default.asp"
End If
%>
<html>
	<head>
		<title>LSS - In-Home Care - PCSP WORKER Details - Uploads</title>
		<script language='JavaScript'>
			<!--
			function upload(xxx)
			{
				newwindow = window.open('upload.asp?wid=' + xxx,'','height=150,width=500,scrollbars=1,directories=0,status=0,toolbar=0,resizable=0');
					if (window.focus) {newwindow.focus()}
			}
			-->	
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
		<form method='post' name='frmUpload' enctype="multipart/form-data">
			<!-- #include file="_boxup.asp" -->
			<!-- #include file="_NavHeader.asp" -->
			<br>
			<table border='0' align='center'>
				<tr>
					<td colspan='4' align='center' width='500px'>
						<font size='2' face='trebuchet MS'><b><u>PCSP Worker Details</u></b></font>
						<a href='A_Worker.asp?WID=<%=WID%>' style='text-decoration:none'><font size='1' face='trebuchet MS'>[Details]</font></a>
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
						<font size='2' face='trebuchet MS'>[import]</font>
					</td>
				</tr>
				<tr><td colspan='2' align='center'><font color='red'  face='trebuchet MS' size='1'>&nbsp;<%=Session("MSG")%>&nbsp;</font></td></tr>
				<tr>
						<td>
							<font size='1' face='trebuchet MS'><u>Name:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 120px;' readonly name='Cname' value="<%=Session("wname")%>"></font>
							<input type='button' name='btnupload' value='Upload File' onclick="upload('<%=wid%>');" class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'">
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
				<tr>
					<td>
						<table style="border: 1px solid;" border=1>
							<tr>
								<td valign='top'>
									<table>
										<tr >
											<td align="center" style="border: 1px solid;"><font size='1' face='trebuchet MS'><u>Timestamp</u></font></td>
											<td align="center" style="border: 1px solid;"><font size='1' face='trebuchet MS'><u>Filename</u></font></td>
											<td align="center" style="border: 1px solid;"><font size='1' face='trebuchet MS'>&nbsp;</font></td>
										</tr>
										<%=strVform%>
									</table>
								</td>
								<td>
									<table>
										<tr>
											<td align='center'>
												<font size='1' face='trebuchet MS'><%=strMSG%></font>
											</td>
										</tr>
										<tr>
											<td align="center"  colspan='2'>
												<iframe src="files.asp?fpath=<%=viewpath%>" width="650" height="450"></iframe>
											</td>
										</tr>
									</table>
								</td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
			<!-- #include file="_boxdown.asp" -->
		</form>
	</body>
</html>
<%
Session("MSG") = "" 
%>