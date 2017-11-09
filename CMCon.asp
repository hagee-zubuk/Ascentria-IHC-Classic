<%Language=VBScript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<%
	Set tblCon = Server.CreateObject("ADODB.RecordSet")
	'Set tblName = Server.CreateObject("ADODB.RecordSet")
	
	sqlCon = "SELECT * FROM [CMCon_t], [Consumer_t] WHERE [CMID] = '" & Request("ID") & "' AND CID = medicaid_number ORDER BY lname, Fname"
	'response.write sqlCon
	tblCon.Open sqlCon, g_strCONN, 1, 3
	If Not tblCon.EOF Then
		ctr = 0
		Do Until tblCon.EOF
			if Z_IsOdd(ctr) = true then 
				kulay = "#FFFAF0" 
			else 
				kulay = "#FFFFFF"
			end if
			tmpName = tblCon("Lname") & ", " & tblCon("Fname")
			strCon = strCon & "<tr bgcolor='" & kulay & "'><td><a href='A_Consumer.asp?Mnum=" & tblCon("CID") & _
					"'>" & _
					"&nbsp;<font size='1'>" & tmpName & "</font>&nbsp;</a></td></tr>" & vbCrLf
			tblCon.MoveNext
			ctr = ctr + 1
		Loop
	Else
		strCon = "<tr><td align='center'><font size='1'>N/A</font></td></tr>"
	End If
	tblCon.Close
	'Set tblName = Nothing
	Set tblCon = Nothing	
%>
<html>
	<head>
		<title>LSS - In-Home Care - Case Manager Details - List</title>
	</head>
	<body bgcolor='white' LEFTMARGIN='0' TOPMARGIN='0'>
		
		<form method='post' name='frmCaseList' action='#'>
		<!-- #include file="_boxup.asp" -->
		<!-- #include file="_NavHeader.asp" -->
		<br>
		<center>
		<table border='0'>
			<tr>
				
				<td colspan='8' align='center' width='500px'>
		<font size='2' face='trebuchet MS'><b><u>Case Manager - List</u></b></font>
		&nbsp;
		<a href='A_Case.asp' style='text-decoration: none;'>
			<font size='1' face='trebuchet MS' color='blue'>[Details]</font></a>
				<font size='2' face='trebuchet MS'>[List]</font>
		</td>
	</tr>
	<tr><td>&nbsp;</td></tr>
	<tr><td>&nbsp;</td></tr>
			<tr>
				
				<td><font size='1' face='trebuchet MS'><u>Name:</u></td>
				<td><input style='font-size: 10px; height: 20px; width: 120px;' readonly value='<%=Session("CMname")%>'></td>
				<td>
					&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				</td><td>
					&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;	
				</td>
				<td>
					&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;	
				</td>
				<td>
					&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;	
				</td>
				<td>
					&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;	
				</td>
				<td>
					&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;	
				</td>
			</tr>
		</table>
		<br><br>
		<table border='1'>
			<tr bgcolor='#040C8B'>
				<td colspan='2' align='center' style="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=0, startColorstr=#040C8B, endColorstr=#c7c9e5);">
					<font size='2' face='trebuchet MS' color='white'><b>Consumer List</b></font></td></tr>
			<%=strCon%>
			
		</table>
		<br>
		<!-- #include file="_boxdown.asp" -->
		</form>
	</body>
</html>
<%
	Session("MSG") = ""
%>