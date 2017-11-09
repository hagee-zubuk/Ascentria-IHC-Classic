<%Language=VBScript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_security.asp" -->
<%
	'''''''WORKER'''''
	Set tblWork = Server.CreateObject("ADODB.RecordSet")
	Set tblNme = Server.CreateObject("ADODB.RecordSet")
	
	sqlWork = "SELECT * FROM [ConWork_t] WHERE [CID] = '" & Request("ID") & "' " 
	'response.write "SQL: " & sqlWork & "<br>"
	tblWork.Open sqlWork, g_strCONN, 1, 3
On Error Resume Next
	If Not tblWork.EOF Then
		ctrW = 0
		'response.write "WID: " & tblWork("WID")
		Do Until tblWork.EOF
			tmpWID = tblWork("WID") 
			sqlName = "SELECT * FROM [Worker_t] WHERE [Worker_t.Index] = " & tmpWID
			'response.write "<br>" & "SQL2: " & sqlName
			tblNme.Open sqlName, g_strCONN, 1, 3
			tmpName = tblNme("Lname") & ", " & tblNme("Fname")
			strWork = strWork & "<tr><td align='center'><input type='checkbox' name='chkW" & ctrW & "' value='" & _
					tblWork("ID") & "'<td><a href='A_Worker.asp?WID=" & _
					tblNme("Social_Security_Number") & "'><font size='2' face='trebuchet MS'>&nbsp;" & tblNme("Social_Security_Number") & _
					"&nbsp;</font></a></td><td>" & _
					"<font size='2'>&nbsp;" & tmpName & "&nbsp;</font></td></tr>"
			tblNme.Close
			tblWork.MoveNext
			ctrW = ctrW + 1
		Loop
	Else
		strWork = "<tr><td align='center'><font size='2'>N/A</font></td></tr>"
	End If
	tblWork.Close
	Set tblNme = Nothing
	Set tblWork = Nothing	
	'''''''''''CASE MANAGER'''''
	
	Set tblCM = Server.CreateObject("ADODB.RecordSet")
	Set tblNme = Server.CreateObject("ADODB.RecordSet")
	
	sqlCM = "SELECT * FROM [CMCon_t] WHERE [CID] = '" & Request("ID") & "' " 
	
	tblCM.Open sqlCM, g_strCONN, 1, 3
	If Not tblCM.EOF Then
		ctr = 0
		Do Until tblCM.EOF
			sqlName = "SELECT * FROM [Case_Manager_t] WHERE [Index] = " & tblCM("CMID") 
			tblNme.Open sqlName, g_strCONN, 1, 3
			tmpName = tblNme("Lname") & ", " & tblNme("Fname")
			strCM = strCM & "<tr><td align='center'><input type='checkbox' name='chkCM" & ctr & "' value='" & _
					tblNme("Index") & "'</td><td><a href='A_Case.asp?CaID=" & tblNme("Index") & "'><font size='2' face='trebuchet MS'>&nbsp;" & _
					tblNme("Index") & "&nbsp;</font></a></td><td>" & _
					"<font size='2'>&nbsp;" & tmpName & "&nbsp;</font></td></tr>"
			tblNme.Close
			tblCM.MoveNext
			ctr = ctr + 1
		Loop
	Else
		strCM = "<tr><td align='center'><font size='2'>N/A</font></td></tr>"
	End If
	tblCM.Close
	Set tblNme = Nothing
	Set tblCM = Nothing	
	'''''''''''''REPRESENTATIVE'''''''
	
	Set tblRep = Server.CreateObject("ADODB.RecordSet")
	Set tblNme = Server.CreateObject("ADODB.RecordSet")
	
	sqlRep = "SELECT * FROM [ConRep_t] WHERE [CID] = '" & Request("ID") & "' " 
	
	tblRep.Open sqlRep, g_strCONN, 1, 3
	If Not tblRep.EOF Then
		ctr = 0
		Do Until tblRep.EOF
			sqlName = "SELECT * FROM [Representative_t] WHERE [Index] = " & tblRep("RID") 
			tblNme.Open sqlName, g_strCONN, 1, 3
			tmpName = tblNme("Lname") & ", " & tblNme("Fname")
			strRep = strRep & "<tr><td align='center'><input type='checkbox' name='chkR" & ctr & "' value='" & _
					tblNme("Index") & "'<td><a href='A_Rep.asp?RID=" & tblNme("Index") & "'><font size='2' face='trebuchet MS'>&nbsp;" & _
					tblNme("Index") & "&nbsp;</font></a></td><td>" & _
					"<font size='2'>&nbsp;" & tmpName & "&nbsp;</font></td></tr>"
			tblNme.Close
			tblRep.MoveNext
			ctr = ctr + 1
		Loop
	Else
		strRep = "<tr><td align='center'><font size='2'>N/A</font></td></tr>"
	End If
	tblRep.Close
	Set tblNme = Nothing
	Set tblRep = Nothing
	
	''''''''''''''''WORKER DROPDOWN
	Set tblLWork = Server.CreateObject("ADODB.Recordset")
strSQLd = "SELECT * FROM [Worker_t] ORDER BY [lname]"
'On Error Resume Next
tblLWork.Open strSQLd, g_strCONN, 3, 1
	tblLWork.Movefirst
	Do Until tblLWork.EOF
		Set tblChkCon = Server.CreateObject("ADODB.Recordset")
		sqlChkCon = "SELECT * FROM ConWork_t WHERE CID = '" & Request("ID") & "' "
		tblChkCon.Open sqlChkCon, g_strCONN, 1, 3
		If Not tblChkCon.EOF Then
			meron = 0
			Do Until tblChkCon.EOF
				If tblLWork("index") = Cint(tblChkCon("WID")) Then 
					'response.write "SQL" & sqlChkCon & " TRUE" 
					meron = 1
				End If
				
				'response.write "SQL: " & sqlChkCon & " FALSE: " & tblLWork("index") & " <> " & tblChkCon("WID") & "<BR>"
				tblChkCon.MoveNext
			Loop
			If meron <> 1 Then strdept = strdept & "<option value='" & tblLWork("index")& "'> "& tblLWork("Lname") & ", " & tblLWork("fname") & " </option>"
		Else
			strdept = strdept & "<option value='" & tblLWork("index")& "'> "& tblLWork("Lname") & ", " & tblLWork("fname") & " </option>"  
			'strdept = strdept & "<option value='" & tbldept("lname")& "'> "& tbldept("Lname") & ", " & tbldept("fname") & " </option>"  
		End If
		tblLWork.Movenext
	loop

tblLWork.Close
set tblLWork = Nothing

	''''''''''''''''CASE MANAGER DROPDOWN
	Set tblCMCon = Server.CreateObject("ADODB.Recordset")
	Set tblCM2 = Server.CreateObject("ADODB.Recordset")
	
	sqlCMCon = "SELECT * FROM [CMCon_t] WHERE CID = '" & Request("ID") & "' "
	sqlCM2 = "SELECT * FROM [Case_Manager_t] ORDER BY [Lname]"
	
	tblCMCon.Open sqlCMCon, g_strCONN, 1, 3
	tblCM2.Open sqlCM2, g_strCONN, 1, 3
	
	If tblCMCon.EOF Then
		CMLocked = ""
		Do Until tblCM2.EOF
			strCM2 = strCM2 & "<option value='" & tblCM2("index")& "'> "& tblCM2("Lname") & ", " & tblCM2("fname") & " </option>"
			tblCM2.MoveNext
		loop
	Else
		CMLocked = "DISABLED"
	End If
tblCM2.close	
tblCMCon.Close
set tblCMCon = Nothing
Set tblCM2 = Nothing

	''''''''''''''''REPRESENTATIVE DROPDOWN
	Set tblRCon = Server.CreateObject("ADODB.Recordset")
	Set tblR2 = Server.CreateObject("ADODB.Recordset")
	
	sqlRCon = "SELECT * FROM [ConRep_t] WHERE CID = '" & Request("ID") & "' "
	sqlR2 = "SELECT * FROM [Representative_t] ORDER BY [Lname]"
	
	tblRCon.Open sqlRCon, g_strCONN, 1, 3
	tblR2.Open sqlR2, g_strCONN, 1, 3
	
	If tblRCon.EOF Then
		RLocked = ""
		Do Until tblR2.EOF
			strR2 = strR2 & "<option value='" & tblR2("Index")& "'> "& tblR2("Lname") & ", " & tblR2("Fname") & " </option>"
			tblR2.MoveNext
		loop
	Else
		RLocked = "DISABLED"
	End If
tblR2.close	
tblRCon.Close
set tblRCon = Nothing
Set tblR2 = Nothing
%>
<html>
	<head>
		<title>LSS - Consumer Details - List</title>
		<script language='JavaScript'>
			function DelList()
			{
				document.frmRepList.action = "ConList.asp?Del=1";
				document.frmRepList.submit();
			}
		</script>
	</head>
	<body bgcolor='white' LEFTMARGIN='0' TOPMARGIN='0'>
		<center>
		<form method='post' name='frmRepList' action='ConList.asp'>
		<!-- #include file="_boxup.asp" -->
		<!-- #include file="_NavHeader.asp" -->
		<center>
		<br>
		<table border='0'>
					<tr>
						<td colspan='4' align='center' width='600px'>
							<font size='3' face='trebuchet MS'><b><u>Consumer Details - List</u></b></font>
		<a href='A_Consumer.asp?MC=<%=Request("MNum")%>' style='text-decoration:none'><font size='2' face='trebuchet MS'>[Details]</font></a>
							<a href='A_C_Status.asp?MNum=<%=MCNum%>' style='text-decoration:none'><font size='2' face='trebuchet MS'>[Status]</font></a>
							<a href='A_C_Health.asp?MNum=<%=Request("MNum")%>' style='text-decoration:none'><font size='2' face='trebuchet MS'>[Health]</font></a>
							<a href='A_C_Files.asp?MNum=<%=Request("MNum")%>' style='text-decoration:none'><font size='2' face='trebuchet MS'>[Files]</font></a>
							<a href='Log.asp' style='text-decoration:none'><font size='2' face='trebuchet MS'>[Log]</font></a>
							<font size='2' face='trebuchet MS'>[List]</font>
						</td>
					</tr>
					<tr><td colspan='2' align='center'><font color='red' face='trebuchet MS' size='2'>&nbsp;<%=Session("MSG")%>&nbsp;</font></td></tr>
					<tr>
						<td>
							<font size='2' face='trebuchet MS'><u>Name:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 120px;' readonly name='Cname' value="<%=Session("Cname")%>"></font>
						</td>
					</tr>
				</table>
		<br><br>
		<table border='0'>
		<tr>
			<td valign='top'>
			<table border='1'>
				<tr><td colspan='3' align='center'><font face='trebuchet MS'><b>PCS Worker List</b></font></td></tr>
				<%=strWork%>
				
			</table>
			</td>
			<td valign='top'>
			<table border='1'>
				<tr><td colspan='3' align='center'><font face='trebuchet MS'><b>Case Manager List</b></font></td></tr>
				<%=strCM%>
				
			</table>
			</td>
			<td valign='top'>
				<table border='1'>
				<tr><td colspan='3' align='center'><font face='trebuchet MS'><b>Representative List</b></font></td></tr>
				<%=strRep%>
				</table>
			</td>
		</tr>
		<tr>
			<td><select style='font-size: 10px; height: 20px; width: 100%;' name='SelWor'>
				<option></option>
				<%=strdept%></select></td>
			<td><select style='font-size: 10px; height: 20px; width: 100%;' <%=CMLocked%> name='SelCM'><option></option><%=strCM2%></select></td>
			<td><select style='font-size: 10px; height: 20px; width: 100%;' <%=RLocked%> name='SelR'><option></option><%=strR2%></select></td>
		</tr>
		<tr><td>&nbsp;</td></tr>
		<tr>
			<td align='center' colspan='3'>
				<input type='submit' value='Save'>
				<input type='button' value='Delete' onclick='JavaScript: DelList();'>
				<input type='hidden' name='tmpID' value='<%=Request("ID")%>'>
				<input type='hidden' name='ctr' value='<%=ctrW%>'>
			</td>
		</tr>
		</table>
		<!-- #include file="_Boxdown.asp" -->
		</form>
	</body>
</html>
<%
	Session("MSG") = ""
%>