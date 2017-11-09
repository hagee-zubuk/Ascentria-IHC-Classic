<%Language=VBScript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<%
	Set tblCon = Server.CreateObject("ADODB.RecordSet")
	'Set tblName = Server.CreateObject("ADODB.RecordSet")
	
	sqlCon = "SELECT * FROM [ConRep_t], consumer_T WHERE [RID] = '" & Request("ID") & "' AND CID = medicaid_number ORDER BY lname, fname" 
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
			'sqlName = "SELECT * FROM [Consumer_t] WHERE [Medicaid_Number] = '" & tblCon("CID") & "' "
			'tblName.Open sqlName, g_strCONN, 1, 3
			tmpName = tblcon("Lname") & ", " & tblcon("Fname")
			'tblName.Close
			strCon = strCon & "<tr bgcolor='" & kulay & "'><td align='center'><input type='checkbox' name='chk" & ctr & _
				"' value='" & tblCon("ID") & "'></td><td><a href='A_Consumer.asp?Mnum=" & tblCon("CID") & "'>&nbsp;<font size='1'>" & tmpName & _
				"</font>&nbsp;</a></td></tr>"
			tblCon.MoveNext
			ctr = ctr + 1
		Loop
	Else
		strCon = "<tr><td align='center'>N/A</td></tr>"
	End If
	tblCon.Close
	'Set tblName = Nothing
	Set tblCon = Nothing	
	''''''''''''''''''''''CONSUMER DROPDOWN
	Set tblLWork = Server.CreateObject("ADODB.Recordset")
	strSQLd = "SELECT * FROM [Consumer_t] ORDER BY [lname]"
	'On Error Resume Next
	tblLWork.Open strSQLd, g_strCONN, 3, 1
	tblLWork.Movefirst
	Do Until tblLWork.EOF
		Set tblChkCon = Server.CreateObject("ADODB.Recordset")
		sqlChkCon = "SELECT * FROM ConRep_t WHERE RID = '" & Request("ID") & "' "
		tblChkCon.Open sqlChkCon, g_strCONN, 1, 3
		If Not tblChkCon.EOF Then
			meron = 0
			Do Until tblChkCon.EOF
				If tblLWork("Medicaid_Number") = tblChkCon("CID") Then 
					'response.write "SQL" & sqlChkCon & " TRUE" 
					meron = 1
				End If
				
				'response.write "SQL: " & sqlChkCon & " FALSE: " & tblLWork("index") & " <> " & tblChkCon("WID") & "<BR>"
				tblChkCon.MoveNext
			Loop
			If meron <> 1 Then strdept = strdept & "<option value='" & tblLWork("Medicaid_Number")& "'> "& tblLWork("Lname") & ", " & tblLWork("fname") & " </option>"
		Else
			strdept = strdept & "<option value='" & tblLWork("Medicaid_Number")& "'> "& tblLWork("Lname") & ", " & tblLWork("fname") & " </option>"  
			'strdept = strdept & "<option value='" & tbldept("lname")& "'> "& tbldept("Lname") & ", " & tbldept("fname") & " </option>"  
		End If
		tblLWork.Movenext
	loop

tblLWork.Close
set tblLWork = Nothing
%>
<html>
	<head>
		<title>LSS - In-Home Care - Representative Details - List</title>
		<script language='JavaScript'>
			function DelLink()
			{
				document.frmRepList.action='A_R_Action.asp?page=2';
				document.frmRepList.submit();
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
		<center>
		<form method='post' name='frmRepList' action='A_R_Action.asp?page=1'>
			<!-- #include file="_boxup.asp" -->
		<!-- #include file="_NavHeader.asp" -->
		<center>
		<br>
		<table border='0'>
			<tr>
				
				<td colspan='8' align='center'>
		<font size='2' face='trebuchet MS'><b><u>Representative - List</u></b></font>
		&nbsp;
		<a href='A_Rep.asp' style='text-decoration: none;'>
			<font size='1' face='trebuchet MS' color='blue'>[Details]</font></a>
				<font size='2' face='trebuchet MS'>[List]</font>
		</td>
	</tr>
	<tr><td>&nbsp;</td></tr>
	<tr><td>&nbsp;</td></tr>
			<tr>
				
				<td><font size='1' face='trebuchet MS'><u>Name:</u></td>
				<td><input style='font-size: 10px; height: 20px; width: 120px;' readonly value='<%=Session("Rname")%>'></td>
				<tr><td colspan='6' align='center'><font face='trebuchet MS' color='red' size='1'>&nbsp;<%=Session("MSG")%>&nbsp;</font></td></tr>
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
		
		
		<table border='1'>
			<tr bgcolor='#040C8B'><td colspan='3' align='center' style="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=0, startColorstr=#040C8B, endColorstr=#c7c9e5);">
				<font face='trebuchet MS' size='2' color='white'><b>Consumer List</b></font></td></tr>
			<%=strCon%>
			</tr>
			<tr>
				<td align='center' colspan='3'>
				<Select name='SelCon'> 
					<option></option>
					<%=strdept%>
				</select>
			</td>
			</tr>
			<tr>
				<td align='center' colspan='3'>
					<input type='hidden' name='ctr' value='<%=ctr%>'>
					<input type='hidden' name='rid' value='<%=Request("ID")%>'>
					<input type='button' class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" value='Save List' onclick='document.frmRepList.submit();'>
					<input type='button' class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" value='Delete Checked List' onclick='JavaScript: DelLink();'>
				</td>
			</tr>
		</table>
		</form>
		<!-- #include file="_boxdown.asp" -->
	</body>
</html>
<%
	Session("MSG") = ""
%>