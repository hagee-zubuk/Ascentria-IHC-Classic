<%Language=VBScript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<%
Function GetWIndex(xxx)
	Set rsI = Server.CreateObject("ADODB.RecordSet")
	sqlI = "SELECT [index] FROM Worker_T WHERE Social_Security_Number = '" & xxx & "'"
	rsI.Open sqlI, g_strCONN, 3, 1
	If Not rsI.EOF Then
		GetWIndex = rsI("Index")
	End If
	rsI.Close
	Set rsI = Nothing
End Function
	If Request("WID") <> "" Then
			tmpWID = GetWIndex(Request("WID"))
			Set tblCon = Server.CreateObject("ADODB.RecordSet")
			Set tblName = Server.CreateObject("ADODB.RecordSet")
			
			sqlCon = "SELECT * FROM [ConWork_t] WHERE [WID] = '" & tmpWID & "' " 
			'response.write Session("WID")
			tblCon.Open sqlCon, g_strCONN, 1, 3
			If Not tblCon.EOF Then
				ctr = 0
				Do Until tblCon.EOF
					if Z_IsOdd(ctr) = true then 
						kulay = "#FFFAF0" 
					else 
						kulay = "#FFFFFF"
					end if
					sqlName = "SELECT * FROM [Consumer_t] WHERE [Medicaid_Number] = '" & tblCon("CID") & "' "
					'response.write sqlname
					tblName.Open sqlName, g_strCONN, 1, 3
					if not tblName.EOF Then tmpName = tblName("Lname") & ", " & tblName("Fname")
					tblName.Close
					strCon = strCon & "<tr bgcolor='" & kulay & "'><td align='center'><input type='checkbox' name='chk" & ctr & _
						"' value='" & tblCon("ID") & "'></td><td><a href='A_Consumer.asp?MNum=" & tblCon("CID") & "'><font size='1'>" & tmpName & _
						"</font>&nbsp;</a></td></tr>"
					tblCon.MoveNext
					ctr = ctr + 1
				Loop
			Else
				strCon = "<tr><td align='center'>N/A</td></tr>"
			End If
			tblCon.Close
			Set tblName = Nothing
			Set tblCon = Nothing	
			''''''''''''''''''''''CONSUMER DROPDOWN
			Set tblLWork = Server.CreateObject("ADODB.Recordset")
			strSQLd = "SELECT * FROM [Consumer_t] ORDER BY [lname]"
			'On Error Resume Next
			tblLWork.Open strSQLd, g_strCONN, 3, 1
			tblLWork.Movefirst
			Do Until tblLWork.EOF
				Set tblChkCon = Server.CreateObject("ADODB.Recordset")
				sqlChkCon = "SELECT * FROM ConWork_t WHERE WID = '" & tmpWID & "' "
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
	Else
		Session("MSG") = "Sorry. To maintain integrity of database, please Sign-in again."
		Response.Redirect "default.asp"
	End If
%>
<html>
	<head>
		<title>LSS - In-Home Care - PCSP Worker Details - List</title>
		<script language='JavaScript'>
			function DelLink()
			{
				document.frmRepList.action='A_W_Del.asp?page=1';
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
		<form method='post' name='frmRepList' action='A_W_Action.asp?page=3'>
			<!-- #include file="_boxup.asp" -->
		<!-- #include file="_NavHeader.asp" -->
		<center>
		<br>
<table border='0'>
			<tr>
				
				<td colspan='8' align='center' >
		<font size='2' face='trebuchet MS'><b><u>PCSP Worker Details - List</u></b></font>
							<a href='A_Worker.asp?WID=<%=Request("WID")%>' style='text-decoration:none'><font size='1' face='trebuchet MS'>[Details]</font>
							<a href='A_W_Files.asp?WID=<%=Request("WID")%>' style='text-decoration:none'><font size='1' face='trebuchet MS'>[Files]</font></a>
							<a href="A_W_Skills.asp?WID=<%=Request("WID")%>" style='text-decoration: none;'>
								<font color='blue' size='1' face='trebuchet MS'>[Skills]</font>
							</a>
							<font size='2' face='trebuchet MS'>[List]</font>
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
	<tr><td>&nbsp;</td></tr>
	<tr><td>&nbsp;</td></tr>
			<tr>
				
				<td><font size='1' face='trebuchet MS'><u>Name:</u></td>
				<td><input style='font-size: 10px; height: 20px; width: 120px;' readonly value="<%=Session("Wname")%>"></td>
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
			<tr bgcolor='#040C8B'><td colspan='3' align='center' style="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=0, startColorstr=#040C8B, endColorstr=#c7c9e5);">
				<font face='trebuchet MS' font size='2' color='white'><b>Consumer List</b></font></td></tr>
			<%=strCon%>
			</tr>
			<tr>
				<td  align='center' colspan='3'>
				<Select name='SelCon'> 
					<option></option>
					<%=strdept%>
				</select>
			</td>
			</tr>
			<tr>
				<td align='center' colspan='3'>
					<input type='hidden' name='WID' value="<%=Request("WID")%>">
					<input type='hidden' name='ctr' value='<%=ctr%>'>
					
					<input type='button' value='Save List' class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='document.frmRepList.submit();'>
					<input type='button' value='Delete Checked List' class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='JavaScript: DelLink();'>
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