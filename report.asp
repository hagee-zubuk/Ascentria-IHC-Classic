<%language=vbscript%>
<!-- #include file="_Files.asp" -->

<!-- #include file="_Utils.asp" -->
<!-- #include file="_security.asp" -->
<%
	DIM empSQL, tblEMP, tblConsumer, cliSQL
	Session("PrintPrev") = ""
Session("PrintPrevRep") = ""
Session("PrintPrevPRoc") = ""
	Set tblEMP = Server.CreateObject("ADODB.RecordSet")
	empSQL = "Select * From worker_t order By lname"
	tblEMP.Open empSQL, g_strCONN, 3, 1
	If Not tblEMP.EOF Then
		Do While Not tblEMP.EOF
			stremp = stremp & "<option>" & tblEMP("Social_Security_Number") & " - " & tblEMP("lname") & ", " & tblEMP("fname") & "</option>"
			tblEMP.MoveNext
		Loop
	End If
	tblEMP.Close
	set tblEMP = Nothing
	
	Set tblConsumer = Server.CreateObject("ADODB.RecordSet")
	cliSQL = "Select * From consumer_t Order By lname"
	tblConsumer.Open cliSQL, g_strCONN, 3, 1
	If Not tblConsumer.EOF Then
		Do While Not tblConsumer.EOF
			If len(tblConsumer("Medicaid_Number")) = 1 Then
				tmpIn = "0" & tblConsumer("Medicaid_Number")
			Else
				tmpIn = tblConsumer("Medicaid_Number")
			End If
			strConsumer = strConsumer & "<option>" & tmpIn & " - " & tblConsumer("lname") & ", " & tblConsumer("fname") & "</option>"
			tblConsumer.MoveNext
		Loop
	End If
	tblConsumer.Close
	set tblConsumer = Nothing

	If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
		server.scripttimeout = 360000
		remi = 0 
		if Request("EMP")  = "For each PCSP Worker" Then remi = 1
		
		Set tblREP = Server.CreateObject("ADODB.RecordSet")
		sqlREP = "SELECT Worker_t.*, Consumer_t.*, tsheets_t.* FROM Worker_t, tsheets_t, Consumer_t WHERE emp_id = Worker_t.social_security_number " & _
			"AND consumer_t.medicaid_number = tsheets_t.Client "
			
		If Request("Emp") = "" And Request("Consumer") = "" Then 
			Session("MSG") = "More information needed."
			Response.Redirect "Report.asp" 
		End If
			
		If Request("EMP") = "For each PCSP Worker" And Request("Consumer") = "For each Consumer" Then
			Session("MSG") = "Invalid Criteria."
			Response.Redirect "Report.asp"
		End If
		If Request("TF1") = "" And Request("TF2") <> "" Then
			Session("MSG") = "Invalid Criteria."
			Response.Redirect "Report.asp"
		End If
		If Request("TF1") <> "" Then
			If Not IsDate(Request("TF1")) Then 
				Session("MSG") = "Invalid Date"
				Response.Redirect "Report.asp"
			End If
			If Request("TF2") <> "" Then
				If Not IsDate(Request("TF2")) Then 
					Session("MSG") = "Invalid Date"
					Response.Redirect "Report.asp"
				ElseIf CDate(Request("TF1")) > CDate(Request("TF2")) Then
					Session("MSG") = "Invalid Date"
					Response.Redirect "Report.asp"
				End If
			End If
		End If
		If Request("Emp") <> "" Then
			If Request("Emp") <> "For each PCSP Worker" Then 
				tmpEMP = split(Request("Emp")," - ")
				sqlREP = sqlREP & "AND emp_id = '" & tmpEMP(0) & "' "
			End If
		End If 
		If Request("Consumer") <> "" Then
			If Request("Consumer") <> "For each Consumer" Then
				tmpCli = split(Request("Consumer")," - ")
				sqlREP = sqlREP & "AND Consumer_t.medicaid_number = '" & tmpCli(0) & "' "
			End If
		End If
		If Request("TF1") <> "" Then
			myDATE = CDate(Request("TF1"))
			If WeekdayName(Weekday(myDATE), true) = "Sun" Then
				D1 = myDATE
			Else
				difDATE = DatePart("w", myDATE)
				D1 = DateAdd("d", -Cint(difDATE - 1), myDATE)
			End If
		End If
		If Request("TF2") <> "" Then
			myDATE = CDate(Request("TF2"))
			If WeekdayName(Weekday(myDATE), true) = "Sun" Then
				D2 = CDate(myDATE) + 6
			Else
				difDATE = DatePart("w", myDATE)
				D2 = DateAdd("d", Cint(difDATE - 1), myDATE)
			End If
		End If
		If Request("TF1") <> "" Then
			If IsDate(Request("TF1")) Then sqlREP = sqlREP & " AND date >= '" & d1 & "' "
			D1 = Request("TF1")
			If Request("TF2") <> "" Then
				If IsDate(Request("TF2")) Then sqlREP = sqlREP & " AND date <= '" & d2 & "' "
				D2 = Request("TF2")	
			End If
		End If
		If Not (Request("chkEXT") <> "") Then sqlREP = sqlREP & " AND EXT = 0 " 
		If Request("EMP") = "For each PCSP Worker" Or (Request("EMP") = "" And (Request("Consumer") <> "" And Request("Consumer") <> "For each Consumer")) Then _
			sqlREP = sqlREP & " ORDER BY worker_t.lname"
		If Request("Consumer") = "For each Consumer" Or ((Request("EMP") <> "" And Request("EMP") <> "For each PCSP Worker") And Request("Consumer") = "") Then _
			sqlREP = sqlREP & "ORDER BY consumer_t.lname"	
		If Request("TF1") <> "" Then
			myDATE = CDate(D1)
			If WeekdayName(Weekday(myDATE), true) = "Sun" Then
				D1 = myDATE
			Else
				difDATE = DatePart("w", myDATE)
				D1 = DateAdd("d", -Cint(difDATE - 1), myDATE)
			End If
		End If
		If Request("TF2") <> "" Then
			myDATE = CDate(D2)
			If WeekdayName(Weekday(myDATE), true) = "Sun" Then
				D2 = CDate(myDATE) + 6
			Else
				difDATE = DatePart("w", myDATE)
				D2 = DateAdd("d", Cint(difDATE - 1), myDATE)
			End If
		End If
		Wk = D1 & " - " & D2	
		Session("Wk") = Wk
		If Request("EMP") = "For each PCSP Worker" Then Session("type") = 1
		If Request("Consumer") = "For each Consumer" Then Session("type") = 2
		'response. write "<br>" & sqlREP & "<br>"
		tblREP.Open sqlREP, g_strCONN, 1, 3
		tmpID2 = ""
		marker = 0
		If Not tblREP.EOF Then
		SHours = 0
		tblREP.MoveFirst
			Do
				If Request("EMP") = "For each PCSP Worker" Then
					tmpID = tblREP("emp_id")
					If tmpID <> tmpID2 Or tblREP.EOF Then
						THours = (Tmon) + (Ttue) + (Twed) + (Tthur) + (Tfri) + (Tsat) + (Tsun)
						Tmon = 0
						Ttue = 0
						Twed = 0
						Tthur = 0
						Tfri = 0
						Tsat = 0
						Tsun = 0
						If tmpID2 <> "" Then 
							Set tblname = Server.CreateObject("ADODB.RecordSet")
							sqlName = "SELECT * FROM Worker_t"
							tblname.Open sqlName, g_strCONN, 1, 3
							strTmp = "Social_Security_Number= '" & tmpID2 & "' "
							tblname.Find(strTmp)
							If Not tblname.EOF Then name = tblname("lname") & ", "	& tblname("fname")
							tblname.Close
							set tblname = Nothing	
									strTBL = strTBL & "<tr><td align='center'>" & Wk & "</td><td align='center'>" & Right(tmpID2, 4) & _
										"</td><td align='center'>" & _
										name & "</td><td align='center'>" & _
										Z_FormatNumber(THours,2) & "</td></tr>" & vbCrLf
										SHours = Shours + THours
						End If
						tmpID2 = tmpID
					End If
					Tmon = (Tmon) + (tblREP("mon"))
					Ttue = (Ttue) + (tblREP("tue"))
					Twed = (Twed) + (tblREP("wed"))
					Tthur = (Tthur) + (tblREP("thu"))
					Tfri = (Tfri) + (tblREP("fri"))
					Tsat = (Tsat) + (tblREP("sat"))
					Tsun = (Tsun) + (tblREP("sun"))
				ElseIf Request("Consumer") = "For each Consumer" Then
					'tmpID2 = "0"
					tmpID = tblREP("medicaid_number")
					If tmpID <> tmpID2 Or tblREP.EOF Then
						THours = (Tmon) + (Ttue) + (Twed) + (Tthur) + (Tfri) + (Tsat) + (Tsun)
						Tmon = 0
						Ttue = 0
						Twed = 0
						Tthur = 0
						Tfri = 0
						Tsat = 0
						Tsun = 0
						'response.write "ID: " & tmpID2 & "<br>"
						'if tmpID2 = "" Then tmpID2 = "0"
						If tmpID2 <> "" Then 
							'If THours <> 0 Then
								Set tblname = Server.CreateObject("ADODB.RecordSet")
								sqlName = "SELECT * FROM Consumer_t"
								tblname.Open sqlName, g_strCONN, 1, 3
								strTmp = "medicaid_number= '" & tmpID2 & "' "
								tblname.Find(strTmp)
								If Not tblname.EOF Then name = tblname("lname") & ", "	& tblname("fname")
								tblname.Close
								set tblname = Nothing
										strTBL = strTBL & "<tr><td align='center'>" & wk & "</td><td align='center'>" & tmpID2 & _
										"</td><td align='center'>" & _
											name & "</td><td align='center'>" & _
											Z_FormatNumber(THours,2) & "</td></tr>" & vbCrLf
								SHours = Shours + THours
							'End If
						End If
						tmpID2 = tmpID
					End If
					Tmon = (Tmon) + (tblREP("mon"))
					Ttue = (Ttue) + (tblREP("tue"))
					Twed = (Twed) + (tblREP("wed"))
					Tthur = (Tthur) + (tblREP("thu"))
					Tfri = (Tfri) + (tblREP("fri"))
					Tsat = (Tsat) + (tblREP("sat"))
					Tsun = (Tsun) + (tblREP("sun"))
				Else
					Tmon = (Tmon) + (tblREP("mon"))
					Ttue = (Ttue) + (tblREP("tue"))
					Twed = (Twed) + (tblREP("wed"))
					Tthur = (Tthur) + (tblREP("thu"))
					Tfri = (Tfri) + (tblREP("fri"))
					Tsat = (Tsat) + (tblREP("sat"))
					Tsun = (Tsun) + (tblREP("sun"))
				End If
				tblREP.MoveNext
			Loop Until tblREP.EOF
			THours = (Tmon) + (Ttue) + (Twed) + (Tthur) + (Tfri) + (Tsat) + (Tsun)
			Tmon = 0
			Ttue = 0
			Twed = 0
			Tthur = 0
			Tfri = 0
			Tsat = 0
			Tsun = 0
			
			If Request("EMP") = "For each PCSP Worker" Then 
				Set tblname = Server.CreateObject("ADODB.RecordSet")
							sqlName = "SELECT * FROM Worker_t"
							tblname.Open sqlName, g_strCONN, 1, 3
							strTmp = "Social_Security_Number= '" & tmpID2 & "' "
							tblname.Find(strTmp)
							If Not tblname.EOF Then name = tblname("lname") & ", "	& tblname("fname")
							tblname.Close
							set tblname = Nothing
					strTBL = strTBL & "<tr><td align='center'>" & wk & "</td><td align='center'>" & right(tmpID2, 4) & _
								"</td><td align='center'>" & _
								name & "</td><td align='center'>" & _
								Z_FormatNumber(THours,2) & "</td></tr>" & vbCrLf
								SHours = Shours + THours
				ElseIf Request("Consumer") = "For each Consumer" Then
					Set tblname = Server.CreateObject("ADODB.RecordSet")
							sqlName = "SELECT * FROM Consumer_t"
							tblname.Open sqlName, g_strCONN, 1, 3
							strTmp = "medicaid_number= '" & tmpID & "' "
							tblname.Find(strTmp)
							If Not tblname.EOF Then name = tblname("lname") & ", "	& tblname("fname")
							tblname.Close
							set tblname = Nothing
					strTBL = strTBL & "<tr><td align='center'>" & wk & "</td><td align='center'>" & tmpID & _
							"<td align='center'>" & name & "</td><td align='center'>" & _
								Z_FormatNumber(THours,2) & "</td></tr>" & vbCrLf
								SHours = Shours + THours
				'End If
			Else
				marker = 1
			End If
			
		End If
		tblREP.Close
		set tblREP = Nothing	
		If Thours <> "" Then
			If Request("EMP") <> "" Then 
				If Request("EMP") <> "For each PCSP Worker" And Request("Consumer") <> "For each Consumer" Then
					strHEAD = strHEAD & "<td align='center' width='200px'>Week</td><td align='center' width='200px'>SNN</td><td align='center' width='200px'>PCSP Worker</td>"
					tmpEMP = split(Request("Emp")," - ")
					strTBL = strTBL & "<td align='center'>" & Wk & "</td><td align='center'>" & Right(tmpEMP(0), 4) & "</td><td align='center'>" & tmpEMP(1) & "</td>"
				ElseIf Request("EMP") = "For each PCSP Worker" Then
					strHEAD = strHEAD & "<td align='center' width='200px'>Week</td><td align='center' width='200px'>SSN" & _
						"</td><td align='center' width='200px'>PCSP Worker</td>"
				End If
			End If 
			If Request("Consumer") <> "" Then 
				If Request("Consumer") <> "For each Consumer" And Request("EMP") <> "For each PCSP Worker" Then 
					strHEAD = strHEAD & "<td align='center' width='200px'>Week</td><td align='center' width='200px'>Medicaid #</td><td align='center' width='200px'>Consumer</td>"
					tmpCLI = split(Request("Consumer")," - ")
					strTBL = strTBL & "<td align='center'>" & Wk & "</td><td align='center'>" & tmpCLI(0) & "</td><td align='center'>" & tmpCLI(1) & "</td>"
				ElseIf Request("Consumer") = "For each Consumer" Then
					strHEAD = strHEAD & "<td align='center' width='200px'>Week</td><td align='center' width='200px'>Medicaid #" & _
						"</td><td align='center' width='200px'>Consumer</td>"
				End If
			End If	 
				strHEAD = strHEAD & "<td align='center' width='100px'>Hours</td>"
				If Marker = 1 Then strTBL =strTBL & "<td align='center'>" & Z_FormatNumber(THours,2) & "</td>"
				If SHours <> 0 Then	strTBL = strTBL & "<tr><td align='right' colspan='3'><b>Total Hours:&nbsp;</b></td><td align='center'><b>" & _
				Z_FormatNumber(SHours,2) & "</b></td><tr>"
				Session("strTableScript") = "<table cellSpacing='0' cellPadding='0' align='center' border='1'>" & _
						"<tr>" & strHEAD & "</tr>" & strTBL & "</table>"
				'Generate csv file
				If (Request("Emp") <> "") Or (Request("Emp") <> "For each PCSP Worker") Then
					'response.write Request("Emp") & vbcrlf
					'response.write remi & vbcrlf
					if (Request("Consumer") <> "For each Consumer") Then
					'
					'	response.write Request("Consumer")
					On Error Resume Next
						temp = Z_DoEncrypt(tmpEMP(1))
						If temp <> "" Then
							tConsumer = Z_DoEncrypt(Request("Consumer")) 
							'Session("strButt") = "<center><a href=""export.asp?exp=1 &thrs=" & THours & "&emp=" & temp & _
							'		"&cli=" & tConsumer & """><font color='red' face='trebuchet MS' size='2'>Export Record for Payroll</font></a></center>" 
							Session("strButt") = "<input type='button' disabled onclick='document.location=""export.asp?exp=1 &thrs=" & THours & "&emp=" & temp & _
									"&cli=" & tConsumer & """' value='Export Record for Payroll'>"
						End If
					End if
					remi = 0
				End If
				LS = 0
			 	If Request("Consumer") <> "" And Request("EMP") <> "For each PCSP Worker" And Request("EMP") <> "" And _
			 		Request("Consumer") <> "For each Consumer" Then LS = 1
			 	 
				strSUM = "Total hours "
				If Request("Consumer") = "" And Request("EMP") = "For each PCSP Worker" Then strSUM = strSUM & "for each PCSP Worker"
				If Request("EMP") = "" And Request("Consumer") = "For each Consumer" Then strSUM = strSUM & "for each Consumer"
				If Request("EMP") <> "" And Not Request("EMP") = "For each PCSP Worker" Then strSUM = strSUM & "of PCSP Worker " & Request("EMP")
				If Request("Consumer") <> "" And Not Request("Consumer") = "For each Consumer" Then strSUM = strSUM & " of Consumer " & Request("Consumer")
				If Request("EMP") = "For Each PCSP Worker" Then strSUM = strSUM & "of PCSP Worker in Consumer" & Request("Consumer")
				If Request("Consumer") = "For Each Consumer" Then strSUM = strSUM & "of " & Request("EMP") & " for each Consumer"
				If Request("TF1") <> "" Then 
					strSUM = strSUM & " from " & d1
					If Request("TF2") <> "" Then strSUM = strSUM & " to " & d2
				End If
				If (Request("chkEXT") <> "")  then  strSUM = strSUM & "<br> with extended hours."
				Session("MSG") = strSUM 
			Else
				Session("MSG") = "No Records Found."
			End If
			'Response.Redirect "report.asp"
			Session("tmpSQL") = Z_DoEncrypt(sqlRep)
			Session("PrintPrev") = ""
			Session("PrintPrevPRoc") = ""
			Session("PrintPrevPRoc") = ""
			Session("PrintPrevRep") = Session("strTableScript") & "|" & Session("MSG")
	End If	
%>		
<html>
	<head>
		<title>Timesheet - Report</title>
		<link href="styles.css" type="text/css" rel="stylesheet" media="print">
		<script language='JavaScript'>
			function ExCSV()
			{
				document.frmSort.action = "Export.asp?sql=1";
				document.frmSort.submit();
			}
			function LandWarn()
			{
				var ans = window.confirm("Please set page orientation to landscape. Click Ok to continue. Click Cancel to stop.");
				if (ans){
				document.frmSort.action = "print.asp";
				document.frmSort.submit();
				}
			}
			function PrintPrev()
			{
				document.frmSort.action = "print.asp";
				document.frmSort.submit();
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
	<body bgcolor='white'  LEFTMARGIN='0' TOPMARGIN='0'>
		<!-- #include file="_boxup.asp" -->
		<!-- #include file="_NavHeader.asp" -->
		
		
		<form method='post' name='frmSort' action="report.asp">
			
	<br><br>
			<table cellSpacing='0' cellPadding='0' align='center' border='0'>
				<tr><td colspan='4' align='center'>
				<font size='2' face='trebuchet MS'>[General Reports]</font>&nbsp;&nbsp;
				<a href='SpecRep.asp' style='text-decoration: none;'><font size='1' color='blue' face='trebuchet MS'>[Advance Reports]&nbsp;</font></a>
				</td></tr>
				<tr
					><td></td
				></tr>
				<tr
					><td><font face='trebuchet MS' size='1'><u>PCSP Worker:</u>&nbsp;</font></td
					><td><select style='width: 250px;' name='emp' align='left'>
						<option></option>
						<%=stremp%>
						<option>For each PCSP Worker</option>
						</select></td 
					><td><font face='trebuchet MS' size='1'>&nbsp;<u>Consumer:</u>&nbsp;</font></td
					><td><select style='width: 250px;' name='Consumer' align='center'>
						<option></option>
						<%=strConsumer%>
						<option>For each Consumer</option>
						</select></td 
				><td>
					<font face='trebuchet MS' size='1'>Extended Hours:</font>
					<input type='checkbox' name='chkEXT'>
				</td>
				
				</tr
				><tr
					><td>&nbsp;</td
				></tr
				><tr align='right'
					><td align='center' colspan='7'>
					<font size='1' face='trebuchet MS'><u>From</u></font><input type='text' class='rightjust' maxlength='10' name='TF1' size='10' maxlength='10'> 
					<font size='1' face='trebuchet MS'><u>To</u></font><input type='text' class='rightjust' size= '11' name='TF2' maxlength='10'><font face='trebuchet MS' size='1'>mm/dd/yyyy</font></td
				></tr
				><tr
					><td align='center' colspan='7'><input type='button' value='Generate Report' class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='document.frmSort.submit();'></td
				></tr
				><tr><td>&nbsp;</td></tr
				><tr
					><td align='center' colspan='7'><font color='red' face='trebuchet MS' size='1'><%=Session("MSG")%></font></td
				></tr>
			</table>
		
		<br>
		<%=Session("strTableScript")%>
		<br>
		
		<%If Session("strTableScript") <> "" Then %>
			<center>
				<%If LS <> 1 Then %>
					<input  style='width: 140px;' align='center' type='button' class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" value='Print Preview' onclick='PrintPrev();'>
				<%Else%>
					<input  style='width: 140px;' align='center' type='button' class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" value='Print table' onclick='JavaScript: LandWarn();'>
				<%End If%>
				<input  style='width: 140px;' align='center' type='button' class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" value='Export table to CSV' onclick='JavaScript: ExCSV();'>
				<%=Session("strButt")%>
		<%End If%>
		
		</form>
		<!-- #include file="_boxdown.asp" -->
	</body>
</html>
<%
'End If
Session("MSG") = ""
Session("strTableScript") = ""
Session("strButt") = ""
%>