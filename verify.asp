<%@Language=VBScript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<%
' load all entries for this user from database.
DIM		sbmtDATE, lngI, tblEMP, strSQL, strTableScript, strComHrs
DIM 	strComMon, strComTue, strComWed, strComThu, strComFri, strComSat, strComSun
DIM		tmpDATE, myDATE, monDATE, tueDATE, wedDATE, thuDATE, friDATE, satDATE, sunDATE
DIM		name, Ename, Edate, Eid, difDATE, myDATE2, finDATE, lname, fname, strLINK, mlMail
DIM		tblOthers, strSQLo

m_date = Request.Cookies("VTSHEETS")("VEdate")
m_id = Request.Cookies("VTSHEETS")("VEid")
m_name = Request.Cookies("VTSHEETS")("VEname")

tmpDATE = m_date
myDATE = CDate(tmpDATE)

If WeekdayName(Weekday(myDATE), true) = "Sun" Then
	finDATE = myDATE
	sunDATE = myDATE
	monDATE = DateAdd("d", 1, sunDATE)
	tueDATE = DateAdd("d", 1, monDATE)
	wedDATE = DateAdd("d", 1, tueDATE)
	thuDATE = DateAdd("d", 1, wedDATE)
	friDATE = DateAdd("d", 1, thuDATE)
	satDATE = DateAdd("d", 1, friDATE)
Else
	difDATE = DatePart("w", myDATE)
	sunDATE = DateAdd("d", -Cint(difDATE - 1), myDATE)
	myDATE2 = sunDATE
	finDATE = myDATE2
	monDATE = DateAdd("d", 1, sunDATE)
	tueDATE = DateAdd("d", 1, monDATE)
	wedDATE = DateAdd("d", 1, tueDATE)
	thuDATE = DateAdd("d", 1, wedDATE)
	friDATE = DateAdd("d", 1, thuDATE)
	satDATE = DateAdd("d", 1, friDATE)
End If
tmpd8 = sunDATEDATE

Set tblEMP = Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT * FROM [tsheets_t],[Worker_t] WHERE [emp_id] = '" & m_id & "' AND [date] = #" & finDATE &  "#  AND [Social_Security_Number] = '" & m_id & "'"
response.write "<!-- sql:" & strSQL & "-->"
tblEMP.Open strSQL, g_strCONN, 3, 1

name = m_name

On Error Resume Next
sbmtDATE = tblEMP("date_sub")

lngI = 0
Do While Not tblEMP.EOF
	If tblEMP("verify") Then 
		strDisabled = " disabled"
		ctr = ctr + 1 
	Else 
		strDisabled = ""
	End if
	
	if Z_IsOdd(lngI) = true then 
		kulay = "#FFFAF0" 
	else 
		kulay = "#FFFFFF"
	end if
	''''''GET NAME OF CONSUMER''''''
	Set tblCon = Server.CreateObject("ADODB.RecordSet")
	sqlCon = "SELECT * FROM Consumer_t WHERE Medicaid_Number ='" & tblEMP("Client") & "' "
	tblCon.Open sqlCon, g_strCONN, 1, 3
		If Not tblCon.EOF Then
			pangalan = tblCon("Lname") & ", " & tblCon("Fname")
			
		End If
	tblCon.Close
	Set tblCon = Nothing
	'''''''''''''''''''''''''''''''
	strTableScript = strTableScript & "<tr bgcolor='" & kulay & "'><td align='center' rowspan='2'>" & vbCrLf & _
			"<input type='checkbox' disabled size= '5' name='chk" & lngI & "' value='" & _
			tblEMP("id") & "'" & strDisabled & ">" & _
			"<input type='hidden' name='tmpID" & lngI & "' value='" & tblEMP("id") & "'></td>" & _
			"<td align='center' rowspan='2'><input type='text' readonly name='hdept" & lngI & "' size= '12' value='" & pangalan & "' " & strDisabled & "" & _
			"></td" & vbCrLf

	strTableScript = strTableScript & "><td align='center' rowspan='2'><input type='text' size= '5' class='rightjust' onblur='ComputeRows()'name='hsun" & _
			lngI & "'  readonly value='" & tblEMP("sun") & "' class='textbox' " & _
			"onblur=""vbscript: compute()"" " & strDisabled & "></td>" & _
			"<td align='center' rowspan='2'><input type='text' readonly size= '5' name='hmon" & _
			lngI & "' value='" & tblEMP("mon") & "' onblur='ComputeRows()' class='rightjust'  " & _
			"onblur=""vbscript: Compute()"" ></td" & vbCrLf & _
			"><td align='center' rowspan='2'><input type='text' onblur='ComputeRows()' class='rightjust' size= '5' name='htue" & _
			lngI & "' class='textbox' readonly value='" & tblEMP("tue") & "' " & _
			"onblur=""vbscript: compute()""></td" & vbCrLf & _
			"><td align='center' rowspan='2'><input type='text' class='rightjust' size= '5' onblur='ComputeRows()' name='hwed" & _
			lngI & "' class='textbox' readonly value='" & tblEMP("wed") & "' " & _
			"onblur=""vbscript: compute()""></td" & vbCrlf & _
			"><td align='center' rowspan='2'><input type='text' class='rightjust' size= '5' onblur='ComputeRows()' name='hthu" & _
			lngI & "' class='textbox'  readonly value='" & tblEMP("thu") &"' " & _
			"onblur=""vbscript: compute()""></td" & vbCrlf & _
			"><td align='center' rowspan='2'><input type='text' size= '5' class='rightjust' onblur='ComputeRows()' name='hfri" & _
			lngI & "' class='textbox' readonly value='" & tblEMP ("fri") & "' " & _
			"onblur=""vbscript: compute()""></td" & vbCrlf & _
			"><td align='center' rowspan='2'><input type='text' size= '5' class='rightjust' onblur='ComputeRows()' name='hsat" & _
			lngI & "'  readonly value='" & tblEMP("sat") & "' class='textbox' " & _
			"onblur=""vbscript: compute()"" ></td" 
			 
	
			strTableScript = strTableScript & "><td align='center' rowspan='2'><input name='htot" & lngI & "' " & _
					" readonly class='rightjust'  size='5' value='" & ( Z_CDbl(tblEmp("mon")) + Z_CDbl(tblEmp("tue")) + _
						Z_CDbl(tblEmp("wed")) + Z_CDbl(tblEmp("thu")) + _
						Z_CDbl(tblEmp("fri")) + Z_CDbl(tblEmp("sat")) + _
						Z_CDbl(tblEmp("sun")) ) & "' " & strDisabled & "></td" & vbCrLf
							
			strComputeCode = "document.frmTime.htot" & lngI & _
					".value = parseFloat(document.frmTime.hmon" & lngI & ".value) + " & _
					"parseFloat(document.frmTime.htue" & lngI & ".value) + " & _
					"parseFloat(document.frmTime.hwed" & lngI & ".value) + " & _
					"parseFloat(document.frmTime.hthu" & lngI & ".value) + " & _
					"parseFloat(document.frmTime.hfri" & lngI & ".value) + " & _
					"parseFloat(document.frmTime.hsat" & lngI & ".value) + " & _
					"parseFloat(document.frmTime.hsun" & lngI & ".value);"
	
	strTableScript = strTableScript & "><td align='center' rowspan='2'><textarea READONLY rows='2' name='Mnotes" & _
			lngI & "' cols='20'>" & tblEMP("misc_notes") & _
			"</textarea"  & vbCrLf 
					
	strTableScript = strTableScript & "><td align='center' rowspan='2'>" & _
			"<input type='checkbox' name='app" & lngI & "' value='" & tblEMP("id") & "'"
	
	If tblEMP("verify") then strTableScript = strTableScript & " checked"
		
	strTableScript = strTableScript & " ></td></tr><tr><td>&nbsp;</td></tr>"
	
	strComMon = strComMon & "parseFloat(document.frmTime.hmon" & lngI & ".value) +"
	strComTue = strComTue & "parseFloat(document.frmTime.htue" & lngI & ".value) +"
	strComWed = strComWed & "parseFloat(document.frmTime.hwed" & lngI & ".value) +"
	strComThu = strComThu & "parseFloat(document.frmTime.hthu" & lngI & ".value) +"
	strComFri = strComFri & "parseFloat(document.frmTime.hfri" & lngI & ".value) +"
	strComSat = strComSat & "parseFloat(document.frmTime.hsat" & lngI & ".value) +"
	strComSun = strComSun & "parseFloat(document.frmTime.hsun" & lngI & ".value) +"
	
	lngI = lngI + 1
	tblEMP.MoveNext
Loop


Set tblOthers = Server.CreateObject("ADODB.Recordset")
strSQLo	= "SELECT * FROM [report_t] WHERE [empid] = '" & m_id & "' AND [d8] = #" & m_date & "#"
On error resume next
tblOthers.Open strSQLo, g_strCONN, 3, 1
		strd8sub = tblOthers("d8_sub")
		strtdes = tblOthers("tdes")
		If tblOthers("stat") = "" then
			strstatus = "For Approval"
		Else
			If ctr = lngI then
				strstatus = "Approved"
			else
				strstatus = "For Approval"
			End if
		End if
	
tblOthers.Close
set tblOthers = Nothing	
Set tblUser = Server.CreateObject("ADODB.Recordset")
sqlUser = "SELECT * FROM [input_t] WHERE index = " & session("UserID")
response.write "<!-- sql" & sqlUser & "-->"
tblUser.Open sqlUser, g_strCONN, 3, 1
	If session("UserID") = "" Then
		session("MSG") = "Session timed out. Sign in again."
		response.redirect "default.asp"
	End IF
	'tmpUser = "ADMINISTRATOR"
	if not tblUser.EOF Then
		tmpUser = UCase(tblUser("lname")) & ", " & UCase(tblUser("fname"))
		response.write "<!-- name" & tmpUser & "-->"
	else
		session("MSG") = "Session timed out. Sign in again."
		response.redirect "default.asp"
		
	End If
tblUser.Close
set tblUser = Nothing
%>
<html>
<head>
<title>Timesheet - View</title>
<link href="styles.css" type="text/css" rel="stylesheet">

<SCRIPT LANGUAGE="JavaScript">
function logout()
{
	document.frmTime.action = "default.asp";
	document.frmTime.submit();
}
function ComputeRows() {
	<%=strComputeCode%>
}
function Compute()
{
var tmpHRS;
	document.frmTime.thmon.value = <%=strComMon%> 0;
	document.frmTime.thtue.value = <%=strComTue%> 0;
	document.frmTime.thwed.value = <%=strComWed%> 0;
	document.frmTime.ththu.value = <%=strComThu%> 0;
 	document.frmTime.thfri.value = <%=strComFri%> 0;
 	document.frmTime.thsat.value = <%=strComSat%> 0;
 	document.frmTime.thsun.value = <%=strComSun%> 0;
	document.frmTime.thtot.value = parseFloat(document.frmTime.thmon.value) +  parseFloat(document.frmTime.thtue.value) + parseFloat(document.frmTime.thwed.value) +  parseFloat(document.frmTime.ththu.value) + parseFloat(document.frmTime.thfri.value) +  parseFloat(document.frmTime.thsat.value) + parseFloat(document.frmTime.thsun.value);

 tmpHRS = document.frmTime.thtot.value;
 document.frmTime.thrs.value = tmpHRS;

}
</script>
</head>
<form method='post' name='frmTime' action="Vsubmit.asp">
	<body bgcolor='white'  LEFTMARGIN='0' TOPMARGIN='0' onload="vbscript:Compute()">

<table width='100%' height='100%' cellSpacing='0' cellPadding='0' border='0'>
<tr><td valign='top' colspan='3'>
	<table cellSpacing='0' cellPadding='0' width="100%" bgColor='#330066' border='0'>
			<tr vAlign=center
				><td align='left' style='height: 110px;'><img height='64' alt="Lutheran Social Services of New England" src="images/BannerTop.jpg" style='width: 100%; height: 100%;' border=0></td
			></tr>
	</table>
</td></tr>
<tr><td colspan='3'>
	<form method='post' name='frmmain'>
		<table align='left' border='0' width='100%' cellpadding='0' cellspacing='0'>
				<tr bgcolor='#A4CADB'>
					<td>
						<img src='images/Welcome.gif' border='0'>
						<a href='admin2.asp' ><img src='images/Home.gif' border='0'></a>
						
						<a href='default.asp' ><img src='images/SignOut.gif' border='0'></a>
					</td>
					<td align='left' style='width: 500px;'>
					<font face='trebuchet MS' size=3>&nbsp;&nbsp;User: <B><%=tmpUser%></b></font>
					</td>
					<td align='right'><a href='#'><img src='images/Help.gif' border='0'></a></td>
				</tr>
				<tr>
					<td >
					&nbsp;
					</td>
				</tr>	
		</table>
</td></tr>
<tr><td colspan='3' valign='top' style='height: 20px;'><img border='0' src='images/Topbar.gif' style='width: 100%;' height='25px'></td></tr>
<tr>
	<td valign='top' style='height: 100%;'><img border='0' src='images/Leftbar.gif' style='width: 31px;' height='100%'></td>
			<td>
				<table border='0' width=100%>
					<tr><td>
						<table border='0'>
							
	<tr>
		<td align='left' colspan='5'><font face='trebuchet MS'>&nbsp;&nbsp;PCS Worker:&nbsp;<b>&nbsp;<%=name%></b></font>
		</td>
	</tr>
</table>
<br>
<center>
<br><br>


<table align='center' border ='0'>
<tr align='right'><td align='right'><font face='trebuchet MS'>From</font></td
><td align='left' width='100px'><input type='text' name='1day' readonly size='10' value='<%=sunDATE%>'></td
><td align='right'><font face='trebuchet MS'>To</font></td
><td align='left' width='100px'><input type='text' size= '10' name='7day' readonly  value='<%=satDATE%>'></td
></tr>
</table>
<br><center><br>
<table bgcolor='white' border ='0'>
<tr class= 'info' bgcolor='white'><td class='title' width = '100px' align='center'><font face='trebuchet MS'><u>Date Submitted</u></font></td
><td class='title' width = '100px' align='center'><font face='trebuchet MS'><u>Total Hours</u></font></td
><td class='title' width = '200px' align='center'><font face='trebuchet MS'><u>Status</u></font></td></tr>
<tr class='curr'>
<td align='center'><input type='text'  name='subDATE' class='rightjust' size='10' readonly value='<%=strd8sub%>'></td>
<td align='center'><input type='text' size='10' class='rightjust' readonly name='thrs'></td>
<td align='center'><input type='text' readonly class='rightjust' size='20' name='status' value='<%=strstatus%>'></td></tr>
</table>
<br><br>
<table id='time_t' align='center' bgcolor='white' border='0' cellpadding='0' cellspacing = '0' >
<tr class='title' bgcolor='white'><td width= '35px'>&nbsp;</td
><td class='title' width= '75px' align='center'><font face='trebuchet MS'><u>Consumer</u></font></td
><td class='title' width= '75px' align='center'><font face='trebuchet MS'><u>Sun<br><%=Z_MDYDate(sunDATE)%></u></font></td
><td class='title' width= '75px' align='center'><font face='trebuchet MS'><u>Mon<br><%=Z_MDYDate(monDATE)%></u></font></td
><td class='title' width= '75px' align='center'><font face='trebuchet MS'><u>Tue<br><%=Z_MDYDate(tueDATE)%></u></font></td
><td class='title' width= '75px' align='center'><font face='trebuchet MS'><u>Wed<br><%=Z_MDYDate(wedDATE)%></u></font></td
><td class='title' width= '75px' align='center'><font face='trebuchet MS'><u>Thu<br><%=Z_MDYDate(thuDATE)%></u></font></td
><td class='title' width= '75px' align='center'><font face='trebuchet MS'><u>Fri<br><%=Z_MDYDate(friDATE)%></u></font></td
><td class='title' width= '75px' align='center'><font face='trebuchet MS'><u>Sat<br><%=Z_MDYDate(satDATE)%></u></font></td
><td class='title' width= '75px' align='center'><font face='trebuchet MS'><u> Total </u></font></td
><td class='title' width= '85px' align='center'><font face='trebuchet MS'><u>Notes</u></font></td
><td class='title' width= '35px'><font face='trebuchet MS'><u>Approved</u></font></td
></tr>
<%=strTableScript%>
<tr ><td ></td
><td></td
><td></td
></tr><tr bgcolor='white'><td border='1' >&nbsp;</td
><td class='title' align='right'><p align='center'><font face='trebuchet MS'><b>Totals</b></font></p></td
><td align='center' ><input type='text' size='6' readonly name='thsun' class='rightjust'></td
><td align='center' ><input type='text' size='6' readonly name='thmon' class='rightjust' ></td
><td align='center' ><input type='text' size='6' readonly name='thtue' class='rightjust' ></td
><td align='center' ><input type='text' size='6' readonly name='thwed' class='rightjust' ></td
><td align='center' ><input type='text' size='6' readonly name='ththu' class='rightjust' ></td
><td align='center' ><input type='text' size='6' readonly name='thfri' class='rightjust' ></td
><td align='center' ><input type='text' size='6' readonly name='thsat' class='rightjust' ></td
><td align='center'><input type='text' size='6' name='thtot' readonly class='rightjust'></td
><td>&nbsp;</td></tr>
</table>

<br><br>

<font face='trebuchet MS'>Admin ID: </font><input type="password" name="empID" size="10" style="font-family:Times New Roman" >
<input type='submit' value='Submit'> <BR>
<span class='error'><font face='trebuchet MS'><%=Session("MSG")%></font></span>
<br><br>
<input type='hidden' name='count' value='<%=lngI%>'>
<input type='hidden' name='tmpd8' value='<%=tmpd8%>'>
<input type='hidden' name='tmpID' value='<%=m_id%>'>
</td></tr>
				</table>
			</td>
			<td align='right' style='height: 100%;'><img border='0' src='images/Rightbar.gif' style='width: 31px;' height='100%'></td>	
		</tr>
	</form> 
</td></tr>
<tr><td colspan='3' valign='top' style='height: 20px;'><img border='0' src='images/Butbar.gif' style='width: 100%;' height='25px'></td></tr>
<tr vAlign=center
				><td colspan='3' align='left' ><img alt="Lutheran Social Services of New England" src="images/BotBanner.Gif" style='width: 100%;' border=0></td
			></tr>
</table>
</body>
</html>
<%Session("MSG") = ""%>