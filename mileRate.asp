<%Language=VBScript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_security.asp" -->
<%
Function GoToSunday(keyDate)
    GoToSunday = DateSerial(Year(keyDate) _
        , Month(keyDate) _
        , Day(keyDate) - DatePart("w", keyDate) + 1)
End Function
If UCase(Session("lngType")) <> "2" Then
	Session("MSG") = "Invalid User Type. Please Sign In again."
	Response.Redirect "default.asp"
End If
'CHANGE RATE
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
	'ERROR CHECK
	If Not IsNumeric(Request("txtNewRate")) Then 
		Session("MSG") = "ERROR: Enter valid rate."
	End If
	If Not IsDate(Request("RateDate")) Then 
		Session("MSG") = Session("MSG") & "<br>ERROR: Enter valid date."
	Else
		If Cdate(Request("RateDate")) < Cdate(Request("txtdateH")) Then
			Session("MSG") = Session("MSG") & "<br>ERROR: Date inputted is earlier than the original date."
		End If
	End If
	If Session("MSG") <> "" Then
		Response.Redirect "mileRate.asp"
	End If
	
	Set rsNRate = Server.CreateObject("ADODB.RecordSet")
	sqlNRate = "SELECT * FROM mileRate_T"
	rsNRate.Open sqlNRate, g_strCONN, 1, 3
	rsNRate.AddNew
	rsNRate("milerate") = Request("txtNewRate")	
	rsNRate("miledate") = GoToSunday(Request("RateDate"))
	rsNRate.Update
	rsNRate.Close
	Set rsNRate = Nothing
on error resume next
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set ALog = fso.OpenTextFile(AdminLog, 8, True)
	Set rsName = CreateObject("ADODB.RecordSet")
	sqlName = "SELECT * FROM input_T WHERE index = " & Session("UserID")
	rsName.Open sqlName, g_strCONN, 3, 1
  tmpName = "N/A"
	If Not rsName.EOF Then
		tmpName = rsName("lname") & ", " & rsName("fname")
	End If
	rsName.Close
	Set rsName = Nothing
	Alog.WriteLine Now & ":: Mileage Rate has been changed from: $" & Z_FormatNumber(Request("txtrateH"), 2) & " to: $" & Z_FormatNumber(Request("txtNewRate"), 2) & _
		 " effective on " & Request("RateDate") & " by " & Session("UserID") & " - " & tmpName & ". "  & vbCrLf
	Set Alog = Nothing
	Set fso = Nothing 
	
End If
'GET RATE
Set rsRate = Server.CreateObject("ADODB.RecordSet")
sqlRate = "SELECT * FROM mileRate_T"
rsRate.Open sqlRate, g_strCONN, 3, 1
If Not rsRate.EOF Then
	rsRate.MoveLast
	tmpRate = rsRate("mileRate")
	tmpDate = rsRate("mileDate")
End If
rsRate.Close
Set rsRate = Nothing
%>
<html>
	<head>
		<title>LSS - In-Home Care - Administrator Tools - Mileage Rate Tools</title>
		<link href="CalendarControl.css" type="text/css" rel="stylesheet">
		<script src="CalendarControl.js" language="javascript"></script>
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
		<form method='post' name='frmRate' action="mileRate.asp">
			<br><br>
			<table cellSpacing='0' cellPadding='0' align='center' border='0'>
				<tr bgcolor='#040C8B'>
					<td colspan='3' align='center' style="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=0, startColorstr=#040C8B, endColorstr=#c7c9e5);">
						<font face='trebuchet MS' size='2' color='white'><b>Mileage Rate Tools</b></font>
					</td>
				</tr>
				<tr><td>&nbsp;</td></tr>
				<tr>
					<td><font size='1' face='trebuchet MS'>Current Mileage Rate:&nbsp;</font></td>
					<td><input type="text" size='4' maxlength='5' name='txtCurRate' disabled value='$<%=tmpRate%>'>
					&nbsp;&nbsp;<font size='1' face='trebuchet MS'>as of</font></td>
					<td><input type="text" size='10' maxlength='10' name='txtCurDate' disabled value='<%=tmpDate%>'></td>
					<input type="hidden" name='txtdateH' value='<%=tmpDate%>'>
					<input type="hidden" name='txtrateH' value='<%=tmpRate%>'>
				</tr>
				<tr><td>&nbsp;</td></tr>
				<tr>
					<td align='right'><font size='1' face='trebuchet MS'>New Mileage Rate:&nbsp;</font></td>
					<td><input type="text" size='4' maxlength='5' name='txtNewRate'></td>
				</tr>
				<tr>
					<td><font size='1' face='trebuchet MS'>Effective On:</font></td>
					<td align='right'>
						<input readonly tabindex="1" name='RateDate' style='width:80px;'
						type="text" maxlength='10' onchange='weeknum();''><input tabindex="2" type="button" value="..." name="cal1" style="width: 15px;"
						onclick="showCalendarControl(document.frmRate.RateDate);" class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'">
					</td>
				</tr>
				<tr>
					<td colspan='3'>
						<font size='1' face='trebuchet MS'>*date chosen will be converted to a sunday of the same week</font>
					</td>
				</tr>
				<tr><td>&nbsp;</td></tr>
				<tr>
					<td colspan='3' align='center'>
						<input type="button" value='Save' style='width: 150px;' class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='document.frmRate.submit();'>
					</td>
				</tr>
				<tr>
					<td colspan='3' align='center'>
						<span><font face='Trebuchet MS' size='1' color='red'><%=Session("MSG")%></font></span>
					</td>
				</tr>
			</table>
		</form>
		<!-- #include file="_boxdown.asp" -->
	</body>
</html>
<%Session("MSG") = ""%>