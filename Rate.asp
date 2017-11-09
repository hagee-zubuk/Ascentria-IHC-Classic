<%Language=VBScript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_security.asp" -->
<%
If UCase(Session("lngType")) <> "2" Then
	Session("MSG") = "Invalid User Type. Please Sign In again."
	Response.Redirect "default.asp"
End If
'CHANGE RATE
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
	If Request("ctrl") = 1 Then
		'ERROR CHECK
		If Not IsNumeric(Request("txtNewRate")) Then 
			Session("MSG") = "ERROR: Enter valid Medicaid rate."
		End If
		If Not IsDate(Request("RateDate")) Then 
			Session("MSG") = Session("MSG") & "<br>ERROR: Enter valid Medicaid date."
		Else
			If Z_FixNull(Request("txtdateH")) <> "" Then
				If Cdate(Request("RateDate")) < Cdate(Request("txtdateH")) Then
					Session("MSG") = Session("MSG") & "<br>ERROR: Medicaid Date inputted is earlier than the original Medicaid date."
				End If
			End If
		End If
		If Session("MSG") <> "" Then
			Response.Redirect "Rate.asp"
		End If
		
		Set rsNRate = Server.CreateObject("ADODB.RecordSet")
		sqlNRate = "SELECT * FROM Rate_T"
		rsNRate.Open sqlNRate, g_strCONN, 1, 3
		rsNRate.AddNew
		rsNRate("Rate") = Z_NullFix(Request("txtNewRate"))	
		rsNRate("rDate") = Z_DateNull(Request("RateDate"))
		rsNRate("VHArate") = Z_NullFix(Request("txtHArateH"))	
		rsNRate("VHAdate") = Z_DateNull(Request("txtHAdateH"))
		rsNRate("VHMrate") = Z_NullFix(Request("txtHMrateH"))	
		rsNRate("VHMdate") = Z_DateNull(Request("txtHMdateH"))
		rsNRate.Update
		rsNRate.Close
		Set rsNRate = Nothing
		
		Set fso = CreateObject("Scripting.FileSystemObject")
		Set ALog = fso.OpenTextFile(AdminLog, 8, True)
		Set rsName = CreateObject("ADODB.RecordSet")
		sqlName = "SELECT * FROM input_T WHERE [index] = " & Session("UserID")
		rsName.Open sqlName, g_strCONN, 3, 1
	  tmpName = "N/A"
		If Not rsName.EOF Then
			tmpName = rsName("lname") & ", " & rsName("fname")
		End If
		rsName.Close
		Set rsName = Nothing
		Alog.WriteLine Now & ":: Medicaid Rate has been changed from: $" & Z_FormatNumber(Request("txtrateH"), 2) & " to: $" & Z_FormatNumber(Request("txtNewRate"), 2) & _
			 " effective on " & Request("RateDate") & " by " & Session("UserID") & " - " & tmpName & ". "  & vbCrLf
		Set Alog = Nothing
		Set fso = Nothing 
	ElseIf Request("ctrl") = 2 Then
		'ERROR CHECK
		If Not IsNumeric(Request("txtHANewRate")) Then 
			Session("MSG") = "ERROR: Enter valid VA-HA rate."
		End If
		If Not IsDate(Request("HARateDate")) Then 
			Session("MSG") = Session("MSG") & "<br>ERROR: Enter valid VA-HA Rate date."
		Else
			If Z_FixNull(Request("txtHAdateH")) <> "" Then
				If Cdate(Request("HARateDate")) < Cdate(Request("txtHAdateH")) Then
					Session("MSG") = Session("MSG") & "<br>ERROR: VA-HA Date inputted is earlier than the original VA-HA date."
				End If
			End If
		End If
		If Session("MSG") <> "" Then
			Response.Redirect "Rate.asp"
		End If
		
		Set rsNRate = Server.CreateObject("ADODB.RecordSet")
		sqlNRate = "SELECT * FROM Rate_T"
		rsNRate.Open sqlNRate, g_strCONN, 1, 3
		rsNRate.AddNew
		rsNRate("VHArate") = Z_NullFix(Request("txtHANewRate"))	
		rsNRate("VHAdate") = Z_DateNull(Request("HARateDate"))
		rsNRate("VHMrate") = Z_NullFix(Request("txtHMrateH"))	
		rsNRate("VHMdate") = Z_DateNull(Request("txtHMdateH"))
		rsNRate("Rate") = Z_NullFix(Request("txtrateH"))	
		rsNRate("rDate") = Z_DateNull(Request("txtdateH"))
		rsNRate.Update
		rsNRate.Close
		Set rsNRate = Nothing
		
		Set fso = CreateObject("Scripting.FileSystemObject")
		Set ALog = fso.OpenTextFile(AdminLog, 8, True)
		Set rsName = CreateObject("ADODB.RecordSet")
		sqlName = "SELECT * FROM input_T WHERE [index] = " & Session("UserID")
		rsName.Open sqlName, g_strCONN, 3, 1
	  tmpName = "N/A"
		If Not rsName.EOF Then
			tmpName = rsName("lname") & ", " & rsName("fname")
		End If
		rsName.Close
		Set rsName = Nothing
		Alog.WriteLine Now & ":: VA-HA Rate has been changed from: $" & Z_FormatNumber(Request("txtHArateH"), 2) & " to: $" & Z_FormatNumber(Request("txtHANewRate"), 2) & _
			 " effective on " & Request("HARateDate") & " by " & Session("UserID") & " - " & tmpName & ". "  & vbCrLf
		Set Alog = Nothing
		Set fso = Nothing 
	ElseIf Request("ctrl") = 3 Then
		'ERROR CHECK
		If Not IsNumeric(Request("txtHMNewRate")) Then 
			Session("MSG") = "ERROR: Enter valid VA-HM rate."
		End If
		If Not IsDate(Request("HMRateDate")) Then 
			Session("MSG") = Session("MSG") & "<br>ERROR: Enter valid VA-HM Rate date."
		Else
			If Z_FixNull(Request("txtHMdateH")) <> "" Then
				If Cdate(Request("HMRateDate")) < Cdate(Request("txtHMdateH")) Then
					Session("MSG") = Session("MSG") & "<br>ERROR: VA-HM Date inputted is earlier than the original VA-HM date."
				End If
			End if
		End If
		If Session("MSG") <> "" Then
			Response.Redirect "Rate.asp"
		End If
		
		Set rsNRate = Server.CreateObject("ADODB.RecordSet")
		sqlNRate = "SELECT * FROM Rate_T"
		rsNRate.Open sqlNRate, g_strCONN, 1, 3
		rsNRate.AddNew
		rsNRate("VHMrate") = Z_NullFix(Request("txtHMNewRate"))	
		rsNRate("VHMdate") = Z_DateNull(Request("HMRateDate"))
		rsNRate("Rate") = Z_NullFix(Request("txtrateH"))	
		rsNRate("rDate") = Z_DateNull(Request("txtdateH"))
		rsNRate("VHArate") = Z_NullFix(Request("txtHArateH"))	
		rsNRate("VHAdate") = Z_DateNull(Request("txtHAdateH"))
		rsNRate.Update
		rsNRate.Close
		Set rsNRate = Nothing
		
		Set fso = CreateObject("Scripting.FileSystemObject")
		Set ALog = fso.OpenTextFile(AdminLog, 8, True)
		Set rsName = CreateObject("ADODB.RecordSet")
		sqlName = "SELECT * FROM input_T WHERE [index] = " & Session("UserID")
		rsName.Open sqlName, g_strCONN, 3, 1
	  tmpName = "N/A"
		If Not rsName.EOF Then
			tmpName = rsName("lname") & ", " & rsName("fname")
		End If
		rsName.Close
		Set rsName = Nothing
		Alog.WriteLine Now & ":: VA-HM Rate has been changed from: $" & Z_FormatNumber(Request("txtHMrateH"), 2) & " to: $" & Z_FormatNumber(Request("txtHMNewRate"), 2) & _
			 " effective on " & Request("HMRateDate") & " by " & Session("UserID") & " - " & tmpName & ". "  & vbCrLf
		Set Alog = Nothing
		Set fso = Nothing 
	End if	
End If
'GET RATE
Set rsRate = Server.CreateObject("ADODB.RecordSet")
sqlRate = "SELECT * FROM Rate_T"
rsRate.Open sqlRate, g_strCONN, 3, 1
If Not rsRate.EOF Then
	rsRate.MoveLast
	tmpRate = rsRate("Rate")
	tmpDate = rsRate("rDate")
	tmpHArate	= rsRate("VHArate")
	tmpHAdate = rsRate("VHAdate")
	tmpHMrate	= rsRate("VHMrate")
	tmpHMdate = rsRate("VHMdate")
End If
rsRate.Close
Set rsRate = Nothing
%>
<html>
	<head>
		<title>LSS - In-Home Care - Administrator Tools - Rate Tools</title>
		<link href="CalendarControl.css" type="text/css" rel="stylesheet">
		<script src="CalendarControl.js" language="javascript"></script>
		<script language='JavaScript'>
		<!--
		function SaveMe(xxx)
		{
			document.frmRate.action = "Rate.asp?ctrl=" + xxx;
			document.frmRate.submit();
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
	<body bgcolor='white'  LEFTMARGIN='0' TOPMARGIN='0'>
		<!-- #include file="_boxup.asp" -->
		<!-- #include file="_NavHeader.asp" -->
		<form method='post' name='frmRate' action="Rate.asp">
			<br><br>
			<table cellSpacing='0' cellPadding='0' align='center' border='0'>
				<tr bgcolor='#040C8B'>
					<td colspan='3' align='center' style="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=0, startColorstr=#040C8B, endColorstr=#c7c9e5);">
						<font face='trebuchet MS' size='2' color='white'><b>Medicaid Rate</b></font>
					</td>
				</tr>
				<tr><td>&nbsp;</td></tr>
				<tr>
					<td><font size='1' face='trebuchet MS'>Current Rate:&nbsp;</font></td>
					<td><input type="text" size='4' maxlength='5' name='txtCurRate' disabled value='$<%=tmpRate%>'>
					&nbsp;&nbsp;<font size='1' face='trebuchet MS'>as of</font></td>
					<td><input type="text" size='10' maxlength='10' name='txtCurDate' disabled value='<%=tmpDate%>'></td>
					<input type="hidden" name='txtdateH' value='<%=tmpDate%>'>
					<input type="hidden" name='txtrateH' value='<%=tmpRate%>'>
				</tr>
				<tr><td>&nbsp;</td></tr>
				<tr>
					<td align='right'><font size='1' face='trebuchet MS'>New Rate:&nbsp;</font></td>
					<td><input type="text" size='4' maxlength='5' name='txtNewRate'></td>
				</tr>
				<tr>
					<td><font size='1' face='trebuchet MS'>Effective On:</font></td>
					<td align='right'>
						<input tabindex="1" name='RateDate' style='width:80px;'
						type="text" maxlength='10' onchange='weeknum();''><input tabindex="2" type="button" value="..." name="cal1" style="width: 15px;"
						onclick="showCalendarControl(document.frmRate.RateDate);" class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'"></td>
				</tr>
				<tr><td>&nbsp;</td></tr>
				<tr>
					<td colspan='3' align='center'>
						<input type="button" value='Save Medicaid Rate' style='width: 150px;' class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='SaveMe(1);'>
					</td>
				</tr>
				<tr><td>&nbsp;</td></tr><tr><td>&nbsp;</td></tr>
				<tr bgcolor='#040C8B'>
					<td colspan='3' align='center' style="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=0, startColorstr=#040C8B, endColorstr=#c7c9e5);">
						<font face='trebuchet MS' size='2' color='white'><b>VA-HA Rate</b></font>
					</td>
				</tr>
				<tr><td>&nbsp;</td></tr>
				<tr>
					<td><font size='1' face='trebuchet MS'>Current Rate:&nbsp;</font></td>
					<td><input type="text" size='4' maxlength='5' name='txtHACurRate' disabled value='$<%=tmpHARate%>'>
					&nbsp;&nbsp;<font size='1' face='trebuchet MS'>as of</font></td>
					<td><input type="text" size='10' maxlength='10' name='txtHACurDate' disabled value='<%=tmpHADate%>'></td>
					<input type="hidden" name='txtHAdateH' value='<%=tmpHADate%>'>
					<input type="hidden" name='txtHArateH' value='<%=tmpHARate%>'>
				</tr>
				<tr><td>&nbsp;</td></tr>
				<tr>
					<td align='right'><font size='1' face='trebuchet MS'>New Rate:&nbsp;</font></td>
					<td><input type="text" size='4' maxlength='5' name='txtHANewRate'></td>
				</tr>
				<tr>
					<td><font size='1' face='trebuchet MS'>Effective On:</font></td>
					<td align='right'>
						<input tabindex="1" name='HARateDate' style='width:80px;'
						type="text" maxlength='10' onchange='weeknum();''><input tabindex="2" type="button" value="..." name="cal1" style="width: 15px;"
						onclick="showCalendarControl(document.frmRate.HARateDate);" class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'"></td>
				</tr>
				<tr><td>&nbsp;</td></tr>
				<tr>
					<td colspan='3' align='center'>
						<input type="button" value='Save VA-HA Rate' style='width: 150px;' class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='SaveMe(2);'>
					</td>
				</tr>
				<tr><td>&nbsp;</td></tr><tr><td>&nbsp;</td></tr>
				<tr bgcolor='#040C8B'>
					<td colspan='3' align='center' style="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=0, startColorstr=#040C8B, endColorstr=#c7c9e5);">
						<font face='trebuchet MS' size='2' color='white'><b>VA-HM Rate</b></font>
					</td>
				</tr>
				<tr><td>&nbsp;</td></tr>
				<tr>
					<td><font size='1' face='trebuchet MS'>Current Rate:&nbsp;</font></td>
					<td><input type="text" size='4' maxlength='5' name='txtHMCurRate' disabled value='$<%=tmpHMRate%>'>
					&nbsp;&nbsp;<font size='1' face='trebuchet MS'>as of</font></td>
					<td><input type="text" size='10' maxlength='10' name='txtHMCurDate' disabled value='<%=tmpHMDate%>'></td>
					<input type="hidden" name='txtHMdateH' value='<%=tmpHMDate%>'>
					<input type="hidden" name='txtHMrateH' value='<%=tmpHMRate%>'>
				</tr>
				<tr><td>&nbsp;</td></tr>
				<tr>
					<td align='right'><font size='1' face='trebuchet MS'>New Rate:&nbsp;</font></td>
					<td><input type="text" size='4' maxlength='5' name='txtHMNewRate'></td>
				</tr>
				<tr>
					<td><font size='1' face='trebuchet MS'>Effective On:</font></td>
					<td align='right'>
						<input tabindex="1" name='HMRateDate' style='width:80px;'
						type="text" maxlength='10' onchange='weeknum();''><input tabindex="2" type="button" value="..." name="cal1" style="width: 15px;"
						onclick="showCalendarControl(document.frmRate.HMRateDate);" class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'"></td>
				</tr>
				<tr><td>&nbsp;</td></tr>
				<tr>
					<td colspan='3' align='center'>
						<input type="button" value='Save VA-HM Rate' style='width: 150px;' class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='SaveMe(3);'>
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