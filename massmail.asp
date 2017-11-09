<%Language=VBScript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_security.asp" -->
<%

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
	server.scripttimeout = 360000
	strEnd = "<br><br><font size='1' face='trebuchet MS'>* Please do not reply to this email. This is a computer generated email.</font>"
	Set rsEmail = Server.CreateObject("ADODB.RecordSet")
	sqlEmail = "SELECT * FROM Worker_T WHERE status = 'Active' ORDER BY lname, fname"
	rsEmail.Open sqlEmail, g_strCONN, 1, 3
	Do Until rsEmail.EOF
		If rsEmail("email") <> "" Then
			Set mlMail = CreateObject("CDO.Message")
			mlMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
			mlMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "localhost"
			mlMail.Configuration.Fields.Update
			myEmailAdr = rsEmail("email")
			mlMail.To = myEmailAdr
			mlMail.From = "info@smart-care.org"
			mlMail.Subject = Request("txtSub")
			strBody = "<font size='2' face='trebuchet MS'>" & Replace(Request("txtMSG"), vbCrLf, "<br>") & "</font>" & strEnd
			mlMail.HTMLBody = "<html><body>" & vbCrLf & strBody & vbCrLf & "</body></html>"
			mlMail.Send
			'response.write strBody & "<br><br><Br>"
			Set mlMail = Nothing
		End If
		rsEmail.MoveNext
	Loop
	rsEmail.Close
	Set rsEmail = Nothing
	Set mlMail = CreateObject("CDO.Message")
	mlMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
	mlMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "localhost"
	mlMail.Configuration.Fields.Update
	mlMail.To = "info@smart-care.org"
	mlMail.From = "info@smart-care.org"
	mlMail.Subject = Request("txtSub") & " - COPY"
	strBody = "<font size='2' face='trebuchet MS'>" & Replace(Request("txtMSG"), vbCrLf, "<br>") & "</font>" & strEnd
	mlMail.HTMLBody = "<html><body>" & vbCrLf & strBody & vbCrLf & "</body></html>"
	mlMail.Send
	'response.write strBody & "<br><br><Br>"
	Set mlMail = Nothing
	Session("MSG") = "Email Sent."
End If
%>
<html>
	<head>
		<title>LSS - In-Home Care - Mass eMail</title>
		<link href="styles.css" type="text/css" rel="stylesheet" media="print">
		<link href="CalendarControl.css" type="text/css" rel="stylesheet">
		<script src="CalendarControl.js" language="javascript"></script>
		<script language='JavaScript'>
		<!--
		function SendMe()
		{
			if (document.frmMemail.txtSub.value == "")
			{
				alert("Please include a subject.")
				return;
			}
			if (document.frmMemail.txtMSG.value == "")
			{
				alert("Please include a message.")
				return;
			}
			var ans = window.confirm("Send eMail? This might take a few minutes to complete.");
			if (ans){
			document.frmMemail.action = "massmail.asp";
			document.frmMemail.submit();
			}
		}
		function bawal(tmpform)
		{
			var iChars = ",|\"\'";
			var tmp = "";
			for (var i = 0; i < tmpform.value.length; i++)
			 {
			  	if (iChars.indexOf(tmpform.value.charAt(i)) != -1)
			  	{
			  		alert ("This character is not allowed.");
			  		tmpform.value = tmp;
			  		return;
		  		}
			  	else
		  		{
		  			tmp = tmp + tmpform.value.charAt(i);
		  		}
		  	}
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

	
		<form method='post' name='frmMemail'>
			<br><br>
			<table cellSpacing='0' cellPadding='0' align='center' border='0'>
				<tr>
								<td align='center' colspan='10'>
									<div name="dErr" style="width: 250px; height:55px;OVERFLOW: auto;">
										<table border='0' cellspacing='1'>		
											<tr>
												<td><span class='error'><%=Session("MSG")%></span></td>
											</tr>
										</table>
									</div>
								</td>
							</tr>
							<tr>
								<td align='center'>
									<table border='0' cellspacing='1'>
											<tr>
											<td valign='top'>Subject:</td>
											<td>
												<input class='main' size='50' maxlength='50' name='txtSub' onkeyup='bawal(this);'>
											</td>
										</tr>
										<tr>
											<td valign='top'>Message:</td>
											<td>
												<textarea class='main' style='width: 400px;' name='txtMSG' cols='75' rows='10' onkeyup='bawal(this);'>
												</textarea>
											</td>
										</tr>
										<tr><td>&nbsp;</td></tr>
										<tr>
											<td colspan='2' align='right'>
												<input class='btn' type='button' value='Send' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='SendMe();'>
											</td>
										</tr>
									</table>
								</td>
							</tr>
						</table>
					</td>
				</tr>
			</form>
		<!-- #include file="_boxdown.asp" -->
	</body>
</html>
<%
If Session("MSG") <> "" Then
	tmpMSG = Session("MSG")
%>
<script><!--
	alert("<%=tmpMSG%>");
--></script>
<%
End If
Session("MSG") = ""
%>