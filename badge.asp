<%@Language=VBScript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_security.asp" -->
<%
Function Z_ChkUBadgeUse(obadge)
	Z_ChkUBadgeUse = False
	Set rsBadge = Server.CreateObject("ADODB.RecordSet")
	rsBadge.Open "SELECT * FROM Worker_T WHERE ubadge = '" & obadge & "'", g_strCONN, 3, 1
	If Not rsBadge.EOF Then Z_ChkUBadgeUse = True
	rsBadge.Close
	Set rsBadge = Nothing
End Function
Function Z_UBadge(obadge)
	Z_UBadge = ""
	Set rsBadge = Server.CreateObject("ADODB.RecordSet")
	rsBadge.Open "SELECT ubadge FROM Worker_T WHERE badge = '" & obadge & "'", g_strCONN, 3, 1
	If Not rsBadge.EOF Then
		Z_UBadge = rsBadge("ubadge")
	End If
	rsBadge.Close
	Set rsBadge = Nothing
End Function
hideme = True
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
	Server.ScriptTimeout = 10800
	'upload file to server
	Set oUpload = Server.CreateObject("SCUpload.Upload")
	oUpload.Upload
	If oUpload.Files.Count = 0 Then
		Set oUpload = Nothing
		Session("MSG") = "Please specify a file to import."
		Response.Redirect "badge.asp"
	End If
	oFileName = oUpload.Files(1).Item(1).filename
	If Z_GetExt(oFileName) <> "TXT" Then
		Set oUpload = Nothing
		Session("MSG") = "Invalid File."
		Response.Redirect "badge.asp"
	End If
	nFileName = "vortex-lssnh.CSV" '"CONVERTED" & oFileName & ".CSV"
	oUpload.Files(1).Item(1).Save ouputPath, oFileName
	Set oUpload = Nothing
	
	'read file
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set oCSV = fso.OpenTextFile(ouputPath & oFileName, 1)
	Set nFile = fso.OpenTextFile(ouputPath & nFileName, 2, True)
	nFile.WriteLine "Card Number,Date,Punch Time"
	ctr = 1
	Do Until oCSV.AtEndofStream
		'read csv
		strLine = oCSV.ReadLine
		If Trim(strLine) <> "" Then
			csvValue = Split(strLine, ",")
			'CHECK VALUES
			If Ubound(csvValue) <> 3 Then
				Session("MSG") = Session("MSG") & "ERROR: Incorrect file format (LINE: " & ctr & "). File not converted.<br>"
				Response.Redirect "badge.asp"
			End If
			If Len(csvValue(0)) <> 6 Then 
				Session("MSG") = Session("MSG") & "ERROR: First value of row is invalid (LINE: " & ctr & "). File not converted.<br>"
				Response.Redirect "badge.asp"
			End If
			If Not IsDate(csvValue(1)) Then
				Session("MSG") = Session("MSG") & "ERROR: Second value of row is invalid (LINE: " & ctr & "). File not converted.<br>"
				Response.Redirect "badge.asp"
			End If
			If Not IsDate(csvValue(2)) Then
				Session("MSG") = Session("MSG") & "ERROR: Third value of row is invalid (LINE: " & ctr & "). File not converted.<br>"
				Response.Redirect "badge.asp"
			End If
			'''''''''''''
			If Z_ChkUBadgeUse(csvValue(0)) Then 
				uBadge = csvValue(0)
			Else
				uBadge = Z_UBadge(csvValue(0))
			End If
			If ubadge <> "" Then
				nFile.WriteLine uBadge & "," & csvValue(1) & "," & csvValue(2)
			Else
				Session("MSG") = Session("MSG") & "Ultipro Badge ID not found (LINE: " & ctr & ").<br>"
			End If
			ctr = ctr + 1
		End If	
	Loop	

	Session("MSG") = Session("MSG") & "File Converted."
	Set nFile = Nothing
	Set oCSV = Nothing
	Set fso = Nothing
	
	hideme = False
End If
%>
<html>
	<head>
		<title>LSS - In-Home Care - Ultipro Tool</title>
		<script language='JavaScript'>
		function dlme(xxx) {
			document.frmbadge.action = "dlbadge.asp?fname=" + xxx;
			document.frmbadge.submit();
		}
		function hideme() {
		<% If hideme Then %>
			document.frmbadge.dl.style.visibility = 'hidden';
		<% Else %>
			document.frmbadge.dl.style.visibility = 'visible';
		<% End If %>
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
	<body bgcolor='white' LEFTMARGIN='0' TOPMARGIN='0' onload='hideme();'>
		<form method="POST" enctype="multipart/form-data" name='frmbadge'>
			<!-- #include file="_boxup.asp" -->
			<!-- #include file="_NavHeader.asp" -->
			<table border='0' align='center'>
				<tr><td colspan='2' align='center'><font color='red'  face='trebuchet MS' size='1'>&nbsp;<%=Session("MSG")%>&nbsp;</font></td></tr>
				<tr>
					<td align='center'>
						<input type="file" name="F1" size="20"><input type='submit' value='Convert File' style='width: 100px;' class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'">
					</td>
				</tr>
				<tr>
				<tr><td>&nbsp;</td></tr>
				<tr><td>&nbsp;</td></tr>
				<tr>
				<td align='center'>
						<input type='button' name='dl' value='Download File' style='width: 200px;' class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick="dlme('<%=nfilename%>');">
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