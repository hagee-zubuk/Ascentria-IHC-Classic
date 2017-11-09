<%@Language=VBScript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_security.asp" -->
<%
'If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
'	If Request("Bill") <> 1 Then
'		tmpFile = Z_DoDecrypt(Session("dload"))
'		Set dload = Server.CreateObject("SCUpload.Upload")
'			dload.Download tmpFile
'		Set dload = Nothing
'	ElseIf Request("Bill") = 1 Then
'		tmpFile = Z_DoDecrypt(Session("dload2"))
'		Set dload = Server.CreateObject("SCUpload.Upload")
'			dload.Download tmpFile
'		Set dload = Nothing
'	End If
'End If

'If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
	'myctr = Request("ctr")
	'If myctr = 1 Then 'oUpload.Form("reqID")
	tmpFile = Z_DoDecrypt(Session("dload"))
	'ElseIf Request("ctr") = 2 Then
	'	tmpFile = Z_DoDecrypt(Session("dload2"))
		
	'End If
	tmpstring = copyfile2 & tmpFile
	'response.write tmpfile
	'Set dload = Server.CreateObject("SCUpload.Upload")
	'	dload.Download "C:\work\lss-dbvortex\export\consumermedicaid.csv",,,False 'tmpFile "C:\INETPUB\ATTACHMENT01294955.DAT", "PICTURE.GIF", , False
	'Set dload = Nothing
	tmpfile2 = Z_DoDecrypt(Session("dload2"))
	tmpstring2 = copyfile2 & tmpFile2
'End If
%>
<html>
	<head>
	<title>LSS - Download</title>
	<script language='JavaScript'>
		function dload()
		{
			document.frmdload.action = "extra.asp";
			document.frmdload.submit();
		}
		function dload2()
		{
			document.frmdload.action = "extra.asp?bill=1";
			document.frmdload.submit();
		}
	</script>
	</head>
	<body>
		<form method='post' name='frmdload' action="extra.asp">
			<a href='#' onclick="document.location='<%=tmpstring%>';"><font size='2' face='trebuchet ms'>download csv file</font></a>
			
				<br>
			<a href='#' onclick="document.location='<%=tmpstring2%>';"><font size='2' face='trebuchet ms'>download import file</font></a>
			<br>
			
			<a href='process.asp'><font size='2' face='trebuchet ms'>&gt;back to Process page</font></a>
			<input type='hidden' name='proc' value='<%=Request("proc")%>'>	
		</form>
	</body>
</html>

