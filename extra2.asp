<%@Language=VBScript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_security.asp" -->
<%
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
			document.frmdload.action = "extra2.asp?ctr=1";
			document.frmdload.submit();
		}
		function dload2()
		{
			document.frmdload.action = "extra2.asp?ctr=2";
			document.frmdload.submit();
		}
	</script>
	</head>
	<body>
		<form method='post' name='frmdload' action="">
			<% If Request("yyy") <> 1 Then%> 
			<a href='#' onclick="document.location='<%=tmpstring%>';"><font size='2' face='trebuchet ms'>download csv file</font></a>
			<% End If %>
			<% If Request("xxx") = 1 And Request("zzz") = 1 Then %>
					<br>
				<% If session("UserID") = 893 Or session("UserID") = 67 Or session("UserID") = 2 Then%>
					<% If Z_DoDecrypt(Session("dload2")) <> "NONE" Then %>
						<a href='#' onclick="document.location='<%=tmpstring2%>';"><font size='2' face='trebuchet ms'>download prfjdepi.csv</font></a>
					<% Else %>
						prfjdepi.csv WAS ALREADY RUN
					<% End If %>
				<% End If %>
			<% End If %>
				<br><br>
			<a href='specRep.asp'><font size='2' face='trebuchet ms'>&gt;back to Reports page</font></a>
			<input type='hidden' name='proc' value='<%=Request("proc")%>'>	
		</form>
	</body>
</html>