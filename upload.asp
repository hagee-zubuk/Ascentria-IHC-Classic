<%@Language=VBScript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_security.asp" -->
<%
	tmpFilename = Z_GenerateGUID()
	Do Until GUIDExists(tmpFilename) = False
		tmpFilename = Z_GenerateGUID()
	Loop
	MNum = Request("MNum")
	wid = Request("wid")
	uploads = 0
	If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
		Set oUpload = Server.CreateObject("SCUpload.Upload")
		oUpload.Upload
		If oUpload.Files.Count = 0 Then
			Set oUpload = Nothing
			Session("MSG") = "Please specify a file to import."
			Response.Redirect "upload.asp"
		End If
		oFileName = oUpload.Files(1).Item(1).filename
		If Z_GetExt(oFileName) <> "PDF" Then
			Set oUpload = Nothing
			Session("MSG") = "Invalid File."
			Response.Redirect "upload.asp"
		End If
		oFileSize = oUpload.Files(1).Item(1).Size
		If oFileSize > 2097152 Then
			Set oUpload = Nothing
			Session("MSG") = "File is too large."
			Response.Redirect "upload.asp"
		End If
		nFileName = tmpFilename & ".PDF"
		oUpload.Files(1).Item(1).Save uploadFilePath, nFileName
		Set oUpload = Nothing
		'save upload
		Set rsUpload = Server.CreateObject("ADODB.RecordSet")
		If Mnum <> "" Then sqlUp = "SELECT * FROM uploads_T WHERE cid = '" & MNum & "' "
		If wid <> "" Then sqlUp = "SELECT * FROM uploads_T WHERE wid = '" & wid & "' "
		rsUpload.Open sqlUp, g_strCONN, 1, 3
		rsUpload.AddNew
		rsUpload("cid") = MNum
		rsUpload("wid") = wid
		rsUpload("ofilename") = oFileName
		rsUpload("datestamp") = Now
		rsUpload("filename") = nFileName
		rsUpload.Update
		rsUpload.Close
		Set rsUpload = Nothing
		uploads = 1
		
	End If
%>
<html>
	<head>
		<title>LSS - In-Home Care - Upload File</title>
		<link href='style.css' type='text/css' rel='stylesheet'>
		<script language='JavaScript'>
		<!--
		function uploadFile() {
			if (document.frmUpload.F1.value != "") {
				filestr = document.frmUpload.F1.value.toUpperCase();
				if (filestr.indexOf(".PDF") == -1) {
					alert("ERROR: Incorrect file extension.")
					document.frmUpload.F1.value = "";
					return;
				}
				else {
					document.frmUpload.action = "upload.asp";
					document.frmUpload.submit();
				}
			}
			else {
				alert("ERROR: Please select a file.")
				return;
			}
		}
		function refreshparent() {
			var ref = <%=uploads%>;
			if (ref == 1) {
				window.opener.location.reload();
				window.close();
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
	<body bgcolor='#FFFFFF' onload='refreshparent();'>
		<form method='post' name='frmUpload' enctype="multipart/form-data">
				<table border='0' align='center'>
				<tr><td colspan='2' align='center'><font color='red'  face='trebuchet MS' size='1'>&nbsp;<%=Session("MSG")%>&nbsp;</font></td></tr>
				<tr>
					<td align='center'>
						<font size='1' face='trebuchet MS'><u>New File:</u>&nbsp;<input type="file" name="F1" size="20">
							<input type='submit' value='Upload' style='width: 100px;' class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'">
					</td>
				</tr>
				<tr><td>&nbsp;</td></tr>
				<tr>
					<td align='center'>
						<input type='button' value='Close' style='width: 100px;' class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick="window.opener.location.reload();window.close();">
					</td>
				</tr>
				<tr><td colspan='2' align='center'><font face='trebuchet MS' size='1'>*2 MB limit</font></td></tr>
				<tr><td colspan='2' align='center'><font face='trebuchet MS' size='1'>*Only PDF's are allowed</font></td></tr>
			</table>
		</form>
	</body>
</html>
<%
If Session("MSG") <> "" Then
	tmpMSG = Replace(Session("MSG"), "<br>", "\n")
%>
<script><!--
	alert("<%=tmpMSG%>");
--></script>
<%
End If
Session("MSG") = ""
%>
