<%Language=VBScript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_security.asp" -->
<%
	If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
		If Request("action") = 1 Then
			Set rsAdd = Server.CreateObject("ADODB.RecordSet")
			sqlAdd = "SELECT * FROM Proj_Man_T"
			rsAdd.Open sqlAdd, g_strCONN, 1, 3
			ctrI = Request("ctr")
			For i = 0 to ctrI 
				rsAdd.Movefirst
				tmpctr = Request("chkID" & i)
				If tmpctr <> "" Then
					strTmp = "mccid='" & tmpctr & "' "
					rsAdd.Find(strTmp)
					If Not rsAdd.EOF Then
						rsAdd.Delete
					End If
				End If
				rsAdd.Update
			Next	
			rsAdd.Close
			Set rsAdd = Nothing
		Else  
			Set rsAdd = Server.CreateObject("ADODB.RecordSet")
			sqlAdd = "SELECT * FROM ManagedCare_T"
			rsAdd.Open sqlAdd, g_strCONN, 1, 3
			ctrI = Request("ctr")
			if not rsadd.eof then
				For i = 0 to ctrI 
					rsAdd.Movefirst
					tmpctr = Request("chkID" & i)
					If tmpctr <> "" Then
						strTmp = "MCCid='" & tmpctr & "' "
						rsAdd.Find(strTmp)
						If Not rsAdd.EOF Then
							rsAdd("mccname") = Request("txtLN" & i )
						End If
					End If
					rsAdd.Update
				Next	
			end if
			If Request("txtLN") <> "" Then
				rsAdd.AddNew
				rsAdd("mccname") = Request("txtLN")
				rsAdd.Update
			End If
			rsAdd.Close
			Set rsAdd = Nothing	
		End If	
	End If
	Set rsPM = Server.CreateObject("ADODB.RecordSet")
	sqlPM = "SELECT * FROM ManagedCare_T ORDER BY MCCname"
	rsPM.Open sqlPM, g_strCONN, 1, 3
	ctr = 0
	Do Until rsPM.EOF
		strPM = strPM & "<tr>" & vbCrLf & _
			"<td><input type='checkbox' name='chkID" & ctr & "' value='" & rsPM("MCCid") & "'></td>" & vbCrLf & _
			"<td><input style='font-size: 10px; height: 20px; width: 150px;' maxlength='50' name='txtLN" & ctr & "' value='" & rsPM("MCCname") & "'></td>" & vbCrLf & _
			"</tr>" & vbCrLf 
		ctr = ctr + 1
		rsPM.MoveNext
	Loop
	rsPM.CLose
	Set rsPMRep = Nothing
%>
<html>
	<head>
		<title>LSS - In-Home Care - Managed Care Company</title>
		<link href="styles.css" type="text/css" rel="stylesheet" media="print">
		<script language="JavaScript">
			function delPM()
			{
				var ans = window.confirm("This action will delete 'checked' Managed Care Company.\n Click OK to delete Managed Care Company. Cancel to Stop.");
				if (ans){
				document.frmPM.action = "mcc.asp?action=1";
				document.frmPM.submit();
				}
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
		<br><br>
		<form method='post' name='frmPM' action='mcc.asp'>
			<table align='center' border='0' cellpadding='0' cellspacing='1'>
				<tr bgcolor='#040C8B'>
					<td colspan='2' align='center' style="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=0, startColorstr=#040C8B, endColorstr=#c7c9e5);">
						<font size='1' face='trebuchet MS' color='white'><b>Managed Care Company<b></font></td> 
				</tr>
				<tr bgcolor='#040C8B'>
					<td style="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=0, startColorstr=#C4B464, endColorstr=#f7efc7);">&nbsp;</td>
					<td align='center' style="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=0, startColorstr=#C4B464, endColorstr=#f7efc7);">
						<font size='1' face='trebuchet MS' color='white'><b>Name</b></font></td> 
				</tr>
				<%=strPM%>
				<tr bgcolor='#040C8B'>
					<td colspan='3' align='center' style="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=0, startColorstr=#040C8B, endColorstr=#c7c9e5);">
						<font size='1' face='trebuchet MS' color='white'><b>New<b></font></td> 
				</tr>
				<tr>
					<td>&nbsp;</td>
					<td><input style='font-size: 10px; height: 20px; width: 150px;' maxlength='50' name='txtLN' ></td>
				</tr>
				<tr>
					<td colspan='3' align='center'>
						<input type='hidden' name='ctr' value='<%=ctr%>'>
						<input type='button' style='width: 91px;' class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" value='Save' onclick='document.frmPM.submit();'>
						<input type='button' style='width: 91px;' class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" value='Delete' onclick='delPM();'>
					</td>
				</tr>
			</table>
			<center>
				<span><font color='red' size='1' face='trebuchet MS'><%=Session("MSG")%></font></span>
			</center>
		</form>
		<!-- #include file="_boxdown.asp" -->
	</body>
</html>
<%Session("MSG") = "" %>