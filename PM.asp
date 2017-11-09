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
					strTmp = "ID='" & tmpctr & "' "
					rsAdd.Find(strTmp)
					If Not rsAdd.EOF Then
						rsAdd.Delete
					End If
				End If
				rsAdd.Update
			Next	
			rsAdd.Close
			Set rsAdd = Nothing
			
		ElseIf Request("action") = 2 Then
			ctr = 0
			ctrI = Request("ctr")
			For i = 0 to ctrI 
				tmpctr = Request("chkID" & i)
				If tmpctr <> "" Then
					ctr = ctr + 1	
					tmpPM = Request("chkID" & i)
				End If
			Next	
			If ctr = 1 Then
				Set rsRPM = Server.CreateObject("ADODB.RecordSet") 'update in consumer
				sqlRPM = "UPDATE Consumer_T SET PMID = " & Request("PMRepsel") & " WHERE PMID = " & tmpPM
				rsRPM.Open sqlRPM, g_strCONN, 1, 3
				Set rsRPM = Nothing
				Set rsRPM = Server.CreateObject("ADODB.RecordSet") 'update in worker pm1
				sqlRPM = "UPDATE Worker_T SET PM1 = " & Request("PMRepsel") & " WHERE PM1 = " & tmpPM
				rsRPM.Open sqlRPM, g_strCONN, 1, 3
				Set rsRPM = Nothing
				Set rsRPM = Server.CreateObject("ADODB.RecordSet") 'update in worker pm2
				sqlRPM = "UPDATE Worker_T SET PM2 = " & Request("PMRepsel") & " WHERE PM2 = " & tmpPM
				rsRPM.Open sqlRPM, g_strCONN, 1, 3
				Set rsRPM = Nothing
				Set rsNPM = Server.CreateObject("ADODB.RecordSet")
				sqlNPM = "SELECT * FROM Proj_Man_T WHERE ID = " & tmpPM
				rsNPM.Open sqlNPM, g_strCONN, 1, 3
				If Not rsNPM.EOF Then
					name1 = rsNPM("lname") & ", " & rsNPM("fname")
				End If
				rsNPM.Close
				Set rsNPM = Nothing
				Set rsNPM = Server.CreateObject("ADODB.RecordSet")
				sqlNPM = "SELECT * FROM Proj_Man_T WHERE ID = " & Request("PMRepsel")
				rsNPM.Open sqlNPM, g_strCONN, 1, 3
				If Not rsNPM.EOF Then
					name2 = rsNPM("lname") & ", " & rsNPM("fname")
				End If
				rsNPM.Close
				Set rsNPM = Nothing
				Session("MSG") = "Replacement succeeded.<br> Consumers with " & name1 & " as Project Manger, now has " & name2 & " as their new Project Manger."
			Else
				Session("MSG") = "Please choose ONLY one(1) existing RIHCC to replace."
			End If
		Else  
			Set rsAdd = Server.CreateObject("ADODB.RecordSet")
			sqlAdd = "SELECT * FROM Proj_Man_T"
			rsAdd.Open sqlAdd, g_strCONN, 1, 3
			ctrI = Request("ctr")
			For i = 0 to ctrI 
				rsAdd.Movefirst
				tmpctr = Request("chkID" & i)
				If tmpctr <> "" Then
					strTmp = "ID='" & tmpctr & "' "
					rsAdd.Find(strTmp)
					If Not rsAdd.EOF Then
						rsAdd("lname") = Request("txtLN" & i )
						rsAdd("Fname") = Request("txtFN" & i )
					End If
				End If
				rsAdd.Update
			Next	
			If Request("txtLN") <> "" Or Request("txtFN") <> "" Then
				rsAdd.AddNew
				rsAdd("lname") = Request("txtLN")
				rsAdd("Fname") = Request("txtFN")
				rsAdd.Update
			End If
			rsAdd.Close
			Set rsAdd = Nothing	
		End If	
	End If
	Set rsPM = Server.CreateObject("ADODB.RecordSet")
	sqlPM = "SELECT * FROM Proj_Man_T ORDER BY Lname, Fname"
	rsPM.Open sqlPM, g_strCONN, 1, 3
	ctr = 0
	Do Until rsPM.EOF
		strPM = strPM & "<tr>" & vbCrLf & _
			"<td><input type='checkbox' name='chkID" & ctr & "' value='" & rsPM("ID") & "'></td>" & vbCrLf & _
			"<td><input style='font-size: 10px; height: 20px; width: 80px;' maxlength='50' name='txtLN" & ctr & "' value='" & rsPM("lname") & "'></td>" & vbCrLf & _
			"<td><input style='font-size: 10px; height: 20px; width: 80px;' maxlength='50' name='txtFN" & ctr & "' value='" & rsPM("fname") & "'></td>" & vbCrLf & _
			"</tr>" & vbCrLf 
		ctr = ctr + 1
		rsPM.MoveNext
	Loop
	rsPM.CLose
	Set rsPMRep = Nothing
	Set rsPMRep = Server.CreateObject("ADODB.RecordSet")
		sqlPMRep = "SELECT * FROM Proj_Man_T ORDER BY Lname, Fname"
		rsPMRep.Open sqlPMRep, g_strCONN, 1, 3
		Do Until rsPMRep.EOF
			PMnameRep = rsPMRep("Lname") & ", " & rsPMRep("Fname")
			strPMRep = strPMRep & "<option " & SelPM & " value='" & rsPMRep("ID") & "' >" & PMnameRep & "</option>" 
			rsPMRep.MoveNext
		Loop
	rsPMRep.Close
	Set rsPMRep = Nothing
%>
<html>
	<head>
		<title>LSS - In-Home Care - RIHCC</title>
		<link href="styles.css" type="text/css" rel="stylesheet" media="print">
		<script language="JavaScript">
			function delPM()
			{
				var ans = window.confirm("This action will delete 'checked' RIHCC. \n If you are going to delete the project manger, please replace the project manger first. \n Click OK to delete project manger. Cancel to Stop.");
				if (ans){
				document.frmPM.action = "PM.asp?action=1";
				document.frmPM.submit();
				}
			}
			function PMRep()
			{
				if (document.frmPM.PMRepsel.value == 0)
					{document.frmPM.RepPM.disabled = true;}
				else
					{document.frmPM.RepPM.disabled = false;}
			}
			function ReplacePM()
			{
				document.frmPM.action = "PM.asp?action=2";
				document.frmPM.submit();
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
	<body bgcolor='white'  LEFTMARGIN='0' TOPMARGIN='0' onload='JavaScript:PMRep();'>
		<!-- #include file="_boxup.asp" -->
		<!-- #include file="_NavHeader.asp" -->
		<br><br>
		<form method='post' name='frmPM' action='pm.asp'>
			<table align='center' border='0' cellpadding='0' cellspacing='1'>
				<tr bgcolor='#040C8B'>
					<td colspan='3' align='center' style="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=0, startColorstr=#040C8B, endColorstr=#c7c9e5);">
						<font size='1' face='trebuchet MS' color='white'><b>RIHCC<b></font></td> 
				</tr>
				<tr bgcolor='#040C8B'>
					<td style="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=0, startColorstr=#C4B464, endColorstr=#f7efc7);">&nbsp;</td>
					<td align='center' style="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=0, startColorstr=#C4B464, endColorstr=#f7efc7);">
						<font size='1' face='trebuchet MS' color='white'><b>Last Name</b></font></td> 
					<td align='center' style="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=0, startColorstr=#C4B464, endColorstr=#f7efc7);">
						<font size='1' face='trebuchet MS' color='white'><b>First Name</b></font></td>
				</tr>
				<%=strPM%>
				<tr bgcolor='#040C8B'>
					<td colspan='3' align='center' style="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=0, startColorstr=#040C8B, endColorstr=#c7c9e5);">
						<font size='1' face='trebuchet MS' color='white'><b>Replace with<b></font></td> 
				</tr>
				<tr>
					<td>&nbsp;</td>
					<td colspan='2'><select style='font-size: 10px; height: 20px; width: 165px;' name='PMRepsel' onchange='JavaScript:PMRep();'>
								<option value='0'></option>
								<%=strPMRep%>
							</select>
					</td>
				</tr>
				<tr bgcolor='#040C8B'>
					<td colspan='3' align='center' style="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=0, startColorstr=#040C8B, endColorstr=#c7c9e5);">
						<font size='1' face='trebuchet MS' color='white'><b>New<b></font></td> 
				</tr>
				<tr>
					<td>&nbsp;</td>
					<td><input style='font-size: 10px; height: 20px; width: 80px;' maxlength='50' name='txtLN' ></td>
					<td><input style='font-size: 10px; height: 20px; width: 80px;' maxlength='50' name='txtFN' ></td>
				</tr>
				<tr>
					<td colspan='3' align='center'>
						<input type='hidden' name='ctr' value='<%=ctr%>'>
						<input type='button' style='width: 91px;' class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" value='Save' onclick='document.frmPM.submit();'>
						<input type='button' style='width: 91px;' class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" value='Delete' onclick='delPM();'>
					</td>
				</tr>
				<tr>
					<td colspan='3' align='center'>
						<input type='button' name='RepPM' style='width: 186px;' class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" value='Replace RIHCC' onclick='ReplacePM();'>
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