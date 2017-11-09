<%language=vbscript%>
<!-- #include file="_Files.asp" -->

<!-- #include file="_Utils.asp" -->
<!-- #include file="_security.asp" -->
<%Response.AddHeader "Pragma", "No-Cache" %>
<%
Function GetCon(xxx)
		GetCon = "N/A"
		Set rsCon = Server.CreateObject("ADODB.RecordSet")
		sqlCon = "SELECT * FROM Consumer_T WHERE Medicaid_Number = '" & xxx & "' "
		rsCon.Open sqlCon, g_strCONN, 3, 1
		If Not rsCon.EOF Then
			GetCon = rsCon("Lname") & ", " & rsCon("Fname")
		End If
		rsCon.Close
		Set rsCon = Nothing
	End Function
	Function GetWor(xxx)
		GetWor = "N/A"
		Set rsWor = Server.CreateObject("ADODB.RecordSet")
		sqlWor = "SELECT * FROM Worker_T WHERE Social_Security_Number = '" & xxx & "' "
		rsWor.Open sqlWor, g_strCONN, 3, 1
		If Not rsWor.EOF Then
			GetWor = rsWor("Lname") & ", " & rsWor("Fname")
		End If
		rsWor.Close
		Set rsWor = Nothing
	End Function
	Function GetUser(xxx)
		GetUser = "N/A"
		Set rsUser = Server.CreateObject("ADODB.RecordSet")
		sqlUser = "SELECT * FROM Input_T WHERE [index] = " & xxx 
		rsUser.Open sqlUser, g_strCONN, 3, 1
		If Not rsUser.EOF Then
			GetUser = rsUser("Lname") & ", " & rsUser("Fname")
		End If
		rsUser.Close
		Set rsUser = Nothing
	End Function
	If UCase(Session("lngType")) = "0" Then
		Session("MSG") = "Invalid User Type. Please Sign In again."
		Response.Redirect "default.asp"
	End If
	If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
		'CHANGE MEDICAID
		If Request("selMed") <> "-1" And Request("txtNewMed") <> "" Then
			Set fso = CreateObject("Scripting.FileSystemObject")
			Set ALog = fso.OpenTextFile(AdminLog, 8, True)
			Alog.WriteLine Now & ":: Consumer's Medicaid Number (" & Request("HidMed") & " - " & GetCon(Request("HidMed")) & _
				") has been changed to " & Request("txtNewMed") & " by " & GetUser(Session("UserID"))
			Set Alog = Nothing
			Set fso = Nothing 
			Set rsMedDiag = Server.CreateObject("ADODB.RecordSet")
			sqlMedDiag = "UPDATE C_Diagnosis_t SET Medicaid_Number = '" & Request("txtNewMed") & "' WHERE Medicaid_Number = '" & Request("HidMed") & "' "
			rsMedDiag.Open sqlMedDiag, g_strCONN, 1, 3
			Set rsMedDiag = Nothing
			Set rsMedFiles = Server.CreateObject("ADODB.RecordSet")
			sqlMedFiles = "UPDATE C_Files_t SET Medicaid_Number = '" & Request("txtNewMed") & "' WHERE Medicaid_Number = '" & Request("HidMed") & "' "
			rsMedFiles.Open sqlMedFiles, g_strCONN, 1, 3
			Set rsMedFiles = Nothing
			Set rsMedHealth = Server.CreateObject("ADODB.RecordSet")
			sqlMedHealth = "UPDATE C_Health_t SET Medicaid_Number = '" & Request("txtNewMed") & "' WHERE Medicaid_Number = '" & Request("HidMed") & "' "
			rsMedHealth.Open sqlMedHealth, g_strCONN, 1, 3
			Set rsMedHealth = Nothing
			Set rsMedLog = Server.CreateObject("ADODB.RecordSet")
			sqlMedLog = "UPDATE C_Site_Visit_Dates_t SET Medicaid_Number = '" & Request("txtNewMed") & "' WHERE Medicaid_Number = '" & Request("HidMed") & "' "
			rsMedLog.Open sqlMedLog, g_strCONN, 1, 3
			Set rsMedLog = Nothing
			Set rsMedStatus = Server.CreateObject("ADODB.RecordSet")
			sqlMedStatus = "UPDATE C_Status_t SET Medicaid_Number = '" & Request("txtNewMed") & "' WHERE Medicaid_Number = '" & Request("HidMed") & "' "
			rsMedStatus.Open sqlMedStatus, g_strCONN, 1, 3
			Set rsMedStatus = Nothing
			Set rsMedCM = Server.CreateObject("ADODB.RecordSet")
			sqlMedCM = "UPDATE CMCon_t SET CID = '" & Request("txtNewMed") & "' WHERE CID = '" & Request("HidMed") & "' "
			rsMedCM.Open sqlMedCM, g_strCONN, 1, 3
			Set rsMedCM = Nothing
			Set rsMedRep = Server.CreateObject("ADODB.RecordSet")
			sqlMedRep = "UPDATE ConRep_t SET CID = '" & Request("txtNewMed") & "' WHERE CID = '" & Request("HidMed") & "' "
			rsMedRep.Open sqlMedRep, g_strCONN, 1, 3
			Set rsMedRep = Nothing
			Set rsMedCon = Server.CreateObject("ADODB.RecordSet")
			sqlMedCon = "UPDATE Consumer_t SET Medicaid_Number = '" & Request("txtNewMed") & "' WHERE Medicaid_Number = '" & Request("HidMed") & "' "
			rsMedCon.Open sqlMedCon, g_strCONN, 1, 3
			Set rsMedCon = Nothing
			Set rsMedWor = Server.CreateObject("ADODB.RecordSet")
			sqlMedWor = "UPDATE ConWork_t SET CID = '" & Request("txtNewMed") & "' WHERE CID = '" & Request("HidMed") & "' "
			rsMedWor.Open sqlMedWor, g_strCONN, 1, 3
			Set rsMedWor = Nothing
			Set rsMedTS = Server.CreateObject("ADODB.RecordSet")
			sqlMedTS = "UPDATE Tsheets_t SET Client = '" & Request("txtNewMed") & "' WHERE Client = '" & Request("HidMed") & "' "
			rsMedTS.Open sqlMedTS, g_strCONN, 1, 3
			Set rsMedTS = Nothing
			Session("MSG") = "Consumer's Medicaid Number has been updated.<br>"
		End If	
		If Request("selSSN") <> "-1" And Request("txtNewSSN") <> "" Then
			Set fso = CreateObject("Scripting.FileSystemObject")
			Set ALog = fso.OpenTextFile(AdminLog, 8, True)
			Alog.WriteLine Now & ":: PCSP Worker's Social Security Number (" & Request("HidSSN") & " - " & GetWor(Request("HidSSN")) & _
				") has been changed to " & Request("txtNewSSN") & " by " & GetUser(Session("UserID"))
			Set Alog = Nothing
			Set fso = Nothing 
			Set rsSSNTS = Server.CreateObject("ADODB.RecordSet")
			sqlSSNTS = "UPDATE Tsheets_t SET emp_id = '" & Request("txtNewSSN") & "' WHERE emp_id = '" & Request("HidSSN") & "' "
			rsSSNTS.Open sqlSSNTS, g_strCONN, 1, 3
			Set rsSSNTS = Nothing
			Set rsSSNFiles = Server.CreateObject("ADODB.RecordSet")
			sqlSSNFiles = "UPDATE W_Files_t SET SSN = '" & Request("txtNewSSN") & "' WHERE SSN = '" & Request("HidSSN") & "' "
			rsSSNFiles.Open sqlSSNFiles, g_strCONN, 1, 3
			Set rsSSNFiles = Nothing
			Set rsSSNLog = Server.CreateObject("ADODB.RecordSet")
			sqlSSNLog = "UPDATE W_log_T SET SSN = '" & Request("txtNewSSN") & "' WHERE SSN = '" & Request("HidSSN") & "' "
			rsSSNLog.Open sqlSSNLog, g_strCONN, 1, 3
			Set rsSSNLog = Nothing
			Set rsSSNTowns = Server.CreateObject("ADODB.RecordSet")
			sqlSSNTowns = "UPDATE W_Towns_T SET SSN = '" & Request("txtNewSSN") & "' WHERE SSN = '" & Request("HidSSN") & "' "
			rsSSNTowns.Open sqlSSNLog, g_strCONN, 1, 3
			Set rsSSNTowns = Nothing
			Set rsSSNWor = Server.CreateObject("ADODB.RecordSet")
			sqlSSNWor = "UPDATE Worker_t SET Social_Security_Number = '" & Request("txtNewSSN") & "' WHERE Social_Security_Number = '" & Request("HidSSN") & "' "
			rsSSNWor.Open sqlSSNWor, g_strCONN, 1, 3
			Set rsSSNWor = Nothing
			Session("MSG") = Session("MSG") & "PCSP Worker's Social Security Number has been updated."
		End If
	End If
	Set rsMed = Server.CreateObject("ADODB.RecordSet")
	sqlMed = "SELECT * FROM Consumer_T ORDER BY lname, Fname"
	rsMed.Open sqlMed, g_strCONN, 3, 1
	Do Until rsMed.EOF
		myName = rsMed("Lname") & ", " & rsMed("Fname")
		strMed = strMed & "<option value='" & rsMed("index") & "'>" & rsMEd("Medicaid_Number") & " - " & myName & "</option>" & vbCrLf
		tmpMed = Replace(Trim(rsMed("Medicaid_Number")), "-", "")
		tmpMed = Replace(tmpMed, "/", "")
		jsMed = jsMed & "if (document.frmAdmin.txtNewMed.value == """ & tmpMed & """) " & vbCrLf & _
			"{alert(""Medicaid Number already exists."");return;} " & vbCrlf
		myMed = myMed & "if (xxx == " & rsMed("Index") & ") " & vbCrLf & _
			"{document.frmAdmin.HidMed.value = """ & rsMEd("Medicaid_Number") &""";} " & vbCrLf
		rsMed.MoveNext
	Loop
	rsMed.Close
	Set rsMEd = Nothing
	Set rsSSN = Server.CreateObject("ADODB.RecordSet")
	sqlSSN = "SELECT * FROM Worker_t ORDER BY lname, Fname"
	rsSSN.Open sqlSSN, g_strCONN, 3, 1
	Do Until rsSSN.EOF
		myName = rsSSN("Lname") & ", " & rsSSN("Fname")
		strSSN = strSSN & "<option value='" & rsSSN("index") & "'>" & rsSSN("Social_Security_Number") & " - " & myName & "</option>" & vbCrLf
		'tmpSSN = Replace(Trim(rsSSN("Social_Security_Number")), "-", "")
		'tmpSSN = Replace(tmpSSN, "/", "")
		tmpSSN = Trim(rsSSN("Social_Security_Number"))
		jsSSN = jsSSN & "if (document.frmAdmin.txtNewSSN.value == """ & tmpSSN & """) " & vbCrLf & _
			"{alert(""Social Security Number already exists."");return;} " & vbCrlf
		mySSN = mySSN & "if (yyy == " & rsSSN("index") & ") " & vbCrLf & _
			"{document.frmAdmin.HidSSN.value = """ & rsSSN("Social_Security_Number") &""";} " & vbCrLf
		rsSSN.MoveNext
	Loop
	rsSSN.Close
	Set rsSSN = Nothing
%>
<html>
	<head>
		<title>LSS - In-Home Care - Administrator Tools - Database Tools</title>
		<script language='JavaScript'>
		<!--
		function myTemp(xxx,yyy)
		{
			<%=myMed%>
			<%=mySSN%>
		}
		function checkMe()
		{
			if (document.frmAdmin.txtNewMed.value == "" && document.frmAdmin.selMed.value !== "-1")
			{
				alert("Please enter the new Medicaid Number for consumer.");
				return;
			}
			if (document.frmAdmin.txtNewMed.value !== "" && document.frmAdmin.selMed.value == "-1")
			{
				alert("Please select the consumer you wish to change.");
				return;
			}
			if (document.frmAdmin.txtNewMed.value !== "" && document.frmAdmin.selMed.value !== "-1")
			{
				<%=jsMed%>
			}
			if (document.frmAdmin.txtNewSSN.value == "" && document.frmAdmin.selSSN.value !== "-1")
			{
				alert("Please enter the new Social Security Number for worker.");
				return;	
			}
			if (document.frmAdmin.txtNewSSN.value !== "" && document.frmAdmin.selSSN.value == "-1")
			{
				alert("Please select the PCSP worker you wish to change.");
				return;
			}
			if (document.frmAdmin.txtNewSSN.value !== "" && document.frmAdmin.selSSN.value !== "-1")
			{
				<%=jsSSN%>
			}
			strMSG = "You are about to EDIT the following:"
			if (document.frmAdmin.txtNewMed.value !== "")
			{
				strMSG = strMSG + "\nMedicaid Number: " + document.frmAdmin.HidMed.value
			}
			if (document.frmAdmin.txtNewSSN.value !== "")
			{
				strMSG = strMSG + "\nSocial Security Number: " + document.frmAdmin.HidSSN.value
			}
			strMSG = strMSG + "\n\n* Change will reflect in ALL database entries."
			var ans = window.confirm(strMSG);
			if (ans)
			{
				document.frmAdmin.submit();
			}
		}
		function maskMe(str,textbox,loc,delim)
		{
			var locs = loc.split(',');
			for (var i = 0; i <= locs.length; i++)
			{
				for (var k = 0; k <= str.length; k++)
				{
					 if (k == locs[i])
					 {
						if (str.substring(k, k+1) != delim)
					 	{
					 		str = str.substring(0,k) + delim + str.substring(k,str.length);
		     			}
					}
				}
		 	}
			textbox.value = str
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
	<body bgcolor='white'  LEFTMARGIN='0' TOPMARGIN='0' onload='myTemp(document.frmAdmin.selMed.value, document.frmAdmin.selSSN.value);'>
		<!-- #include file="_boxup.asp" -->
		<!-- #include file="_NavHeader.asp" -->
		<form name='frmAdmin' method='post' action='changeprime.asp'>
			<table border='0' cellspacing='0' cellpadding='0' align='center'>
				<tr><td>&nbsp;</td></tr>
				<tr bgcolor='#040C8B'>
					<td colspan='3' align='center' style="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=0, startColorstr=#040C8B, endColorstr=#c7c9e5);">
						<font face='trebuchet MS' size='2' color='white'><b>Database Tools</b></font>
					</td>
				</tr>
				<tr>
					<td>
						<font size='1' face='trebuchet MS'>Consumer:</font>
					</td>
					<td>
						<font size='1' face='trebuchet MS'>New Medicaid:</font>
					</td>
				</tr>
				<tr>
					<td>
						<select name='selMed' onchange='myTemp(document.frmAdmin.selMed.value, document.frmAdmin.selSSN.value);'>
							<option value='-1'>&nbsp;</option>
							<%=strMed%>
						</select>
						<input type='hidden' name='HidMed'>
					</td>
					<td>
						<input type="text" size='14' maxlength='14' name='txtNewMed'>
					</td>
				</tr>
				<tr><td>&nbsp;</td></tr>
				<tr>
					<td>
						<font size='1' face='trebuchet MS'>PCSP Worker:</font>
					</td>
					<td>
						<font size='1' face='trebuchet MS'>New SSN:</font>
					</td>
				</tr>
				<tr>
					<td>
						<select name='selSSN' onchange='myTemp(document.frmAdmin.selMed.value, document.frmAdmin.selSSN.value);'>
							<option value='-1'>&nbsp;</option>
							<%=strSSN%>
						</select>
						<input type='hidden' name='HidSSN'>
					</td>
					<td>
						<input type="text" size='14' maxlength='11' name='txtNewSSN' onKeyUp="javascript:return maskMe(this.value,this,'3,6','-');" onBlur="javascript:return maskMe(this.value,this,'3,6','-');">
					</td>
				</tr>
				<tr><td>&nbsp;</td></tr>
				<tr>
					<td colspan='2'>
						<input type="button" value='Save' style='width: 150px;' class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='checkMe();'>
					</td>
				</tr>
				<tr>
					<td colspan='2'>
						<span><font face='Trebuchet MS' size='1' color='red'><%=Session("MSG")%></font></span>
					</td>
				</tr>
			</table>
		</form>
		<!-- #include file="_boxdown.asp" -->
	</body>
</html>
<% Session("MSG") = "" %>