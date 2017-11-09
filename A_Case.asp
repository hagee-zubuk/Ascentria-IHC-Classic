<%language=vbscript%>
<!-- #include file="_Files.asp" -->

<!-- #include file="_Utils.asp" -->
<%
	If Request("Case") <> "" Then
		tmpMCNum = split(Request("Case"), " - ") 
		ID = tmpMCNum(0)
		Session("CaID") = ID
	Else
		If Request("Caid") <> "" Then Session("CaID") = Request("CaID")
	End If
	Set tblCase = Server.CreateObject("ADODB.Recordset")
	sqlCase = "SELECT * FROM Case_Manager_t WHERE [Index] = " & Session("CaID") 
	tblCase.Open sqlCase, g_strCONN, 1, 3
		If Not tblCase.EOF Then
			Index = tblCase("Index")
			fname = tblCase("Fname")
			lname = tblCase("Lname")
			Session("CMname") = lname & ", " & fname
			Agency = tblCase("Agency")
			adr = tblCase("Address")
			cty = tblCase("City")
			ste = tblCase("State")
			zcode = tblCase("Zip")
			Ono = tblCase("OfficeNo")
			Cno = tblCase("CellNo")
			Fno = tblCase("FaxNo")
			ext = tblCase("ext")
			em = tblCase("emAil")
			cmc = tblCase("cmcid")
		Else
			tblCase.AddNew
			tblCase("Index") = Session("RID")
			tblCase.Update
		End If
	tblCase.Close
	Set tblCase = Nothing	
	'''''''''''''''CMC
		Set rsPM = Server.CreateObject("ADODB.RecordSet")
		sqlPM = "SELECT * FROM CaseMngmt_T ORDER BY cmcname"
		rsPM.Open sqlPM, g_strCONN, 1, 3
		Do Until rsPM.EOF
			cmcname = rsPM("cmcname")
			Selcmc = ""
			If rsPM("cmcid") = cmc Then Selcmc = "SELECTED"
			strCMC = strCMC & "<option " & Selcmc & " value='" & rsPM("CMCid") & "' >" & cmcname & "</option>" 
			rsPM.MoveNext
		Loop
		rsPM.Close
		Set rsPM = Nothing
%>
<html>
	<head>
		<title>LSS - In-Home Care - Case Manager Details</title>
		<script language='JavaScript'>
			function CaD_Del()
			{
				var ans = window.confirm("Delete Case Manager? Click Cancel to stop.");
				if (ans){
				document.frmCaseDet.action = "A_Del.asp?act=8";
				document.frmCaseDet.submit();
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
	<body bgcolor='white' LEFTMARGIN='0' TOPMARGIN='0'>
		<form method='post' name='frmCaseDet' action='A_Ca_Action.asp'>
			<!-- #include file="_boxup.asp" -->
			<!-- #include file="_NavHeader.asp" -->
			<br>
			<center>
				<table border='0'>
					<tr>
						<td colspan='4' align='center' width='500px'>
							<font size='2' face='trebuchet MS'><b><u>Case Manager - Details</u></b></font>&nbsp;&nbsp;&nbsp;
							<font size='2' face='trebuchet MS'>[Details]</font>
							<a href="CMCon.asp?ID=<%=Index%>" style='text-decoration: none;'>
								<font color='blue' size='1' face='trebuchet MS'>[List]</font>
							</a>
						</td>
					</tr>
					
					<tr><td>&nbsp;</td></tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td>
							<font size='1' face='trebuchet MS'><u>Name:</u>&nbsp;<input style='font-size: 10px; height: 20px; width: 90px;' type='text' maxlength='20' name='lname' value="<%=Lname%>">,&nbsp; 
							<input style='font-size: 10px; height: 20px; width: 90px;' type='text' maxlength='20' name='fname' value="<%=Fname%>">
							<font size='1' face='trebuchet MS'>(Last Name, First Name)</font>
							<input type='hidden' name='MCNum' value='<%=Index%>'>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<!--<tr>
						<td>
							<font size='1' face='trebuchet MS'><u>Address:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 350px;' maxlength='49' name='Addr' value="<%=Adr%>"></font>
						</td>
						<tr>
						<td colspan='2'>
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							<font size='1' face='trebuchet MS'><u>City:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 100px;' name='cty' maxlength='49' value="<%=cty%>">
							<font size='1' face='trebuchet MS'><u>State:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 50px;' name='ste' maxlength='2' value="<%=ste%>">
							<font size='1' face='trebuchet MS'><u>Zip Code:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 100px;' name='zcode' maxlength='10' value="<%=zcode%>">
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>//-->
					<tr>
						<td>
							<font size='1' face='trebuchet MS'><u>Case Management Company:</u></font>
							<select style='font-size: 10px; height: 20px; width: 150px;' name='selCMC'>
								<option value='0'></option>
								<%=strCMC%>
							</select>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<!--<tr>
						<td><font size='1' face='trebuchet MS'><u>Agency:</u>&nbsp;<input style='font-size: 10px; height: 20px; width: 120px;' maxlength='40' type='text' name='Agency' value="<%=Agency%>"></font></td>	
					</tr>
					<tr><td>&nbsp;</td></tr>//-->
					<tr>
						<td>
							<font size='1' face='trebuchet MS'><u>Office No:</u>&nbsp;<input style='font-size: 10px; height: 20px; width: 80px;' type='text' maxlength='12' name='PhoneNo' value="<%=Ono%>"></font>
							&nbsp;
							<font size='1' face='trebuchet MS'><u>Ext:</u>&nbsp;<input style='font-size: 10px; height: 20px; width: 80px;' type='text' maxlength='12' name='Ext' value="<%=Ext%>"></font>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td><font size='1' face='trebuchet MS'><u>Cell No:</u>&nbsp;<input style='font-size: 10px; height: 20px; width: 80px;' type='text' maxlength='12' name='CelNo' value="<%=Cno%>"></font></td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td><font size='1' face='trebuchet MS'><u>Fax No:</u>&nbsp;<input style='font-size: 10px; height: 20px; width: 80px;' type='text' maxlength='12' name='FaxNo' value="<%=Fno%>"></font></td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td><font size='1' face='trebuchet MS'><u>eMail:</u>&nbsp;<input style='font-size: 10px; height: 20px; width: 120px;' type='text' maxlength='50' name='email' value="<%=Em%>"></font></td>
					</tr>
					<tr><td>&nbsp;</td></tr>
				</table>	
				<table border='0'>
					<tr>
						<td align='center' colspan='2'>
							<input type='button' value='Save' class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" style='width: 110px;' onclick='document.frmCaseDet.submit();'>
							<input type='button' value='Delete' class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" style='width: 110px;' onclick='JavaScript:CaD_Del();'>
						</td>
					</tr>
				</table>
			</center>	
			<!-- #include file="_boxdown.asp" -->
		</form>
	</body>
</html>
<% Session("CaDet") = ""
	 Session("MSG") = ""
	%>