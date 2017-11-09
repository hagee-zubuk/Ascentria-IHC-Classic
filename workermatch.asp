<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_security.asp" -->
<%
	Function Z_workername(wid)
		Z_workername = ""
		If Z_CZero(wid) = 0 Then Exit Function
		Set rsWor = Server.CreateObject("ADODB.RecordSet")
		rsWor.Open "SELECT lname, fname FROM Worker_T WHERE [index] = " & wid, g_strCONN, 3, 1
		If Not rsWor.EOF Then Z_workername = rsWor("lname") & ", " & rsWor("fname")
		rsWor.Close
		Set rsWor = Nothing
	End Function
	Function Z_consumername(cid)
		Z_consumername = ""
		If Z_CZero(cid) = 0 Then Exit Function
		Set rsCon = Server.CreateObject("ADODB.RecordSet")
		rsCon.Open "SELECT lname, fname FROM Consumer_T WHERE [index] = " & cid, g_strCONN, 3, 1
		If Not rsCon.EOF Then Z_consumername = rsCon("lname") & ", " & rsCon("fname")
		rsCon.Close
		Set rsCon = Nothing
	End Function
	Function Z_getcountyid(id, ptype)
		Z_getcounty = ""
		If Z_CZero(id) = 0 Then Exit Function
		Set rsCountID = Server.CreateObject("ADODB.RecordSet")
		rsCountID.Open "SELECT CountID FROM " & ptype & " WHERE [index] = " & id, g_strCONN, 3, 1
		If Not rsCountID.EOF Then Z_getcountyid = Z_CZero(rsCountID("countID"))
		rsCountID.Close
		Set rsCountID = Nothing
	End function
	Function Z_County(id)
		Z_County = ""
		If Z_CZero(id) = 0 Then Exit Function 
		Set rsCount = Server.CreateObject("ADODB.RecordSet")
		rsCount.Open "SELECT county FROM county_T WHERE [uid] = " & id, g_strCONN, 3, 1
		If Not rsCount.EOF Then Z_County = rsCount("county") 
		rsCount.Close
		Set rsCount = Nothing
	End Function
	Function Z_ConMedicaid(id)
		Z_ConMedicaid = ""
		If Z_CZero(id) = 0 Then Exit Function
		Set rsMed = Server.CreateObject("ADODB.RecordSet")
		rsMed.Open "SELECT medicaid_number FROM Consumer_T WHERE [index] = " & id, g_strCONN, 3, 1
		If Not rsMed.EOF Then Z_ConMedicaid = rsMed("medicaid_number")
		rsMed.Close
		Set rsMed = Nothing
	End Function
	Function Z_WorSSN(id)
		Z_WorSSN = ""
		If Z_CZero(id) = 0 Then Exit Function
		Set rsMed = Server.CreateObject("ADODB.RecordSet")
		rsMed.Open "SELECT Social_Security_Number FROM Worker_T WHERE [index] = " & id, g_strCONN, 3, 1
		If Not rsMed.EOF Then Z_WorSSN = rsMed("Social_Security_Number")
		rsMed.Close
		Set rsMed = Nothing
	End Function
	Function Z_MatchSkills(con, wor)
		Z_MatchSkills = 0
		Set rsCon = Server.CreateObject("ADODB.RecordSet")
		rsCon.Open "SELECT top 1 * FROM c_health_T WHERE medicaid_number = '" & con & "' ORDER BY [datestamp] desc", g_strCONN, 3, 1
		sklctr = 0
		sklmtch = 0
		If Not rsCon.EOF Then
			Set rsWor = Server.CreateObject("ADODB.RecordSet")
			rsWor.Open "SELECT * FROM w_skills_T WHERE wid = '" & wor & "' ", g_strCONN, 3, 1
			If Not rsWor.EOF Then
				If rsCon("housekeep") Then 
					sklctr = sklctr + 1
					If rsWor("housekeep") Then sklmtch = sklmtch + 1
				End If
				If rsCon("laundry") Then 
					sklctr = sklctr + 1
					If rsWor("laundry") Then sklmtch = sklmtch + 1
				End If
				If rsCon("meal") Then 
					sklctr = sklctr + 1
					If rsWor("meal") Then sklmtch = sklmtch + 1	
				End If
				If rsCon("grocery") Then 
					sklctr = sklctr + 1
					If rsWor("grocery") Then sklmtch = sklmtch + 1	
				End If
				If rsCon("dress") Then 
					sklctr = sklctr + 1
					If rsWor("dress") Then sklmtch = sklmtch + 1	
				End If
				If rsCon("eat") Then 
					sklctr = sklctr + 1
					If rsWor("eat") Then sklmtch = sklmtch + 1	
				End If
				If rsCon("asstwalk") Then 
					sklctr = sklctr + 1
					If rsWor("asstwalk") Then sklmtch = sklmtch + 1	
				End If
				If rsCon("asstwheel") Then 
					sklctr = sklctr + 1
					If rsWor("asstwheel") Then sklmtch = sklmtch + 1	
				End If
				If rsCon("asstmotor") Then 
					sklctr = sklctr + 1
					If rsWor("asstmotor") Then sklmtch = sklmtch + 1	
				End If
				If rsCon("commeal") Then 
					sklctr = sklctr + 1
					If rsWor("commeal") Then sklmtch = sklmtch + 1	
				End If
				If rsCon("medical") Then 
					sklctr = sklctr + 1
					If rsWor("medical") Then sklmtch = sklmtch + 1	
				End If
				If rsCon("shower") Then 
					sklctr = sklctr + 1
					If rsWor("shower") Then sklmtch = sklmtch + 1	
				End If
				If rsCon("tub") Then 
					sklctr = sklctr + 1
					If rsWor("tub") Then sklmtch = sklmtch + 1	
				End If
				If rsCon("oral") Then 
					sklctr = sklctr + 1
					If rsWor("oral") Then sklmtch = sklmtch + 1	
				End If
				If rsCon("commode") Then 
					sklctr = sklctr + 1
					If rsWor("commode") Then sklmtch = sklmtch + 1	
				End If
				If rsCon("sit") Then 
					sklctr = sklctr + 1
					If rsWor("sit") Then sklmtch = sklmtch + 1	
				End If
				If rsCon("medication") Then 
					sklctr = sklctr + 1
					If rsWor("medication") Then sklmtch = sklmtch + 1	
				End If
				If rsCon("undress") Then 
					sklctr = sklctr + 1
					If rsWor("undress") Then sklmtch = sklmtch + 1	
				End If
				If rsCon("shampoosink") Then 
					sklctr = sklctr + 1
					If rsWor("shampoosink") Then sklmtch = sklmtch + 1	
				End If
				If rsCon("oralcare") Then 
					sklctr = sklctr + 1
					If rsWor("oralcare") Then sklmtch = sklmtch + 1	
				End If
				If rsCon("massage") Then 
					sklctr = sklctr + 1
					If rsWor("massage") Then sklmtch = sklmtch + 1	
				End If
				If rsCon("shampoobed") Then 
					sklctr = sklctr + 1
					If rsWor("shampoobed") Then sklmtch = sklmtch + 1	
				End If
				If rsCon("shave") Then 
					sklctr = sklctr + 1
					If rsWor("shave") Then sklmtch = sklmtch + 1	
				End If
				If rsCon("bedbath") Then 
					sklctr = sklctr + 1
					If rsWor("bedbath") Then sklmtch = sklmtch + 1	
				End If
				If rsCon("bedpan") Then 
					sklctr = sklctr + 1
					If rsWor("bedpan") Then sklmtch = sklmtch + 1	
				End If
				If rsCon("ptexer") Then 
					sklctr = sklctr + 1
					If rsWor("ptexer") Then sklmtch = sklmtch + 1	
				End If
				If rsCon("hoyer") Then 
					sklctr = sklctr + 1
					If rsWor("hoyer") Then sklmtch = sklmtch + 1	
				End If
				If rsCon("eye") Then 
					sklctr = sklctr + 1
					If rsWor("eye") Then sklmtch = sklmtch + 1	
				End If
				If rsCon("transferbelt") Then 
					sklctr = sklctr + 1
					If rsWor("transferbelt") Then sklmtch = sklmtch + 1	
				End If
				If rsCon("alz") Then 
					sklctr = sklctr + 1
					If rsWor("alz") Then sklmtch = sklmtch + 1	
				End If
				If rsCon("Incontinence") Then 
					sklctr = sklctr + 1
					If rsWor("Incontinence") Then sklmtch = sklmtch + 1	
				End If
				If rsCon("Hospice") Then 
					sklctr = sklctr + 1
					If rsWor("Hospice") Then sklmtch = sklmtch + 1	
				End If
				If rsCon("lotion") Then 
					sklctr = sklctr + 1
					If rsWor("lotion") Then sklmtch = sklmtch + 1	
				End If
			End If
			rsWor.Close
			Set rsWor = Nothing
		End If
		rsCon.Close
		Set rsCon = Nothing
		If sklctr > 0 Then Z_MatchSkills = sklmtch / sklctr
	End Function
	
	Dim arrWor(), arrPer(), arrWor2(), arrper2()
	
	consumername = Z_consumername(Request("cid"))
	county = Z_County(Z_getcountyid(Request("cid"), "Consumer_T"))
	
	'same county
	Set rsMatch = Server.CreateObject("ADODB.RecordSet")
	sqlMatch = "SELECT lname, fname, Social_Security_Number, [index] FROM worker_T WHERE [status] = 'Active' AND countID = " & _
		Z_getcountyid(Request("cid"), "Consumer_T") & " Order By lname, fname"
	rsMatch.Open sqlMatch, g_strCONN, 3, 1
	x = 0
	Do Until rsMatch.EOF
		conMatch = Z_MatchSkills(Z_ConMedicaid(Request("cid")), rsMatch("Social_Security_Number"))
		ReDim Preserve arrWor(x)
		ReDim Preserve arrPer(x)
		If conMatch >= 0.75 Then
			arrWor(x) = rsMatch("index")
			arrPer(x) = conMatch
			x = x + 1
		End If
		rsMatch.MoveNext
	Loop
	rsMatch.Close
	Set rsMatch = Nothing
	
	If x > 0 Then
		n = UBound(arrWor)
		Do
		  nn = -1
		  For j = LBound(arrWor) to n - 1
		      If arrPer(j) < arrPer(j + 1) Then
		         TempValue = arrWor(j + 1)
		         arrWor(j + 1) = arrWor(j)
		         arrWor(j) = TempValue
		         TempValue2 = arrPer(j + 1)
		         arrPer(j + 1) = arrPer(j)
		         arrPer(j) = TempValue2
		         nn = j
		      End If
		  Next
		  n = nn
		Loop Until nn = -1
		For i = LBound(arrWor) To UBound(arrWor)
			strWork = strWork & "<tr><td align='center'><font size='1' face='trebuchet ms'><a href='#' onclick=""PassMe('" & Z_WorSSN(arrWor(i)) & "');"">" & Z_workername(arrWor(i)) & _
				"</a></font></td><td align='center'><font size='1' face='trebuchet ms'>" & Z_County(Z_getcountyid(arrWor(i), "Worker_T")) & _
				"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & FormatPercent(arrPer(i), 2) & _
				"</font></td></tr>"
		Next 
	End If
	
	'not same 
	Set rsMatch = Server.CreateObject("ADODB.RecordSet")
	sqlMatch = "SELECT lname, fname, Social_Security_Number, [index] FROM worker_T WHERE [status] = 'Active' AND countID <> " & Z_getcountyid(Request("cid"), "Consumer_T") & " Order By lname, fname"
	rsMatch.Open sqlMatch, g_strCONN, 3, 1
	x = 0
	Do Until rsMatch.EOF
		conMatch = Z_MatchSkills(Z_ConMedicaid(Request("cid")), rsMatch("Social_Security_Number"))
		ReDim Preserve arrWor2(x)
		ReDim Preserve arrPer2(x)
		If conMatch >= 0.75 Then
			arrWor2(x) = rsMatch("index")
			arrPer2(x) = conMatch
			x = x + 1
		End If
		rsMatch.MoveNext
	Loop
	rsMatch.Close
	Set rsMatch = Nothing
	
	If x > 0 Then
		n = UBound(arrWor2)
		If x > 0 Then strWork = strWork & "<tr><td align='left' colspan='3'><font size='1' face='trebuchet ms'>*Not in the same county</font></td></tr>"
		Do
		  nn = -1
		  For j = LBound(arrWor2) to n - 1
		      If arrPer2(j) < arrPer2(j + 1) Then
		         TempValue = arrWor2(j + 1)
		         arrWor2(j + 1) = arrWor2(j)
		         arrWor2(j) = TempValue
		         TempValue2 = arrPer2(j + 1)
		         arrPer2(j + 1) = arrPer2(j)
		         arrPer2(j) = TempValue2
		         nn = j
		      End If
		  Next
		  n = nn
		Loop Until nn = -1
		For i = LBound(arrWor2) To UBound(arrWor2)
			strWork = strWork & "<tr><td align='center'><font size='1' face='trebuchet ms'>" & Z_workername(arrWor2(i)) & _
				"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & Z_County(Z_getcountyid(arrWor2(i), "Worker_T")) & _
				"</font></td><td align='center'><font size='1' face='trebuchet ms'>" & FormatPercent(arrPer2(i), 2) & _
				"</font></td></tr>"
		Next 
	End If
%>
<html>
	<head>
		<title>In Home Care - Find Worker</title>
		<script language='JavaScript'>
		function PassMe(xxx){
			//alert(xxx);
			window.opener.document.frmConDet.action = "a_worker.asp?wid=" + xxx;
			window.opener.document.frmConDet.submit();
			//window.opener.location.href='a_worker.asp?wid=' + xxx;
			self.close();
		}
		</script>
	</head>
	<body>
		<form method='post' name='frmWorkmatch'>
			<table width='100%'>
				<tr><td colspan='2' align='left'><font size='4' color='#040c8b'><b><i>----Match Worker</i></b></font></td></tr>
				<tr>
					<td width='75px'>Consumer:</td>
					<td><%=consumername%></td>
				</tr>
				<tr>
					<td>County:</td>
					<td><%=county%></td>
				</tr>
				<tr><td>&nbsp;</td></tr>
				<tr>
					<td colspan='2' align='center'>
						<table border='1' width='100%'>
							<tr bgcolor='#040c8b'>
								<td style="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=0, startColorstr=#040C8B, endColorstr=#c7c9e5);" align='center'><font face='trebuchet MS' size='2' color='white'><b>PCSP Worker</b></font></td>
								<td style="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=0, startColorstr=#040C8B, endColorstr=#c7c9e5);" align='center'><font face='trebuchet MS' size='2' color='white'><b>County</b></font></td>
								<td style="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=0, startColorstr=#040C8B, endColorstr=#c7c9e5);" align='center'><font face='trebuchet MS' size='2' color='white'><b>Match</b></font></td>
							</tr>
							<%=strWork%>
						</table>
					</td>
				</tr>
			</table>
		</form>
	</body>
</html>
