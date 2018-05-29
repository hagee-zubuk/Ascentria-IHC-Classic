<%language=vbscript%>
<!-- #include file="_Files.asp" -->

<!-- #include file="_Utils.asp" -->
<!-- #include file="_security.asp" -->
<%
	Function GetWID(tmpLine)
		SplitLine = Split(tmpLine, "|")
		GetWID = SplitLine(0)
	End Function
	Function FixLine(tmpLine)
		SplitLine = Split(tmpLine, "|")
		FixLine = SplitLine(1)
	End Function
	FinOnly = "ReadOnly"
	If Session("lngType") = 1 Or Session("lngType") = 2 Or session("UserID") = 2 Then
		FinOnly = ""
	End If
	'If Request("consumer") <> "" Then
	'	tmpMCNum = split(Request("consumer"), " - ") 
	'	ID = tmpMCNum(0)
	'	Session("CID") = Request("consumer")
	'Else
	'	If Request("CID") <> "" Then Session("CID") = Request("CID")
	'End If
	If Request("consumer") <> "" Then
		CID = Request("consumer")
	ElseIf Request("MNum") <> "" Then
		CID = Request("MNum") 
	Else
		Session("MSG") = "Sorry. To maintain integrity of database, please Sign-in again."
		Response.Redirect "default.asp"
	End If
	If CID <> "" Then
		Set tblConsumer = Server.CreateObject("ADODB.Recordset")
		sqlConsumer = "SELECT * FROM Consumer_t WHERE Medicaid_Number = '" & CID & "' "
		tblConsumer.Open sqlConsumer, g_strCONN, 1, 3
			If Not tblConsumer.EOF Then
				Index = tblConsumer("Index")
				Session("Cidx") = Index
				MCNum = tblConsumer("Medicaid_Number")
				SSN = tblConsumer("Social_Security_Number")
				Lname = tblConsumer("Lname")
				Fname = tblConsumer("Fname")
				Session("Cname") = lname & ", " & fname
				DOB = tblConsumer("DOB")
				If tblConsumer("Gender")= "Male" Then
					M = "selected"
					F = ""
				ElseIf tblConsumer("Gender")= "Female" Then
					M = ""
					F = "selected"
				End If
				Addr = tblConsumer("Address")
				cty = tblConsumer("City")
				ste = tblConsumer("State")
				zcode = tblConsumer("Zip")
				gmapsaddr = Replace(Addr, "#", "") & ", " & cty & ", " & ste & ", " & zcode  
				MAddr = tblConsumer("mAddress")
				Mcty = tblConsumer("mCity")
				Mste = tblConsumer("mState")
				Mzcode = tblConsumer("mZip")
				FonNum = tblConsumer("PhoneNo")
				FonNum2 = tblConsumer("secphone")
				celNum = tblConsumer("celphone")
				rela = ""
				If tblConsumer("workrelative") Then rela = "checked"
				emerinfo = tblConsumer("emerinfo")
				emerrel = tblConsumer("emerrel")
				emerphone = tblConsumer("emerphone")
				Direct = tblConsumer("Directions")
				If tblConsumer("Driving") = True Then Drive = "checked"
				Ref_Dte = tblConsumer("Referral_Date")
				Strt_Dte = tblConsumer("Start_Date")
				EndDte = tblConsumer("End_Date")
				termDte = tblConsumer("TermDate")
				If tblConsumer("Representative") = True Then Rep = "checked"
				maxhrs = tblConsumer("MaxHrs")
				cmt = tblConsumer("Pcomment")
				AmendS = tblConsumer("AmendSigned")
				AmendR = tblConsumer("DteRecd")
				'PM = tblConsumer("PM")
				PM = tblConsumer("PMID")
				CareDte = tblConsumer("CarePlan")
				EffDte = tblConsumer("EffDate")
				milecap = tblConsumer("milecap")
				email = tblConsumer("email")
				code = tblConsumer("code")
				mcode = ""
				pcode = ""
				ccode = ""
				acode = ""
				vcode = ""
				If code = "M" Then mcode = "selected"
				If code = "P" Then pcode = "selected"
				If code = "C" Then ccode = "selected"
				If code = "A" Then acode = "selected"
				If code = "V" Then vcode = "selected"	
				contract = tblConsumer("contract") 
				tmprate = tblConsumer("rate")
				cliid = tblConsumer("cliid")
				aphone1 = tblConsumer("aphone1")
				aphone2 = tblConsumer("aphone2")
				aphone3 = tblConsumer("aphone3")
				aphone4 = tblConsumer("aphone4")
				aphone5 = tblConsumer("aphone5")
				percode= tblConsumer("pcode")
				vahm = ""
				If tblconsumer("vahm") Then vahm = "checked"
				hrshm = tblconsumer("vahmhrs")
				vaha = ""
				If tblconsumer("vaha") Then vaha = "checked"
				hrsha = tblconsumer("vahahrs")
				concounty = tblConsumer("countid")
				' new -- 2017-10-27
				langsel = Z_ListLanguages(tblConsumer("langid"))
			Else
				Session("MSG") = "Session has expired. Please sign in again."
				Response.Redirect "Default.asp"
			End If
		tblConsumer.Close
		Set tblConsumer = Nothing
		Set tblRsel = Server.CreateObject("ADODB.RecordSet")
		sqlRsel = "SELECT * FROM ConRep_t WHERE CID = '" & CID & "' "
		tblRsel.Open sqlRsel, g_strCONN, 1, 3
		If Not tblRsel.EOF Then
			Repsel = tblRsel("RID") 
		End If	
		tblRsel.Close
		Set tblRsel = Nothing
		''''''LIST''''''''''''''''''''''''''''''''''''''''''''''
			'''''''WORKER'''''
		Set tblWork = Server.CreateObject("ADODB.RecordSet")
		Set tblNme = Server.CreateObject("ADODB.RecordSet")
		'response.write "ID: " & Session("Cidx")
		sqlWork = "SELECT * FROM [ConWork_t] WHERE [CID] = '" & CID & "' AND [WID] IS NOT NULL" 
		Response.Write vbCrlf & "<!-- SQL: " & sqlWork & " -->" & vbCrlf
		tblWork.Open sqlWork, g_strCONN, 1, 3
	'On Error Resume Next
		If Not tblWork.EOF Then
			ctrW = 0
			'response.write "WID: " & tblWork("WID")
			Do Until tblWork.EOF
				if Z_IsOdd(ctrW) = true then 
					kulay = "#FFFAF0" 
				else 
					kulay = "#FFFFFF"
				end if
				tmpWID = tblWork("WID") 
				sqlName = "SELECT [lname], [fname], [Social_Security_Number], [status], [Index] as myindex, address, state, city, zip FROM [Worker_t] WHERE Worker_t.[Index] = " & tmpWID
				'response.write "<br>" & "sqlName: " & sqlName
				Response.Write vbCrLf & "<!-- [" & sqlNAme & "] -->" & vbCrLf
				tblNme.Open sqlName, g_strCONN, 1, 3
				If Not tblNme.EOF Then
					tmpName = tblNme("Lname") & ", " & tblNme("Fname")
					strWork = strWork & "<tr bgcolor='" & kulay & "'><td align='center'><input type='checkbox' name='chkW" & ctrW & "' value='" & _
							tblWork("ID") & "'><td><a href='A_Worker.asp?WID=" & _
							tblNme("Social_Security_Number") & "'><font size='1' face='trebuchet MS'>&nbsp;" & Right(tblNme("Social_Security_Number"), 4) & _
							"&nbsp;</font></a></td><td>"
					If tblNme("Status")  = "InActive" Then		
						strWork = strWork & "<font size='2'>&nbsp;*" & tmpName & "&nbsp;</font></td>"
					Else
						strWork = strWork & "<font size='2'>&nbsp;" & tmpName & "&nbsp;</font></td>"
					End If
					woradr = Replace(tblNme("address"), "#", "") & ", " & tblNme("city") & ", " & tblNme("state") & ", " & tblNme("zip")  
					strWork = strWork & "<td align='center'><a target='_blank' href='http://maps.google.com/maps?saddr=" & woradr & "&daddr=" & gmapsaddr & "'><img src='images/gmaps.png' title='Show directions/distance from Consumer'></a></td></tr>"
					tblNme.Close
				End If
				tblWork.MoveNext
				ctrW = ctrW + 1
			Loop
		Else
			strWork = "<tr><td align='center'><font size='2'>N/A</font></td></tr>"
		End If
		tblWork.Close
		Set tblNme = Nothing
		Set tblWork = Nothing	
		'''''''BACKUP WORKER'''''
		Set tblWork = Server.CreateObject("ADODB.RecordSet")
		Set tblNme = Server.CreateObject("ADODB.RecordSet")
		'response.write "ID: " & Session("Cidx")
		sqlWork = "SELECT [WID], [ID] FROM [ConWorkBack_t] WHERE [CID] = '" & CID & "' AND [WID] IS NOT NULL" 
		Response.Write vbCrLf & "<!-- SQL: " & sqlWork & " -->" & vbCrLf
		tblWork.Open sqlWork, g_strCONN, 1, 3
	'On Error Resume Next
		If Not tblWork.EOF Then
			ctrW = 0
			'response.write "WID: " & tblWork("WID")
			Do Until tblWork.EOF
				if Z_IsOdd(ctrW) = true then 
					kulay = "#FFFAF0" 
				else 
					kulay = "#FFFFFF"
				end if
				tmpWID = (tblWork("WID") )
				If IsNumeric( tmpWID ) AND tmpWID > 0 Then 
					sqlName = "SELECT Lname, Fname, Status, Social_Security_Number, [index]  FROM [Worker_t] WHERE [Index] = " & tmpWID
					Response.Write vbCrLf & "<!-- SQL2: " & sqlName & " -->" & vbCrLf
					tblNme.Open sqlName, g_strCONN, 1, 3
					tmpName = tblNme("Lname") & ", " & tblNme("Fname")
					strWorkback = strWorkback & "<tr bgcolor='" & kulay & "'><td align='center'><input type='checkbox' name='chkWBack" & ctrW & "' value='" & _
							tblWork("ID") & "'><td><a href='A_Worker.asp?WID=" & _
							tblNme("Social_Security_Number") & "'><font size='1' face='trebuchet MS'>&nbsp;" & Right(tblNme("Social_Security_Number"), 4) & _
							"&nbsp;</font></a></td><td>"
					If tblNme("Status")  = "InActive" Then		
						strWorkback =strWorkbackstrWork & "<tr><td align='center'><font size='2'>&nbsp;*" & tmpName & "&nbsp;</font></td></tr>"
					Else
						strWorkback = strWorkback & "<tr><td align='center'><font size='2'>&nbsp;" & tmpName & "&nbsp;</font></td></tr>"
					End If
					tblNme.Close
					tblWork.MoveNext
					ctrWBack = ctrWBack + 1
				End If
			Loop
		Else
			strWorkback = "<tr><td align='center'><font size='2'>N/A</font></td></tr>"
		End If
		tblWork.Close
		Set tblNme = Nothing
		Set tblWork = Nothing	
		'''''''''''CASE MANAGER'''''
		
		Set tblCM = Server.CreateObject("ADODB.RecordSet")
		Set tblNme = Server.CreateObject("ADODB.RecordSet")
		
		sqlCM = "SELECT * FROM [CMCon_t] WHERE [CID] = '" & CID & "' " 
		
		tblCM.Open sqlCM, g_strCONN, 1, 3
		If Not tblCM.EOF Then
			ctr = 0
			Do Until tblCM.EOF
				sqlName = "SELECT Lname, fname, [index] as myindex FROM [Case_Manager_t] WHERE [Index] = " & tblCM("CMID") 
				tblNme.Open sqlName, g_strCONN, 1, 3
				tmpName = tblNme("Lname") & ", " & tblNme("Fname")
				strCM = strCM & "<tr><td align='center'><input type='checkbox' name='chkCM" & ctr & "' value='" & _
						tblNme("myindex") & "'></td><td><a href='A_Case.asp?CaID=" & tblNme("myindex") & "'><font size='1' face='trebuchet MS'>&nbsp;" & _
						tblNme("myindex") & "&nbsp;</font></a></td><td>" & _
						"<font size='2'>&nbsp;" & tmpName & "&nbsp;</font></td></tr>"
				tblNme.Close
				tblCM.MoveNext
				ctr = ctr + 1
			Loop
		Else
			strCM = "<tr><td align='center'><font size='2'>N/A</font></td></tr>"
		End If
		tblCM.Close
		Set tblNme = Nothing
		Set tblCM = Nothing	
		'''''''''''''REPRESENTATIVE'''''''
		
		Set tblRep = Server.CreateObject("ADODB.RecordSet")
		Set tblNme = Server.CreateObject("ADODB.RecordSet")
		
		sqlRep = "SELECT * FROM [ConRep_t] WHERE [CID] = '" & CID & "' " 
		
		tblRep.Open sqlRep, g_strCONN, 1, 3
		If Not tblRep.EOF Then
			ctr = 0
			Do Until tblRep.EOF
				sqlName = "SELECT Lname, fname, [index] FROM [Representative_t] WHERE [Index] = " & tblRep("RID") 
				tblNme.Open sqlName, g_strCONN, 1, 3
				tmpName = tblNme("Lname") & ", " & tblNme("Fname")
				strRep = strRep & "<tr><td align='center'><input type='checkbox' name='chkR" & ctr & "' value='" & _
						tblNme("Index") & "'><td><a href='A_Rep.asp?RID=" & tblNme("Index") & "'><font size='1' face='trebuchet MS'>&nbsp;" & _
						tblNme("Index") & "&nbsp;</font></a></td><td>" & _
						"<font size='2'>&nbsp;" & tmpName & "&nbsp;</font></td></tr>"
				tblNme.Close
				tblRep.MoveNext
				ctr = ctr + 1
			Loop
		Else
			strRep = "<tr><td align='center'><font size='2'>N/A</font></td></tr>"
		End If
		tblRep.Close
		Set tblNme = Nothing
		Set tblRep = Nothing
		''''''CHECK WORKER LIST
		Set fso = CreateObject("Scripting.FileSystemObject")
		If Not fso.FileExists(WorkerList) Then
			Set NewWorkList = fso.CreateTextFile(WorkerList)
			Set tblLWork = Server.CreateObject("ADODB.Recordset")
			strSQLd = "SELECT * FROM [Worker_t] WHERE Status = 'Active' OR Status = 'InActive' ORDER BY [lname], [Fname]"
			tblLWork.Open strSQLd, g_strCONN, 3, 1
			tblLWork.Movefirst
			Do Until tblLWork.EOF
				NewWorkList.WriteLine tblLWork("index") & "|<option value='" & tblLWork("Social_Security_Number")& "'> "& tblLWork("Lname") & ", " & tblLWork("fname") & " (" & tblLWork("city") & ", " & tblLWork("state") & "  " & tblLWork("zip") & ")</option>"
				tblLWork.Movenext
			Loop
			tblLWork.Close
			Set tblLWork = Nothing
			Set NewWorkList = Nothing
		End If
		
		Set tblChkCon = Server.CreateObject("ADODB.Recordset") 'GET LINKS
		sqlChkCon = "SELECT * FROM ConWork_t WHERE CID = '" & CID & "' "
		tblChkCon.Open sqlChkCon, g_strCONN, 1, 3
		ConLink = ""
		Do Until tblChkCon.EOF
			ConLink = ConLink & tblChkCon("WID") & ","
			tblChkCon.MoveNext
		Loop
		tblChkCon.Close
		Set tblChkCon = Nothing
		ConLinkList = Split(ConLink, ",")
		'response.write ConLink
		Set WorkList = fso.OpenTextFile(WorkerList, 1) 'CREATE DROPDOWN
		Do Until WorkList.AtEndofStream
			WorkerLine = WorkList.ReadLine
			WorkerIndex = GetWID(WorkerLine)
			WorkerLine = FixLine(WorkerLine)
			ctrWork = 0
			Meron = 0
			Do Until ctrWork = Ubound(ConLinkList) + 1
				If Z_CZero(ConLinkList(ctrWork)) = Z_CZero(WorkerIndex) Then 
					Meron = 1	
				End If
				ctrWork = ctrWork + 1
			Loop
			If Meron = 0 Then strdept = strdept & WorkerLine
		Loop
		Set WorkList = Nothing
		Set fso = Nothing
		''''''
		''''''''''''''''CASE MANAGER DROPDOWN
		Set tblCMCon = Server.CreateObject("ADODB.Recordset")
		Set tblCM2 = Server.CreateObject("ADODB.Recordset")
		
		sqlCMCon = "SELECT * FROM [CMCon_t] WHERE CID = '" & CID & "' "
		sqlCM2 = "SELECT * FROM [Case_Manager_t] ORDER BY [Lname]"
		
		tblCMCon.Open sqlCMCon, g_strCONN, 1, 3
		tblCM2.Open sqlCM2, g_strCONN, 1, 3
		
		If tblCMCon.EOF Then
			CMLocked = ""
			Do Until tblCM2.EOF
				strCM2 = strCM2 & "<option value='" & tblCM2("index")& "'> "& tblCM2("Lname") & ", " & tblCM2("fname") & " </option>"
				tblCM2.MoveNext
			loop
		Else
			CMLocked = "DISABLED"
			
		End If
	tblCM2.close	
	tblCMCon.Close
	set tblCMCon = Nothing
	Set tblCM2 = Nothing
	
		''''''''''''''''REPRESENTATIVE DROPDOWN
		Set tblRCon = Server.CreateObject("ADODB.Recordset")
		Set tblR2 = Server.CreateObject("ADODB.Recordset")
		
		sqlRCon = "SELECT * FROM [ConRep_t] WHERE CID = '" & CID & "' "
		sqlR2 = "SELECT * FROM [Representative_t] ORDER BY [Lname]"
		
		tblRCon.Open sqlRCon, g_strCONN, 1, 3
		tblR2.Open sqlR2, g_strCONN, 1, 3
		
		If tblRCon.EOF Then
			RLocked = ""
			Do Until tblR2.EOF
				strR2 = strR2 & vbTab & vbTab & vbTab & "<option value=""" & tblR2("Index")& """>" & _
						tblR2("Lname") & ", " & tblR2("Fname") & " </option>" & vbCrLf				
				tblR2.MoveNext
			loop
		Else
			RLocked = "DISABLED"
		End If
		tblR2.close	
		tblRCon.Close
		set tblRCon = Nothing
		Set tblR2 = Nothing
		'''''''''''''''PM
		Set rsPM = Server.CreateObject("ADODB.RecordSet")
		sqlPM = "SELECT * FROM Proj_Man_T ORDER BY Lname, Fname"
		rsPM.Open sqlPM, g_strCONN, 1, 3
		Do Until rsPM.EOF
			PMname = rsPM("Lname") & ", " & rsPM("Fname")
			SelPM = ""
			If rsPM("ID") = PM Then SelPM = "SELECTED"
			strPM = StrPM & "<option " & SelPM & " value='" & rsPM("ID") & "' >" & PMname & "</option>" & vbCrLf
			rsPM.MoveNext
		Loop
		rsPM.Close
		Set rsPM = Nothing
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
		'''''''''''''''MCC
		Set rsPM = Server.CreateObject("ADODB.RecordSet")
		sqlPM = "SELECT * FROM ManagedCare_T ORDER BY MCCname"
		rsPM.Open sqlPM, g_strCONN, 1, 3
		Do Until rsPM.EOF
			MCCname = rsPM("MCCname")
			Selmmc = ""
			If rsPM("MCCid") = cmc Then Selmmc = "SELECTED"
			strmcc = strmcc & "<option " & Selmmc & " value='" & rsPM("MCCid") & "' >" & MCCname & "</option>" 
			rsPM.MoveNext
		Loop
		rsPM.Close
		Set rsPM = Nothing
		''''''''''''''''BACKUP WORKER DROPDOWN
		Set tblLWork = Server.CreateObject("ADODB.Recordset")
	strSQLd = "SELECT * FROM [Worker_t] WHERE Status = 'Active' OR Status = 'InActive' ORDER BY [lname], [Fname]"
	'On Error Resume Next
	tblLWork.Open strSQLd, g_strCONN, 3, 1
		tblLWork.Movefirst
		Do Until tblLWork.EOF
			Set tblChkCon = Server.CreateObject("ADODB.Recordset")
			sqlChkCon = "SELECT * FROM ConWorkBack_t WHERE CID = '" & Session("CID") & "' "
			tblChkCon.Open sqlChkCon, g_strCONN, 1, 3
			If Not tblChkCon.EOF Then
				meron = 0
				Do Until tblChkCon.EOF
					If tblLWork("index") = Cint(tblChkCon("WID")) Then 
						'response.write "SQL" & sqlChkCon & " TRUE" 
						meron = 1
					End If		
					'response.write "SQL: " & sqlChkCon & " FALSE: " & tblLWork("index") & " <> " & tblChkCon("WID") & "<BR>"
					tblChkCon.MoveNext
				Loop
				If meron <> 1 Then	strdeptback = strdeptback & vbTab & vbTab & vbTab & vbTab & _
						"<option value=""" & tblLWork("index")& """>" & tblLWork("Lname") & ", " & _
						tblLWork("fname") & " </option>" & vbCrLf
			Else
				strdeptback = strdeptback & vbTab & vbTab & vbTab & vbTab & _
						"<option value='" & tblLWork("index")& "'>" & tblLWork("Lname") & ", " & _
						tblLWork("fname") & "</option>" & vbCrLf
			End If
			tblLWork.Movenext
		loop
	
	tblLWork.Close
	set tblLWork = Nothing
	''''county
		Set rsPM = Server.CreateObject("ADODB.RecordSet")
		sqlPM = "SELECT * FROM county_T ORDER BY county"
		rsPM.Open sqlPM, g_strCONN, 1, 3
		Do Until rsPM.EOF
			selcount = ""
			if concounty = rsPM("uid") then selcount = "selected" 
			strcount = strcount & "<option " &  selCount & " value='" & rsPM("uid") & "' >" & rsPM("county") & "</option>" & vbCrLf
			rsPM.MoveNext
		Loop
		rsPM.Close
		Set rsPM = Nothing
	Else
		Session("MSG") = "Sorry. To maintain integrity of database, please Sign-in again."
		Response.Redirect "default.asp"
	End If
%>
<html>
	<head>
		<title>LSS - In-Home Care - Consumer Details</title>
		<script language='JavaScript'>
			function CD_Edit()
			{
				//check driver
				
				document.frmConDet.action = "A_C_Action.asp?page=1";
				document.frmConDet.submit();
			}
			function CD_Del()
			{
				var ans = window.confirm("Click OK to continue deletion of ALL Consumer details. Click Cancel to stop.");
				if (ans){
				document.frmConDet.action = "A_Del.asp?act=4";
				document.frmConDet.submit();
				}
			}
			function DelList()
			{
				document.frmConDet.action = "ConList.asp?Del=1";
				document.frmConDet.submit();
			}
			function SavList()
			{
				document.frmConDet.action = "ConList.asp";
				document.frmConDet.submit();
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
			function myRate()	{
				if (document.frmConDet.selcode.value == "M") {
					document.frmConDet.txtrate.disabled = true;
					document.frmConDet.chkVAHM.disabled = true;
					document.frmConDet.hrshm.disabled = true;
					document.frmConDet.chkVAHA.disabled = true;
					document.frmConDet.hrsha.disabled = true;
					document.frmConDet.chkVAHM.checked = false;	
					document.frmConDet.chkVAHA.checked = false;
				}
				else if (document.frmConDet.selcode.value == "V") {
					document.frmConDet.chkVAHM.disabled = false;
					document.frmConDet.hrshm.disabled = false;
					document.frmConDet.chkVAHA.disabled = false;
					document.frmConDet.hrsha.disabled = false;
					document.frmConDet.txtrate.disabled = true;
				}
				else
					{document.frmConDet.txtrate.disabled = false;
						document.frmConDet.chkVAHM.disabled = true;
				document.frmConDet.hrshm.disabled = true;
				document.frmConDet.chkVAHA.disabled = true;
				document.frmConDet.hrsha.disabled = true;
				document.frmConDet.chkVAHM.checked = false;
				//document.frmConDet.hrshm.value = "";
				document.frmConDet.chkVAHA.checked = false;
				//document.frmConDet.hrsha.value = "";
						}
			}
			function hrschk(xxx) {
				if (xxx == 1) {
					if (document.frmConDet.chkVAHM.checked == false) {
						document.frmConDet.hrshm.value = "";
					}
				}
				else if(xxx == 2){
					if (document.frmConDet.chkVAHA.checked == false) {
						document.frmConDet.hrsha.value = "";
					}
				}
			}
			function workermatch(xxx) {
				newwindow = window.open("workermatch.asp?cid=" + xxx,'Worker Match','height=600,width=550,scrollbars=1,directories=0,status=0,toolbar=0,resizable=0');
				if (window.focus) {newwindow.focus()}
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
	<body bgcolor='white' LEFTMARGIN='0' TOPMARGIN='0' onload='myRate();'>
		<form method='post' name='frmConDet'>
			<!-- #include file="_boxup.asp" -->
			<!-- #include file="_NavHeader.asp" -->
			<br>
			<center>
				<table border='0'  width='575px'>
					<tr>
						<td colspan='4' align='center' width='550px'>
							<font size='2' face='trebuchet MS'><b><u>Consumer Details</u></b></font>
							<font size='2' face='trebuchet MS'>[Details]</font>
							<a href='A_C_Status.asp?MNum=<%=CID%>' style='text-decoration:none'><font size='1' face='trebuchet MS'>[Status]</font></a>
							<a href='A_C_health.asp?MNum=<%=CID%>' style='text-decoration:none'><font size='1' face='trebuchet MS'>[Health]</font></a>
							<a href='A_C_Files.asp?MNum=<%=CID%>' style='text-decoration:none'><font size='1' face='trebuchet MS'>[Files]</font></a>
							<a href='Log.asp?MNum=<%=CID%>' style='text-decoration:none'><font size='1' face='trebuchet MS'>[Log]</font></a>
							<a href='cimport.asp?MNum=<%=CID%>' style='text-decoration:none'><font size='1' face='trebuchet MS'>[Uploads]</font></a>
						</td>
					</tr>
					<tr><td colspan='2' align='center'><font color='red'  face='trebuchet MS' size='1'>&nbsp;<%=Session("MSG")%>&nbsp;</font></td></tr>
					<% If Session("lngType") = 1 Or Session("lngType") = 2 Or session("UserID") = 893 Then %>
					<tr>
						<td colspan='2'>
							<font size='1' face='trebuchet MS'><u>Medicaid Number:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 90px;' readonly name='MCnum' <%=FinOnly%> value="<%=MCNum%>"></font>
						&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
						&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
						&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							<font size='1'face='trebuchet MS' ><u>SSN:</u>&nbsp;
							<input style='font-size: 10px; height: 20px; width: 80px;' maxlength='11' type='text' name='SSN' <%=FinOnly%> value="<%=SSN%>"></td>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
				<% Else %>
					<tr>
						<td colspan='2'>
							<font size='1' face='trebuchet MS'><u>Medicaid Number:</u>&nbsp;<%=MCNum%></font>
					<input type='hidden' name='MCnum' value="<%=MCNum%>">
					<input type='hidden' name='SSN' value="<%=SSN%>">
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
				<% End If %>
					<tr>
						<td>
							<font size='1' face='trebuchet MS'><u>Name:</u>&nbsp;<input <%=FinOnly%> style='font-size: 10px; height: 20px; width: 80px;' type='text' maxlength='20' name='lname' value="<%=Lname%>">,&nbsp; 
							<input <%=FinOnly%> style='font-size: 10px; height: 20px; width: 80px;' type='text' name='fname' maxlength='20' value="<%=Fname%>">
							<font size='1' face='trebuchet MS'>(Last Name, First Name)</font>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td>
							<font size='1' face='trebuchet MS'><u>Address:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 250px;' name='Addr' maxlength='49' value="<%=Addr%>"></font>
						</td>
						
					</tr>
					<tr>
						<td colspan='2' >
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							<font size='1' face='trebuchet MS'><u>City:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 100px;' name='cty' maxlength='50' value="<%=cty%>">
							<font size='1' face='trebuchet MS'><u>State:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 50px;' name='ste' maxlength='2' value="<%=ste%>">
							<font size='1' face='trebuchet MS'><u>Zip Code:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 100px;' name='zcode' maxlength='10' value="<%=zcode%>" onKeyUp="javascript:return maskMe(this.value,this,'5','-');" onBlur="javascript:return maskMe(this.value,this,'5','-');">
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td>
							<font size='1' face='trebuchet MS'><u>Mailing Address:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 250px;' name='MAddr' maxlength='49' value="<%=MAddr%>"></font>
							<input type='checkbox' name='chkMail'><font size='1' face='trebuchet MS'><u>Same as Residence:</u></font>
						</td>
					</tr>
					<tr>
						<td colspan='2'>
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							<font size='1' face='trebuchet MS'><u>City:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 100px;' name='Mcty' maxlength='49' value="<%=Mcty%>">
							<font size='1' face='trebuchet MS'><u>State:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 50px;' name='Mste' maxlength='2' value="<%=Mste%>">
							<font size='1' face='trebuchet MS'><u>Zip Code:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 100px;' name='Mzcode' maxlength='10' value="<%=Mzcode%>" onKeyUp="javascript:return maskMe(this.value,this,'5','-');" onBlur="javascript:return maskMe(this.value,this,'5','-');">
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td>
							<font size='1' face='trebuchet MS'><u>Phone No:</u>&nbsp;<input style='font-size: 10px; height: 20px; width: 80px;' type='text' maxlength='12' name='PhoneNo' value="<%=FonNum%>" onKeyUp="javascript:return maskMe(this.value,this,'3,7','-');" onBlur="javascript:return maskMe(this.value,this,'3,7','-');"></font>
							&nbsp;&nbsp;&nbsp;&nbsp;
							<font size='1' face='trebuchet MS'><u>eMail:</u>&nbsp;<input style='font-size: 10px; height: 20px; width: 125px;' type='text' maxlength='50' name='email' value="<%=email%>"></font>
						</td>
					</tr>
						<tr><td>&nbsp;</td></tr>
					<tr>
						<td>
							<font size='1' face='trebuchet MS'><u>County:</u>&nbsp;
								<select style='font-size: 10px; height: 20px; width: 100px;' name='selcount'></font>
								<option value="0">&nbsp;</option>
								<%=strcount%>
								</select>
								&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							<font size='1' face='trebuchet MS'><u>Language:</u>&nbsp;</font>
							<select style='font-size: 10px; height: 20px; width: 200px;' name='langid' id='langid'>
								<%=langsel%>
							</select>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
						<tr>
						<td>
							<font size='1' face='trebuchet MS'><u>Secondary Phone No:</u>&nbsp;<input style='font-size: 10px; height: 20px; width: 80px;' type='text' maxlength='12' name='PhoneNo2' value="<%=FonNum2%>" onKeyUp="javascript:return maskMe(this.value,this,'3,7','-');" onBlur="javascript:return maskMe(this.value,this,'3,7','-');"></font>
							&nbsp;&nbsp;&nbsp;&nbsp;
							<font size='1' face='trebuchet MS'><u>Mobile Number:</u>&nbsp;<input style='font-size: 10px; height: 20px; width: 80px;' type='text' maxlength='12' name='mobilenum' value="<%=celNum%>" onKeyUp="javascript:return maskMe(this.value,this,'3,7','-');" onBlur="javascript:return maskMe(this.value,this,'3,7','-');"></font>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
							<td valign='center'><font size='1' face='trebuchet MS'><u>Emergency Contact Info:</u></font>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 250px;' name='emerinfo' maxlength='49' value="<%=emerinfo%>"</td>
						</td>
					</tr>
					<tr>
						<td>
								&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							<font size='1' face='trebuchet MS'><u>Relationship:</u>&nbsp;<input style='font-size: 10px; height: 20px; width: 80px;' type='text' maxlength='49' name='emerrel' value="<%=emerrel%>" ></font>
							&nbsp;&nbsp;&nbsp;&nbsp;
							<font size='1' face='trebuchet MS'><u>Contact Number:</u>&nbsp;<input style='font-size: 10px; height: 20px; width: 80px;' type='text' maxlength='12' name='emerphone' value="<%=emerphone%>" onKeyUp="javascript:return maskMe(this.value,this,'3,7','-');" onBlur="javascript:return maskMe(this.value,this,'3,7','-');"></font>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td>
							<font size='1' face='trebuchet MS'><u>DOB:</u>&nbsp;
							<input <%=FinOnly%> type='text' style='font-size: 10px; height: 20px; width: 80px;' maxlength='10' name='DOB' value="<%=DOB%>" onKeyUp="javascript:return maskMe(this.value,this,'2,5','/');" onBlur="javascript:return maskMe(this.value,this,'2,5','/');"></font>
							<font size='1' face='trebuchet MS'>(mm/dd/yyyy)</font>
							<font size='1' face='trebuchet MS'><u>Gender:</u>&nbsp;
							<select style='font-size: 10px; height: 20px; width: 80px;' name='Gen'></font>
								<option <%=M%>>Male</option>
								<option <%=F%>>Female</option>
							</select>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td>
							<font size='1' face='trebuchet MS'><u>Directions:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 400px;' maxlength='100' name='Direct' value="<%=Direct%>"></font>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td>
							<font size='1' face='trebuchet MS'><u>Referral Date:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 80px;' name='RefDate' maxlength='10' value="<%=Ref_Dte%>" onKeyUp="javascript:return maskMe(this.value,this,'2,5','/');" onBlur="javascript:return maskMe(this.value,this,'2,5','/');"></font>
							<font size='1' face='trebuchet MS'>(mm/dd/yyyy)</font>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td>
							<font size='1' face='trebuchet MS'><u>Consumer Start Date:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 80px;' name='StrtDte' maxlength='10' value="<%=Strt_Dte%>" onKeyUp="javascript:return maskMe(this.value,this,'2,5','/');" onBlur="javascript:return maskMe(this.value,this,'2,5','/');"></font>
							<font size='1' face='trebuchet MS'>(mm/dd/yyyy)</font>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td>
							<font size='1' face='trebuchet MS'><u>Amendment Effective Date:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 80px;' name='EffDte' maxlength='10' value="<%=EffDte%>" onKeyUp="javascript:return maskMe(this.value,this,'2,5','/');" onBlur="javascript:return maskMe(this.value,this,'2,5','/');"></font>
							<font size='1' face='trebuchet MS'>(mm/dd/yyyy)</font>
							<font size='1' face='trebuchet MS'><u>Amendment Expiraton Date:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 80px;' name='EndDte' maxlength='10' value="<%=EndDte%>" onKeyUp="javascript:return maskMe(this.value,this,'2,5','/');" onBlur="javascript:return maskMe(this.value,this,'2,5','/');"></font>
							<font size='1' face='trebuchet MS'>(mm/dd/yyyy)</font>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td>
							<font size='1' face='trebuchet MS'><u>Amendment Signed:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 80px;' name='ASDte' maxlength='10' value="<%=AmendS%>" onKeyUp="javascript:return maskMe(this.value,this,'2,5','/');" onBlur="javascript:return maskMe(this.value,this,'2,5','/');"></font>
							<font size='1' face='trebuchet MS'>(mm/dd/yyyy)</font>
							<font size='1' face='trebuchet MS'><u>Amendment Received:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 80px;' name='ARDte' maxlength='10' value="<%=AmendR%>" onKeyUp="javascript:return maskMe(this.value,this,'2,5','/');" onBlur="javascript:return maskMe(this.value,this,'2,5','/');"></font>
							<font size='1' face='trebuchet MS'>(mm/dd/yyyy)</font>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td>
							<font size='1' face='trebuchet MS'><u>Consumer End Date:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 80px;' name='TermDte' maxlength='10' value="<%=termDte%>" onKeyUp="javascript:return maskMe(this.value,this,'2,5','/');" onBlur="javascript:return maskMe(this.value,this,'2,5','/');"></font>
							<font size='1' face='trebuchet MS'>(mm/dd/yyyy)</font>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td>
							<font size='1' face='trebuchet MS'><u>Current Care Plan Date:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 80px;' name='CareDte' maxlength='10' value="<%=CareDte%>" onKeyUp="javascript:return maskMe(this.value,this,'2,5','/');" onBlur="javascript:return maskMe(this.value,this,'2,5','/');"></font>
							<font size='1' face='trebuchet MS'>(mm/dd/yyyy)</font>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td>
							<font size='1' face='trebuchet MS'><u>Code:</u>&nbsp;
								<select style='font-size: 10px; height: 20px; width: 40px;' name='selcode' onchange='JavaScript:myRate();'>
								<option value='M' <%=mcode%> >M</option>
								<option value='P' <%=pcode%> >P</option>
								<option value='C' <%=ccode%> >C</option>
								<option value='A' <%=acode%> >A</option>
								<option value='V' <%=vcode%> >V</option>
							</select>
							<font size='1' face='trebuchet MS'><u>Rate:</u>&nbsp;
								<input type='text' style='font-size: 10px; height: 20px; width: 80px;' name='txtrate' maxlength='10' value="<%=tmprate%>" >
							</font>
						</td>
					</tr>
					<tr>
						<td align='left' colspan='2'>
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							<input type='checkbox' name='chkVAHM' value='1' <%=vahm%> onclick='hrschk(1);'>
							<font size='1' face='trebuchet MS'><u>VA-HM</u>&nbsp;
						    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
						    <font size='1' face='trebuchet MS'><u>Hours:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 80px;' name='hrshm' maxlength='6' value="<%=hrshm%>"></font>
						</td>
					</tr>
					<tr>
						<td align='left' colspan='2'>
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							<input type='checkbox' name='chkVAHA' value='1' <%=vaha%> onclick='hrschk(2);'>
							<font size='1' face='trebuchet MS'><u>VA-HA</u>&nbsp;
						    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
						    <font size='1' face='trebuchet MS'><u>Hours:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 80px;' name='hrsha' maxlength='6' value="<%=hrsha%>"></font>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td>
							<font size='1' face='trebuchet MS'><u>Private Pay Contract Hours:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 80px;' name='txtcon' maxlength='10' value="<%=contract%>"></font>
						</td>
						
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td align='left' colspan='2'>
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							<input type='checkbox' name='chkDrive' <%=Drive%>>
							<font size='1' face='trebuchet MS'><u>Driving Part of Care Plan:</u>&nbsp;
						    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
						    <font size='1' face='trebuchet MS'><u>Max Hours:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 80px;' name='maxhrs' maxlength='6' value="<%=maxhrs%>"></font>
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
						    <font size='1' face='trebuchet MS'><u>Mileage Cap:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 80px;' name='txtmile' maxlength='6' value="<%=milecap%>"></font>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td>
							<font size='1' face='trebuchet MS'><u>Managed Care Company:</u></font>
							<select style='font-size: 10px; height: 20px; width: 150px;' name='selmmc'>
								<option value='0'></option>
								<%=strmcc%>
							</select>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
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
					<tr>
						<td valign='center'>
							<font size='1' face='trebuchet MS'><u>RIHCC:</u></font>
							<select style='font-size: 10px; height: 20px; width: 150px;'name='PMsel'>
								<option value='0'></option>
								<%=strPM%>
							</select>
							&nbsp;&nbsp;&nbsp;&nbsp;
							<font size='1' face='trebuchet MS'><u>Comments:</u>&nbsp;
							<textarea rows='2' name='Pcomments' cols="20" ><%=cmt%></textarea>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td>
							<font size='1' face='trebuchet MS'><u>Client ID:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 80px;' name='txtcid' maxlength='6' readonly value="<%=cliid%>"></font>
						</td>
					</tr>
					<tr>
						<td>
							<font size='1' face='trebuchet MS'><u>Person Code:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 80px;' name='txtpcode' maxlength="8" size="9" <%=FinOnly%> value="<%=percode%>"></font>
						</td>
					</tr>
					<tr>
						<td align='left' colspan='2'>
							<font size='1' face='trebuchet MS'><u>PCSP is a relative :</u>&nbsp;
							<input type='checkbox' name='chkwrel' <%=rela%> value=1>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td>
							<font size='1' face='trebuchet MS'><u>Authorized CID number:</u>
						</td>
					</tr>
					<tr>
						<td>
							&nbsp;&nbsp;<font size='1' face='trebuchet MS'><u>1:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 80px;' name='txtafon1' maxlength='11' value="<%=aphone1%>" ></font>
							&nbsp;&nbsp;<font size='1' face='trebuchet MS'><u>4:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 80px;' name='txtafon4' maxlength='11' value="<%=aphone4%>" ></font>
						</td>
					</tr>
					<tr>
						<td>
							&nbsp;&nbsp;<font size='1' face='trebuchet MS'><u>2:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 80px;' name='txtafon2' maxlength='11' value="<%=aphone2%>" ></font>
							&nbsp;&nbsp;<font size='1' face='trebuchet MS'><u>5:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 80px;' name='txtafon5' maxlength='11' value="<%=aphone5%>" ></font>
						</td>
					</tr>
					<tr>
						<td>
							&nbsp;&nbsp;<font size='1' face='trebuchet MS'><u>3:</u>&nbsp;
							<input type='text' style='font-size: 10px; height: 20px; width: 80px;' name='txtafon3' maxlength='11' value="<%=aphone3%>" ></font>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					</table>
					<table>
					<tr>
						<td align='center' colspan='2'>
							<table border='1'>
								<tr bgcolor='#C4B464'>
									<td colspan='4' style="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=0, startColorstr=#C4B464, endColorstr=#f7efc7);">
										<a href='JavaScript: SavList();' style='text-decoration: none'><font size='2' face='trebuchet ms'>[Save List]</font></a>
										<a href='JavaScript: DelList();' style='text-decoration: none'><font size='2' face='trebuchet ms'>[Delete Checked List]</font></a>
										<font size='1' face='trebuchet MS'>* Inactive</font>
										<input type='hidden' name='tmpID' value='<%=CID%>'>
										<input type='hidden' name='ctr' value='<%=ctrW%>'>
										<input type='hidden' name='ctr2' value='<%=ctrWBack%>'>
								</td>
								</tr>
							
							<tr>
			<td valign='top'>
			<table border='1'>
				<tr bgcolor='#040c8b'>
					<td style="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=0, startColorstr=#040C8B, endColorstr=#c7c9e5);" colspan='4' align='center' width='170px'>
						<font face='trebuchet MS' size='2' color='white'><b>PCSP Worker List</b></font>&nbsp;<a href='#' onclick='workermatch(<%=index%>);'><img src='images/zoom.gif' title='Find Worker'></a>
					</td></tr>
				<%=strWork%>
				
			</table>
			</td>
			<td valign='top'>
			<table border='1'>
				<tr bgcolor='#040c8b'>
					<td style="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=0, startColorstr=#040C8B, endColorstr=#c7c9e5);" colspan='3' align='center' width='170px'>
						<font face='trebuchet MS' size='2' color='white'><b>Case Manager List</b></font>
						</td></tr>
				<%=strCM%>
				
			</table>
			</td>
			<td valign='top'>
				<table border='1'>
				<tr bgcolor='#040c8b'>
					<td style="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=0, startColorstr=#040C8B, endColorstr=#c7c9e5);" colspan='3' align='center' width='170px'>
						<font face='trebuchet MS' size='2' color='white'><b>Representative List</b></font>
						</td></tr>
				<%=strRep%>
				</table>
			</td>
			<td valign='top'>
			<table border='1'>
				<tr bgcolor='#040c8b'>
					<td style="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=0, startColorstr=#040C8B, endColorstr=#c7c9e5);" colspan='3' align='center' width='170px'>
						<font face='trebuchet MS' size='2' color='white'><b>PCSP Backup Worker List</b></font>
					</td></tr>
				<%=strWorkback%>
				
			</table>
			</td>
		</tr>
		<tr>
			<td><select style='font-size: 10px; height: 20px; width: 200px;' name='SelWor'>
				<option></option>
				<%=strdept%></select></td>
			<td><select style='font-size: 10px; height: 20px; width: 179px;' <%=CMLocked%> name='SelCM'><option></option><%=strCM2%></select></td>
			<td><select style='font-size: 10px; height: 20px; width: 179px;' <%=RLocked%> name='SelR'><option></option><%=strR2%></select></td>
			<td><select style='font-size: 10px; height: 20px; width: 179px;' name='SelWorBack'>
				<option></option>
				<%=strdeptback%></select></td>
		</tr>
		
		
	</td></tr>
</table></td></tr>
<tr><td>&nbsp;</td></tr>
					<tr>
						<td align='center' >
							<input type='button' value='Save' style='width: 110px;' class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='JavaScript:CD_Edit();'>
							<% If UCase(Session("lngType")) = "2" Then %>
								<input type='button' value='Delete' style='width: 110px;' class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='JavaScript:CD_Del();'>
							<% End If %>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
				</table>
			</td></tr>
				<!-- #include file="_boxdown.asp" -->
			</center>	
		</form>
	</body>
</html>
<%
Session("MSG") = "" 
Session("CFiles") = ""
%>