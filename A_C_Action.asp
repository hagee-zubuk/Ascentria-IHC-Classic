<%Language=VBScript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<%
Function GetNEWCID()
	Set rsCID = Server.CreateObject("ADODB.RecordSet")
	sqlCID = "SELECT MAX(CliID) as mycliid FROM Consumer_T"
	rsCID.Open sqlCID, g_strCONN, 3, 1
	GetNEWCID = rsCID("mycliid") + 1
	rsCID.Close
	Set rsCID = Nothing
End Function
Function GetWIndex(xxx)
	Set rsWID = Server.CreateObject("ADODB.RecordSet")
	sqlWID = "SELECT [index] FROM Worker_T WHERE Social_Security_Number = '" & xxx & "'"
	rsWID.Open sqlWID, g_strCONN, 3, 1
	If Not rsWID.EOF Then
			GetWIndex = rsWID("Index")
	End If
	rsWID.CLose
	Set rsWID = Nothing
End Function
	If Request("page") = 1 Then
		Set tblConsumer = Server.CreateObject("ADODB.Recordset")
		If Request("new") <> 1 Then
			sqlConsumer = "SELECT * FROM Consumer_t WHERE Medicaid_Number = '" & Request("MCNum") & "' "
		Else
			Session("CFiles") = Z_DoEncrypt(Request("MCnum") & "|" & Request("SSN") & "|" & Request("lname") & "|" & Request("fname") & "|" & _
					Request("Addr") & "|" & Request("PhoneNo") & "|" & Request("DOB") & "|" & Request("Gen") & "|" & _
					Request("Direct") & "|" & Request("RefDate") & "|" & Request("StrtDte") & "|" &  Request("chkDrive") & "|" & _
					Request("chkrep") & "|" & Request("maxhrs") & "|" & Request("comments") & "|" & Request("cty") & _
					"|" & Request("ste") & "|" & Request("zip") & "|" & Request("ASDte") & "|" & Request("ARDte") & _
					"|" & Request("PMsel") & "|" & Request("TermDate") & "|" & Request("EndDte") & "|" & Request("CareDte") & "|" & _
					Request("Effdte") & "|" & Request("txtmile")& "|" & Request("email")& "|" & Request("mAddr") & _
					"|" & Request("mcty") & "|" & Request("mste") & "|" & Request("mzcode") & "|" & Request("chkMail")& "|" & Request("selcode") & "|" & _
					Request("txtcon")  & "|" & Request("txtrate"))
			If Request("MCnum") = "" Then
				Session("MSG") = "Medicaid Number is blank."
				Response.Redirect "A_New_Consumer.asp"
			Else
				If Request("selcode") = "M" And Len(Request("MCnum")) <> 11 Then
					Session("MSG") = "Medicaid Number should be 11 characters for code M consumers."
					Response.Redirect "A_New_Consumer.asp"
				End If
			End If
			If Request("Strtdte") = "" Then
				Session("MSG") = Session("MSG") & "<br>" & "Consumer Start Date is required."
				Response.Redirect "A_New_Consumer.asp"
			End If
			sqlConsumer = "SELECT * FROM Consumer_t WHERE Medicaid_Number = '" & Trim(Request("MCNum")) & "' "
		End If
		tblConsumer.Open sqlConsumer, g_strCONN, 1, 3
			If Request("new") = 1 Then
				If tblConsumer.EOF Then
					If Request("MCNum") <> "" Then
						'get NEW CID
						tmpCID = GetNEWCID()
						tblConsumer.AddNew
						tblConsumer("CliID") = tmpCID
						tblConsumer("Medicaid_Number") = Trim(Request("MCNum"))
						Session("CID") = Trim(Request("MCNum"))
						Set tblStatus = Server.CreateObject("ADODB.RecordSet")
						sqlStatus = "SELECT * FROM C_Status_t WHERE Medicaid_Number = '" & Request("MCNum") & "' "	
						tblStatus.Open sqlStatus, g_strCONN, 1, 3
						If tblStatus.EOF Then
							tblStatus.AddNew
							tblStatus("Medicaid_Number") = Request("MCNum")
							tblStatus.Update
						End If
						tblStatus.Close
						Set tblStatus = Nothing
					End If
				Else
					Session("MSG") = "Medicaid Number is either blank or already exists."
					Response.Redirect "A_New_Consumer.asp"
					
				End If
			End If
		
		tblConsumer("Social_Security_Number") = Request("SSN")
		tblConsumer("Lname") = Request("Lname")
		tblConsumer("Fname") = Request("Fname")
		If Request("DOB") <> "" Then
			If IsDate(Request("DOB")) Then
				If CDate(Request("DOB")) < Date Then 
					tblConsumer("DOB") = Z_FormatDate(Request("DOB"))
				Else
					
					Session("MSG") = "Invalid DOB." 
					
				End If
			Else
				Session("MSG") = Session("MSG") & "<br>" & "Enter a valid date For DOB."
				
			End If
		Else
			tblConsumer("DOB") = Empty
		End If
		tblConsumer("Gender") = Request("Gen")
		tblConsumer("Address") = Request("Addr")
		tblConsumer("City") = Request("cty")
		tblConsumer("State") = UCase(Request("ste"))
		tblConsumer("zip") = Request("zcode")
		If Request("chkMail") = "" Then
			tblConsumer("mAddress") = Request("mAddr")
			tblConsumer("mCity") = Request("mcty")
			tblConsumer("mState") = UCase(Request("mste"))
			tblConsumer("mzip") = Request("mzcode")
		Else
			tblConsumer("mAddress") = Request("Addr")
			tblConsumer("mCity") = Request("cty")
			tblConsumer("mState") = UCase(Request("ste"))
			tblConsumer("mzip") = Request("zcode")
		End If
		if Request("PhoneNo") <> "" then
			if len(Request("PhoneNo")) = 12 then
				tblConsumer("PhoneNo") = Request("PhoneNo")
			else
				Session("MSG") = Session("MSG") & "<br>" & "Enter a valid format for Phone Number (xxx-xxx-xxxx)."
			end if
		else
			tblConsumer("PhoneNo") = ""
		end if
		if Request("PhoneNo2") <> "" then
			if len(Request("PhoneNo2")) = 12 then
				tblConsumer("secphone") = Request("PhoneNo2")
			else
				Session("MSG") = Session("MSG") & "<br>" & "Enter a valid format for Secondary Phone Number (xxx-xxx-xxxx)."
			end if
		else
			tblConsumer("secphone") = ""
		end if
		if Request("mobilenum") <> "" then
			if len(Request("mobilenum")) = 12 then
				tblConsumer("celphone") = Request("mobilenum")
			else
				Session("MSG") = Session("MSG") & "<br>" & "Enter a valid format for Mobile Number (xxx-xxx-xxxx)."
			end if
		else
			tblConsumer("celphone") = ""
		end if
		tblConsumer("emerinfo") = Request("emerinfo")
		tblConsumer("emerrel") = Request("emerrel")
		tblConsumer("emerphone") = Request("emerphone")
		tblConsumer("mcc") = Request("selmcc")
		tblConsumer("cmc") = Request("selcmc")
		tblconsumer("workrelative") = false
		if request("chkwrel") <> "" Then tblconsumer("workrelative") = True
		tblConsumer("Directions") = Request("Direct")
		tblConsumer("PMID") = Request("PMsel")
		If Request("RefDate") <> "" Then
			If IsDate(Request("RefDate")) Then
				tblConsumer("Referral_Date") = Z_FormatDate(Request("RefDate"))
			Else
				Session("MSG") = Session("MSG") & "<br>" & "Enter a valid date for Referral Date."
			End If
		Else
			tblConsumer("Referral_Date") = Empty
		End If
		If Request("Strtdte") <> "" Then
			If IsDate(Request("Strtdte")) Then
				tblConsumer("Start_Date") = Z_FormatDate(Request("Strtdte"))
			Else
				Session("MSG") = Session("MSG") & "<br>" & "Enter a valid date for Consumer Start Date."
			End If
		Else
			Session("MSG") = Session("MSG") & "<br>" & "Consumer Start Date is required."
		End If
		If Request("Effdte") <> "" Then
			If IsDate(Request("Effdte")) Then
				tblConsumer("EffDate") = Z_FormatDate(Request("Effdte"))
			Else
				Session("MSG") = Session("MSG") & "<br>" & "Enter a valid date for Amendment Effective Date."
			End If
		Else
			tblConsumer("EffDate") = Empty
		End If
		If Request("EndDte") <> "" Then
			If IsDate(Request("EndDte")) Then
				tblConsumer("End_Date") = Z_FormatDate(Request("EndDte"))
			Else
				Session("MSG") = Session("MSG") & "<br>" & "Enter a valid date for Amendment Expiration Date."
			End If
		Else
			tblConsumer("End_Date") = Empty
		End If
		If Request("Termdte") <> "" Then
			If IsDate(Request("Termdte")) Then
				tblConsumer("TermDate") = Z_FormatDate(Request("Termdte"))
			Else
				Session("MSG") = Session("MSG") & "<br>" & "Enter a valid date for Consumer End Date."
			End If
		Else
			tblConsumer("TErmDate") = Empty
		End If
		If Request("ASDte") <> "" Then
			If IsDate(Request("ASDte")) Then
				tblConsumer("AmendSigned") = Z_FormatDate(Request("ASDte"))
			Else
				Session("MSG") = Session("MSG") & "<br>" & "Enter a valid date for Start Date Ammendment Signed."
			End If
		Else
			tblConsumer("AmendSigned") = Empty
		End If
		If Request("ARDte") <> "" Then
			If IsDate(Request("ARDte")) Then
				tblConsumer("DteRecd") = Z_FormatDate(Request("ARDte"))
			Else
				Session("MSG") = Session("MSG") & "<br>" & "Enter a valid date for Start Date Ammendment Received."
			End If
		Else
			tblConsumer("DteRecd") = Empty
		End If
		If Request("chkDrive") <> "" Then 
			tblConsumer("Driving") = True 
		Else
			tblConsumer("Driving") = False
		End If
		If Request("maxhrs") <> "" Then
			If IsNumeric(Request("maxhrs")) Then 
				If Z_Cdbl(Request("maxhrs")) < 169 Then
					tblConsumer("MaxHrs") = Request("maxhrs")
				Else
					Session("MSG") = Session("MSG") & "<br>" & "Max Hours should be lower or equal to 168 hours."	
				End If
			Else
				Session("MSG") = Session("MSG") & "<br>" & "Invalid Max Hours."
			End If
		Else
			tblConsumer("MaxHrs") = 0
		End If
		tblConsumer("PComment") = Request("Pcomments")
		If Request("CareDte") <> "" Then
			If IsDate(Request("CareDte")) Then
				tblConsumer("CarePlan") = Request("CareDte")
			Else
				Session("MSG") = Session("MSG") & "<br>" & "Enter a valid date for Current Care Plan."
			End If
		Else
			tblConsumer("CarePlan") = Empty
		End If
		If Request("txtmile") <> "" Then
			If IsNumeric(Request("txtmile")) Then
				tblConsumer("milecap") = Request("txtmile")
			Else
				Session("MSG") = Session("MSG") & "<br>" & "Invalid Mileage Cap."
			End If
		End If
		tblConsumer("code") = Request("selcode")
		tblconsumer("vahm") = false
		tblconsumer("vaha") = false
		if request("chkVAHM") = 1 Then 
			tblconsumer("vahm") = True 
			if IsNumeric(request("hrshm")) Then
				tblconsumer("vahmhrs") = Z_CZero(request("hrshm"))
			else
				Session("MSG") = Session("MSG") & "<br>" & "Invalid VA-HM hours."
			end if
		end if
		if request("chkVAHa") = 1 Then 
			tblconsumer("vaha") = True 
			if IsNumeric(request("hrsha")) Then
				tblconsumer("vahahrs") = Z_CZero(request("hrsha"))
			else
				Session("MSG") = Session("MSG") & "<br>" & "Invalid VA-HA hours."
			end if
		end if
		If Request("txtcon") <> "" Then
			If IsNumeric(Request("txtcon")) Then
				tblConsumer("contract") = Z_CZero(Request("txtcon"))
			Else
				Session("MSG") = Session("MSG") & "<br>" & "Invalid private pay contract hours."
			End If
		End If
		If Request("txtrate") <> "" Then
			If IsNumeric(Request("txtrate")) Then
				tblConsumer("rate") = Z_CZero(Request("txtrate"))
			Else
				Session("MSG") = Session("MSG") & "<br>" & "Invalid rate."
			End If
		End If
		If Request("selcode") = "M" Then tblConsumer("rate") = 0
		tblConsumer("email") = Request("email")
		if Request("txtafon1") <> "" then
			if len(Request("txtafon1")) = 11 then
				tblConsumer("aphone1") = Request("txtafon1")
			else
				Session("MSG") = Session("MSG") & "<br>" & "Enter a valid format for (1) Authorized CID number (11 digits)."
			end if
		else
			tblConsumer("aphone1") = ""
		end if
		if Request("txtafon2") <> "" then
			if len(Request("txtafon2")) = 11 then
				tblConsumer("aphone2") = Request("txtafon2")
			else
				Session("MSG") = Session("MSG") & "<br>" & "Enter a valid format for (2) Authorized CID number (11 digits)."
			end if
		else
			tblConsumer("aphone2") = ""
		end if
		if Request("txtafon3") <> "" then
			if len(Request("txtafon3")) = 11 then
				tblConsumer("aphone3") = Request("txtafon3")
			else
				Session("MSG") = Session("MSG") & "<br>" & "Enter a valid format for (3) Authorized CID number (11 digits)."
			end if
		else
			tblConsumer("aphone3") = ""
		end if
		if Request("txtafon4") <> "" then
			if len(Request("txtafon4")) = 11 then
				tblConsumer("aphone4") = Request("txtafon4")
			else
				Session("MSG") = Session("MSG") & "<br>" & "Enter a valid format for (4) Authorized CID number (11 digits)."
			end if
		else
			tblConsumer("aphone4") = ""
		end if
		if Request("txtafon5") <> "" then
			if len(Request("txtafon5")) = 11 then
				tblConsumer("aphone5") = Request("txtafon5")
			else
				Session("MSG") = Session("MSG") & "<br>" & "Enter a valid format for (5) Authorized CID number (11 digits)."
			end if
		else
			tblConsumer("aphone5") = ""
		end if
		tblConsumer("pcode") = Request("txtpcode")
		'tblConsumer("aphone1") = Request("txtafon1")
		'tblConsumer("aphone2") = Request("txtafon2")
		'tblConsumer("aphone3") = Request("txtafon3")
		'tblConsumer("aphone4") = Request("txtafon4")
		'tblConsumer("aphone5") = Request("txtafon5")
		tblConsumer("countid") = Request("selcount")
		' new -- 2017-10-27
		tblConsumer("langid") = Z_CLng(Request("langid"))
		tblConsumer.Update
		tblConsumer.Close
		Set tblConsumer = Nothing
		If Request("SelWor") <> "" Then
			Set rsWork = Server.CreateObject("ADODB.RecordSet")
			sqlWork = "SELECT * FROM [ConWork_t]"
			rsWork.Open sqlWork, g_strCONN, 1, 3
			rsWork.AddNew
			rsWork("WID") = GetWIndex(Request("SelWor"))
			rsWork("CID") = Request("tmpID")
			rsWork.Update
			rsWork.close
			Set rsWork = Nothing
		End If	
		If Request("SelCM") <> "" Then
			Set rsCM = Server.CreateObject("ADODB.RecordSet")
			sqlCM = "SELECT * FROM [CMCon_t]"
			rsCM.Open sqlCM, g_strCONN, 1, 3
			rsCM.AddNew
			rsCM("CMID") = Request("SelCM")
			rsCM("CID") = Request("tmpID")
			rsCM.Update
			rsCM.close
			Set rsCM = Nothing
		End If
		If Request("SelR") <> "" Then
			Set rsR = Server.CreateObject("ADODB.RecordSet")
			sqlR = "SELECT * FROM [ConRep_t]"
			rsR.Open sqlR, g_strCONN, 1, 3
			rsR.AddNew
			rsR("RID") = Request("SelR")
			rsR("CID") = Request("tmpID")
			rsR.Update
			rsR.close
			Set rsR = Nothing
		End If	
		If Request("SelWorBack") <> "" Then
			Set rsWork = Server.CreateObject("ADODB.RecordSet")
			sqlWork = "SELECT * FROM [ConWorkBack_t]"
			rsWork.Open sqlWork, g_strCONN, 1, 3
			rsWork.AddNew
			rsWork("WID") = Request("SelWorBack")
			rsWork("CID") = Request("tmpID")
			rsWork.Update
			rsWork.close
			Set rsWork = Nothing
		End If	
		Response.Redirect "A_Consumer.asp?MNum=" & Request("MCNum")
	ElseIf Request("page") = 2 Then
		Set tblStatus = Server.CreateObject("ADODB.RecordSet")
		sqlStatus = "SELECT * FROM C_Status_t WHERE Medicaid_Number = '" & Request("MNum") & "' "
		tblStatus.Open sqlStatus, g_strCONN, 1, 3
		If Not tblStatus.EOF Then
			If Request("chkActive") = 1 Then
				tblStatus("Active") = False
				
				If Request("Inactive") <> "" Then
					If IsDate(Request("Inactive")) Then
						tblStatus("Inactive_Date") = Z_FormatDate(Request("Inactive"))
					Else
						Session("MSG") = Session("MSG") & "<br>" & "Enter a valid date in Inactive Date."
						tblStatus("Inactive_Date") = Empty
					End If
				End If
				If Request("chkNurse") <> "" Then 
					tblStatus("Enter_Nursing_Home") = True 
				Else
					tblStatus("Enter_Nursing_Home") = False
				End If
				If Request("chkDir") <> "" Then 
					tblStatus("Unable_Self_Direct") = True 
				Else
					tblStatus("Unable_Self_Direct") = False
				End If
				If Request("chkWork") <> "" Then 
					tblStatus("Unable_Suitable_Worker") = True 
				Else
					tblStatus("Unable_Suitable_Worker") = False
				End If
				If Request("chkDeath") <> "" Then 
					tblStatus("Death") = True 
				Else
					tblStatus("Death") = False
				End If
				tblStatus("A_Other") = Request("Others")
				tblStatus.Update
			Else 
				tblStatus("Active") = True
				tblStatus("Inactive_Date") = Empty
				tblStatus("Enter_Nursing_Home") = False
				tblStatus("Unable_Self_Direct") = False
				tblStatus("Unable_Suitable_Worker") = False
				tblStatus("Death") = False
				tblStatus("A_Other") = ""
				tblStatus.Update
			End If
			'If Request("chkOnHold") = 1 Then
			'	'tblStatus("On_Hold") = True
			'	'tblStatus("Status") = "OnHold"
			'	tblStatus("onHold") = True
			'	If Request("chkHos") <> "" Then 
			'		tblStatus("In_Hospital") = True 
			'	Else
			'		tblStatus("In_Hospital") = False
			'	End If
			'	If Request("chkNew") <> "" Then 
			'		tblStatus("New_Worker") = True 
			'	Else
			'		tblStatus("New_Worker") = False
			'	End If
			'	tblStatus("H_Other") = Request("Ot")
			'	If Request("frmDate") <> "" Then
			'		If Not IsDate(Request("frmDate")) Then
			'			Session("MSG") = Session("MSG") & "<br>" & "Invalid 'From Date' in Services Temporarily on Hold."
			'		Else
			'			tblStatus("H_From_Date") = Z_FormatDate(Request("frmDate"))
			'		End If
			'		If Request("toDate") <> "" Then
			'			If Not IsDate(Request("toDate")) Then
			'				Session("MSG") = Session("MSG") & "<br>" & "Invalid 'To Date' in Services Temporarily on Hold."
			'			Else
			'				tblStatus("H_to_Date") = Z_FormatDate(Request("toDate"))
			'				If CDate(Request("frmDate")) > CDate(Request("toDate")) Then
			'					Session("MSG") = Session("MSG") & "<br>" & "Invalid dates in Services Temporarily on Hold."
			'					tblStatus("H_From_Date") = Empty
			'					tblStatus("H_to_Date") = Empty
			'				Else
			'					tblStatus("H_From_Date") = Z_FormatDate(Request("frmDate"))
			'					tblStatus("H_to_Date") = Z_FormatDate(Request("toDate"))
			'				End If
			'			End If 
			'			
			'		
			'		End If
			'		If Request("toDate") = "" Then tblStatus("H_to_Date") = Empty
			'	Else
			'		tblStatus("H_From_Date") = Empty
			'		tblStatus("H_To_Date") = Empty
			'	End If
			'	
			'Else
			'	'tblStatus("On_Hold") = False
			'	tblStatus("onHold") = False
			'	If Request("chkActive") <> "" Then
			'		tblStatus("Status") = "Inactive"
			'	Else
			'		tblStatus("Status") = "Active"
			'	End If
			'	tblStatus("Status") = "Inactive"
			'	tblStatus("In_Hospital") = False
			'	tblStatus("New_Worker") = False
			'	tblStatus("H_Other") = ""
			'	tblStatus("H_To_Date") = Empty
			'	tblStatus("H_From_Date") = Empty
			'End If
			tblStatus.Update
		End If
		tblStatus.Close
		Set tblStatus = Nothing
		'on hold new
		Set rsHold = Server.CreateObject("ADODB.RecordSet")
		rsHold.Open "SELECT TOP 1 * FROM C_OnHold_T WHERE Cid = '" & Request("MNum") & "' ORDER BY [datestamp] DESC", g_strCONN, 1, 3
		If rsHold.EOF Then '1st instance
			If Request("chkOnHold") = 1 Then
				rsHold.AddNew
				rsHold("cid") = Request("MNum")
				rsHold("DateStamp") = Now
				rsHold("On_Hold") = True
				If Request("chkHos") <> "" Then 
					rsHold("In_Hospital") = True 
				Else
					rsHold("In_Hospital") = False
				End If
				If Request("chkNew") <> "" Then 
					rsHold("New_Worker") = True 
				Else
					rsHold("New_Worker") = False
				End If
				rsHold("H_Other") = Z_FixNull(Request("Ot"))
				If Request("frmDate") <> "" Then
					If Not IsDate(Request("frmDate")) Then
						Session("MSG") = Session("MSG") & "<br>" & "Invalid 'From Date' in Services Temporarily on Hold."
					Else
						rsHold("H_From_Date") = Z_FormatDate(Request("frmDate"))
					End If
					If Request("toDate") <> "" Then
						If Not IsDate(Request("toDate")) Then
							Session("MSG") = Session("MSG") & "<br>" & "Invalid 'To Date' in Services Temporarily on Hold."
						Else
							rsHold("H_to_Date") = Z_FormatDate(Request("toDate"))
							If CDate(Request("frmDate")) > CDate(Request("toDate")) Then
								Session("MSG") = Session("MSG") & "<br>" & "Invalid dates in Services Temporarily on Hold."
								rsHold("H_From_Date") = Empty
								rsHold("H_to_Date") = Empty
							Else
								rsHold("H_From_Date") = Z_FormatDate(Request("frmDate"))
								rsHold("H_to_Date") = Z_FormatDate(Request("toDate"))
							End If
						End If 
					End If
					If Request("toDate") = "" Then rsHold("H_to_Date") = Empty
				Else
					rsHold("H_From_Date") = Empty
					rsHold("H_To_Date") = Empty
				End If
				rsHold.Update
			End If
		Else
			'compare
			theSame = True
			
			If Request("chkOnHold") = 1 Then
				If Not rsHold("On_Hold") Then 
					holdsame = False
				Else
					holdsame = True
				End If
			Else
				If Not rsHold("On_Hold") Then 
					holdsame = True
				Else
					holdsame = False
				End If
			End If
			If Z_DateNull(rsHold("H_From_Date")) <> Z_DateNull(Request("frmDate")) Then 
				FdteSame = False
			Else
				FdteSame = True
			End If
			If Z_DateNull(rsHold("H_to_Date")) <> Z_DateNull(Request("toDate")) Then 
				tdteSame = False
			Else
				tdteSame = True
			End If
			If Request("chkHos") = 1 Then
				If Not rsHold("In_Hospital") Then 
					hosSame = False
				Else
					hosSame = True
				End If
			Else
				If Not rsHold("In_Hospital") Then 
					hosSame = True
				Else
					hosSame = False
				End If
			End If
			If Request("chkNew") = 1 Then
				If Not rsHold("New_Worker") Then 
					NewSame = False
				Else
					NewSame = True
				End If
			Else
				If Not rsHold("New_Worker") Then 
					NewSame = True
				Else
					NewSame = False
				End If
			End If
			If Z_FixNull(Trim(UCase(Request("Ot")))) <> Z_FixNull(Trim(UCase(rsHold("H_Other")))) Then 
				othrSame = False
			Else
				othrSame = True
			End If
			If Not (holdsame And FdteSame And tdteSame And HosSame And NewSame And othrSame) Then theSame = False
			If Not theSame Then
				rsHold.AddNew
				rsHold("cid") = Request("MNum")
				rsHold("DateStamp") = Now
				If Request("chkOnHold") <> "" Then 
					rsHold("On_Hold") = True 
				
					If Request("chkHos") <> "" Then 
						rsHold("In_Hospital") = True 
					Else
						rsHold("In_Hospital") = False
					End If
					If Request("chkNew") <> "" Then 
						rsHold("New_Worker") = True 
					Else
						rsHold("New_Worker") = False
					End If
					rsHold("H_Other") = Z_FixNull(Request("Ot"))
					If Request("frmDate") <> "" Then
						If Not IsDate(Request("frmDate")) Then
							Session("MSG") = Session("MSG") & "<br>" & "Invalid 'From Date' in Services Temporarily on Hold."
						Else
							rsHold("H_From_Date") = Z_FormatDate(Request("frmDate"))
						End If
						If Request("toDate") <> "" Then
							If Not IsDate(Request("toDate")) Then
								Session("MSG") = Session("MSG") & "<br>" & "Invalid 'To Date' in Services Temporarily on Hold."
							Else
								rsHold("H_to_Date") = Z_FormatDate(Request("toDate"))
								If CDate(Request("frmDate")) > CDate(Request("toDate")) Then
									Session("MSG") = Session("MSG") & "<br>" & "Invalid dates in Services Temporarily on Hold."
									rsHold("H_From_Date") = Empty
									rsHold("H_to_Date") = Empty
								Else
									rsHold("H_From_Date") = Z_FormatDate(Request("frmDate"))
									rsHold("H_to_Date") = Z_FormatDate(Request("toDate"))
								End If
							End If 
						End If
						If Request("toDate") = "" Then rsHold("H_to_Date") = Empty
					Else
						rsHold("H_From_Date") = Empty
						rsHold("H_To_Date") = Empty
					End If
				Else
					rsHold("On_Hold") = False
					rsHold("In_Hospital") = False
					rsHold("New_Worker") = False
					rsHold("H_Other") = ""
					rsHold("H_From_Date") = Empty
					rsHold("H_To_Date") = Empty
				End If
				rsHold.Update
			End If
		End If
		rsHold.Close
		Set rsHold = Nothing
		'hosp
		Set tblEMP4 = Server.CreateObject("ADODB.Recordset")
		strSQL4 = "SELECT * FROM [C_Site_Visit_Dates_t] WHERE Medicaid_Number ='" & Request("MNum") & "' AND NOT hospdate IS NULL"
		'On Error Resume Next
		tblEMP4.Open strSQL4, g_strCONN, 1, 3
		ctr = Request("ctr4")
		For i = 0 to ctr - 1 
			tmpctr = Request("chkhosp" & i)
			If tmpctr <> "" Then
				strTmp = "Index= " & tmpctr 
				tblEMP4.Find(strTmp)
				If Not tblEMP4.EOF Then
					If IsDate(Request("txthospdate" & i)) Then 
						tblEMP4("hospdate") = Request("txthospdate" & i)
					Else
						Session("MSG") = Session("MSG") & "Invalid date in Hospitalization log. " & tblEMP4("hospdate") & " changed to " & Request("txthospdate" & i) & ".<br>"
					End If
					tblEMP4("hospcom") = Request("hospcom" & i)
					tblEMP4.Update
				End If
			End If
			tblEMP4.MoveFirst
		Next
		tblEMP4.Close
		Set tblEMP4 = Nothing
		If Request("txtHospdate") <> "" Then
			If IsDate(Request("txtHospdate")) Then
				Set tblPC = Server.CreateObject("ADODB.RecordSet")
				sqlPC = "SELECT * FROM [C_Site_Visit_Dates_t] WHERE [index] = 0 "
				tblPC.Open sqlPC, g_strCONN, 1, 3
				tblPC.AddNew
				tblPC("Medicaid_Number") = Request("MNum")
				tblPC("hospdate") = Request("txtHospdate")
				If Request("hospcom") <> "" Then 
					tblPC("hospcom") = Request("hospcom")
				End If
				tblPC.Update
				tblPC.Close
				Set tblPC = Nothing
			Else
				Session("MSG") = Session("MSG") & "Invalid Entry in Hospitalizations."
			End If
		End If
		'sup
		Set tblEMP4 = Server.CreateObject("ADODB.Recordset")
		strSQL4 = "SELECT * FROM [C_Site_Visit_Dates_t] WHERE Medicaid_Number ='" & Request("MNum") & "' AND NOT supdate IS NULL"
		'On Error Resume Next
		tblEMP4.Open strSQL4, g_strCONN, 1, 3
		ctr = Request("ctr5")
		For i = 0 to ctr - 1 
			tmpctr = Request("chksup" & i)
			If tmpctr <> "" Then
				strTmp = "Index= " & tmpctr 
				tblEMP4.Find(strTmp)
				If Not tblEMP4.EOF Then
					If IsDate(Request("txtsupdate" & i)) Then 
						tblEMP4("supdate") = Request("txtsupdate" & i)
					Else
						Session("MSG") = Session("MSG") & "Invalid date in Supervisory Notes log. " & tblEMP4("supdate") & " changed to " & Request("txtsupdate" & i) & ".<br>"
					End If
					tblEMP4("supnotes") = Request("supcom" & i)
					tblEMP4.Update
				End If
			End If
			tblEMP4.MoveFirst
		Next
		tblEMP4.Close
		Set tblEMP4 = Nothing
		If Request("txtsupdate") <> "" Then
			If IsDate(Request("txtsupdate")) Then
				Set tblPC = Server.CreateObject("ADODB.RecordSet")
				sqlPC = "SELECT * FROM [C_Site_Visit_Dates_t] WHERE [index] = 0 "
				tblPC.Open sqlPC, g_strCONN, 1, 3
				tblPC.AddNew
				tblPC("Medicaid_Number") = Request("MNum")
				tblPC("supdate") = Request("txtsupdate")
				If Request("supcom") <> "" Then 
					tblPC("supnotes") = Request("supcom")
				End If
				tblPC.Update
				tblPC.Close
				Set tblPC = Nothing
			Else
				Session("MSG") = Session("MSG") & "Invalid Entry in Supervisory Notes."
			End If
		End If
		Response.Redirect "A_C_Status.asp?MNum=" & Request("MNum")
	ElseIf Request("page") = 3 Then

	Set tblHealth = Server.CreateObject("ADODB.RecordSet")
		sqlHealth = "SELECT * FROM C_Health_t WHERE Medicaid_Number = '" & Request("MNum") & "' "
		tblHealth.Open sqlHealth, g_strCONN, 1, 3
		'If Not tblHealth.EOF Then
			''''''''''''''''''''''''''''''''''''''''''''
			'Age
			'If Request("age") < 50 Then
			'	RAge = 0
			'ElseIf Request("age") >= 50 And Request("age") <= 59 Then
			'	RAge = 1
			'ElseIf Request("age") >= 60 And Request("age") <= 69 Then
			'	RAge = 2
			'ElseIf Request("age") >= 70 And Request("age") <= 79 Then
			'	RAge = 3
			'ElseIf Request("age") >= 80 Then
			'	RAge = 4
			'End If
			
			'Ambulation
			'Select Case Request("Amb")
			'	Case 1 RAmb = 0
			'	Case 2 RAmb = 1
			'	Case 3 RAmb = 2
			'	Case 4 RAmb = 3
			'	Case 5 RAmb = 4
			'	Case Else Ramb = 0
			'End Select
			
			'ADL
			'Select Case Request("ADL")
			'	Case 1 RADL = 0
			'	Case 2 RADL = 1
			'	Case 3 RADL = 2
			'	Case 4 RADL = 3
			'	Case 5 RADL = 4
			'	Case Else RADL = 0
			'End Select
			
			'Others
			'ctr = 0
			'If Request("ChkUse") <> "" Then
			'	ctr = ctr + 1
			'End If
			'If Request("ChkMH") <> "" Then
			'	ctr = ctr + 1
			'End If
			'If Request("ChkDrug") <> "" Then
			'	ctr = ctr + 1
			'End If
			'If Request("ChkIso") <> "" Then
			'	ctr = ctr + 1
			'End If
			'If Request("ChkDem") <> "" Then
			'	ctr = ctr + 1
			'End If
			'If Request("ChkTerm") <> "" Then
			'	ctr = ctr + 1
			'End If
			'If Request("ChkTob") <> "" Then
			'	ctr = ctr + 1
			'End If
			'If Request("ChkObes") <> "" Then
			'	ctr = ctr + 1
			'End If
			'If Request("ChkPar") <> "" Then
			'	ctr = ctr + 1
			'End If
			'If Request("ChkQuad") <> "" Then
			'	ctr = ctr + 1
			'End If
			'If ctr < 3 Then 
			'	ROthers = ctr
			'ElseIf ctr >= 3 Then
			'	ROthers = 4
			'End If 
			
			'tmpRate = CDbl((RAge + RAmb + RADL + ROthers) / 4)

			'tblHealth("Rating") = tmpRate
			'''''''''''''''''''''''''''''''''''''''''
			tblHealth.AddNew
			tblHealth("datestamp") = Now
			tblHealth("Medicaid_Number") = Request("MNum")
			'Ambulation
			tblHealth("Indept") = False
			tblHealth("Cane") = False
			tblHealth("Walker") = False
			tblHealth("Walk") = False
			tblHealth("WheelChair") = False
			If Request("Amb") = 1 Then
				tblHealth("Indept") = True
			ElseIf Request("Amb") = 2 Then
				tblHealth("Cane") = True
			ElseIf Request("Amb") = 3 Then
				tblHealth("Walker") = True
			ElseIf Request("Amb") = 4 Then
				tblHealth("Walk") = True
			ElseIf Request("Amb") = 5 Then
				tblHealth("WheelChair") = True
			End If
			'ADL
			tblHealth("ADL_Indep") = False
			tblHealth("Monitor") = False
			tblHealth("MinAss") = False
			tblHealth("Ass") = False
			tblHealth("Complete") = False
			If Request("ADL") = 1 Then
				tblHealth("ADL_Indep") = True
			ElseIf Request("ADL") = 2 Then
				tblHealth("Monitor") = True
			ElseIf Request("ADL") = 3 Then
				tblHealth("MinAss") = True
			ElseIf Request("ADL") = 4 Then
				tblHealth("Ass") = True
			ElseIf Request("ADL") = 5 Then
				tblHealth("Complete") = True
			End If
			'new
			tblHealth("oxy") = False
			tblHealth("gait") = False
			tblHealth("blance") = False
			tblHealth("fear") = False
			tblHealth("furni") = False
			tblHealth("noncompl") = False
			tblHealth("substance") = False
			tblHealth("chew") = False
			tblHealth("teeth") = False
			tblHealth("specdiet") = False
			tblHealth("allergy") = False
			tblHealth("diabetes") = False
			tblHealth("hyper") = False
			tblHealth("hypo") = False
			tblHealth("heartdis") = False
			tblHealth("heartfail") = False
			tblHealth("deepvain") = False
			tblHealth("hyperten") = False
			tblHealth("hypoten") = False
			tblHealth("neuro") = False
			tblHealth("vasc") = False
			tblHealth("gerd") = False
			tblHealth("ulcers") = False
			tblHealth("arth") = False
			tblHealth("hip") = False
			tblHealth("limb") = False
			tblHealth("osteo") = False
			tblHealth("bone") = False
			tblHealth("als") = False
			tblHealth("cereb") = False
			tblHealth("stroke") = False
			tblHealth("dementia") = False
			tblHealth("hunting") = False
			tblHealth("sclerosis") = False
			tblHealth("parapal") = False
			tblHealth("park") = False
			tblHealth("quadri") = False
			tblHealth("seize") = False
			tblHealth("TIA") = False
			tblHealth("Trauma") = False
			tblHealth("anx") = False
			tblHealth("depress") = False
			tblHealth("bipolar") = False
			tblHealth("schiz") = False
			tblHealth("abuse") = False
			tblHealth("otherpsy") = False
			tblHealth("asthma") = False
			tblHealth("copd") = False
			tblHealth("TB") = False
			tblHealth("cat") = False
			tblHealth("retin") = False
			tblHealth("glaucoma") = False
			tblHealth("macular") = False
			tblHealth("hear") = False
			tblHealth("allergies") = False
			tblHealth("anemia") = False
			tblHealth("cancer") = False
			tblHealth("devdis") = False
			tblHealth("morbid") = False
			tblHealth("renal") = False
			tblHealth("otherdiag") = False
			tblHealth("housekeep") = False
			tblHealth("laundry") = False
			tblHealth("meal") = False
			tblHealth("grocery") = False
			tblHealth("dress") = False
			tblHealth("eat") = False
			tblHealth("asstwalk") = False
			tblHealth("asstwheel") = False
			tblHealth("asstmotor") = False
			tblHealth("commeal") = False
			tblHealth("medical") = False
			tblHealth("shower") = False
			tblHealth("tub") = False
			tblHealth("oral") = False
			tblHealth("commode") = False
			tblHealth("sit") = False
			tblHealth("medication") = False
			tblHealth("undress") = False
			tblHealth("shampoosink") = False
			tblHealth("oralcare") = False
			tblHealth("massage") = False
			tblHealth("shampoobed") = False
			tblHealth("shave") = False
			tblHealth("bedbath") = False
			tblHealth("bedpan") = False
			tblHealth("ptexer") = False
			tblHealth("hoyer") = False
			tblHealth("eye") = False
			tblHealth("transferbelt") = False
			tblHealth("alz") = False
			tblHealth("Incontinence") = False
			tblHealth("Hospice") = False
			tblHealth("lotion") = False
			tblHealth("whatdiet") = "" 
			tblHealth("whatallergy") = ""
			tblHealth("whatdiag") = ""
			tblHealth("whatallergies") = ""
			
			If Request("oxy") = 1 Then tblHealth("oxy")	=	True
			If Request("gait") = 1 Then tblHealth("gait")	=	True
			If Request("blance") = 1 Then tblHealth("blance")	=	True
			If Request("fear") = 1 Then tblHealth("fear")	=	True
			If Request("furni") = 1 Then tblHealth("furni")	=	True
			If Request("noncompl") = 1 Then tblHealth("noncompl")	=	True
			If Request("substance") = 1 Then tblHealth("substance")	=	True
			If Request("chew") = 1 Then tblHealth("chew")	=	True
			If Request("teeth") = 1 Then tblHealth("teeth")	=	True
			If Request("specdiet") = 1 Then tblHealth("specdiet")	=	True
			If Request("specdiet") = 1 Then tblHealth("whatdiet") = Request("whatdiet")					
			If Request("allergy") = 1 Then tblHealth("allergy")	=	True
			If Request("allergy") = 1 Then tblHealth("whatallergy") = Request("whatallergy")					
			If Request("diabetes") = 1 Then tblHealth("diabetes")	=	True
			If Request("hyper") = 1 Then tblHealth("hyper")	=	True
			If Request("hypo") = 1 Then tblHealth("hypo")	=	True
			If Request("heartdis") = 1 Then tblHealth("heartdis")	=	True
			If Request("heartfail") = 1 Then tblHealth("heartfail")	=	True
			If Request("deepvain") = 1 Then tblHealth("deepvain")	=	True
			If Request("hyperten") = 1 Then tblHealth("hyperten")	=	True
			If Request("hypoten") = 1 Then tblHealth("hypoten")	=	True
			If Request("neuro") = 1 Then tblHealth("neuro")	=	True
			If Request("vasc") = 1 Then tblHealth("vasc")	=	True
			If Request("gerd") = 1 Then tblHealth("gerd")	=	True
			If Request("ulcers") = 1 Then tblHealth("ulcers")	=	True
			If Request("arth") = 1 Then tblHealth("arth")	=	True
			If Request("hip") = 1 Then tblHealth("hip")	=	True
			If Request("limb") = 1 Then tblHealth("limb")	=	chew
			If Request("osteo") = 1 Then tblHealth("osteo")	=	True
			If Request("bone") = 1 Then tblHealth("bone")	=	True
			If Request("als") = 1 Then tblHealth("als")	=	True
			If Request("cereb") = 1 Then tblHealth("cereb")	=	True
			If Request("stroke") = 1 Then tblHealth("stroke")	=	True
			If Request("dementia") = 1 Then tblHealth("dementia")	=	True
			If Request("hunting") = 1 Then tblHealth("hunting")	=	True
			If Request("sclerosis") = 1 Then tblHealth("sclerosis")	=	True
			If Request("parapal") = 1 Then tblHealth("parapal")	=	True
			If Request("park") = 1 Then tblHealth("park")	=	chew
			If Request("quadri") = 1 Then tblHealth("quadri")	=	True
			If Request("seize") = 1 Then tblHealth("seize")	=	True
			If Request("TIA") = 1 Then tblHealth("TIA")	=	True
			If Request("Trauma") = 1 Then tblHealth("Trauma")	=	True
			If Request("anx") = 1 Then tblHealth("anx")	=	True
			If Request("depress") = 1 Then tblHealth("depress")	=	True
			If Request("bipolar") = 1 Then tblHealth("bipolar")	=	True
			If Request("schiz") = 1 Then tblHealth("schiz")	=	True
			If Request("abuse") = 1 Then tblHealth("abuse")	=	True
			If Request("otherpsy") = 1 Then tblHealth("otherpsy")	=	True
			If Request("asthma") = 1 Then tblHealth("asthma")	=	True
			If Request("copd") = 1 Then tblHealth("copd")	=	True
			If Request("TB") = 1 Then tblHealth("TB")	=	True
			If Request("cat") = 1 Then tblHealth("cat")	=	True
			If Request("retin") = 1 Then tblHealth("retin")	=	True
			If Request("glaucoma") = 1 Then tblHealth("glaucoma")	=	True
			If Request("macular") = 1 Then tblHealth("macular")	=	True
			If Request("hear") = 1 Then tblHealth("hear")	=	True
			If Request("allergies") = 1 Then tblHealth("allergies")	=	True
			If Request("allergies") = 1 Then tblHealth("whatallergies") = Request("whatallergies")					
			If Request("anemia") = 1 Then tblHealth("anemia")	=	True
			If Request("cancer") = 1 Then tblHealth("cancer")	=	True
			If Request("devdis") = 1 Then tblHealth("devdis")	=	True
			If Request("morbid") = 1 Then tblHealth("morbid")	=	True
			If Request("renal") = 1 Then tblHealth("renal")	=	True
			If Request("otherdiag") = 1 Then tblHealth("otherdiag")	=	True
			If Request("otherdiag") = 1 Then tblHealth("whatdiag") = Request("whatdiag")	
			If Request("housekeep") = 1 Then tblHealth("housekeep")	=	True
			If Request("laundry") = 1 Then tblHealth("laundry")	=	True
			If Request("meal") = 1 Then tblHealth("meal")	=	True
			If Request("grocery") = 1 Then tblHealth("grocery")	=	True
			If Request("dress") = 1 Then tblHealth("dress")	=	True
			If Request("eat") = 1 Then tblHealth("eat")	=	True
			If Request("asstwalk") = 1 Then tblHealth("asstwalk")	=	True
			If Request("asstwheel") = 1 Then tblHealth("asstwheel")	=	True
			If Request("asstmotor") = 1 Then tblHealth("asstmotor")	=	True
			If Request("commeal") = 1 Then tblHealth("commeal")	=	True
			If Request("medical") = 1 Then tblHealth("medical")	=	True
			If Request("shower") = 1 Then tblHealth("shower")	=	True
			If Request("tub") = 1 Then tblHealth("tub")	=	True
			If Request("oral") = 1 Then tblHealth("oral")	=	True
			If Request("commode") = 1 Then tblHealth("commode")	=	True
			If Request("sit") = 1 Then tblHealth("sit")	=	True
			If Request("medication") = 1 Then tblHealth("medication")	=	True
			If Request("undress") = 1 Then tblHealth("undress")	=	True
			If Request("shampoosink") = 1 Then tblHealth("shampoosink")	=	True
			If Request("oralcare") = 1 Then tblHealth("oralcare")	=	True
			If Request("massage") = 1 Then tblHealth("massage")	=	True
			If Request("shampoobed") = 1 Then tblHealth("shampoobed")	=	True
			If Request("shave") = 1 Then tblHealth("shave")	=	True
			If Request("bedbath") = 1 Then tblHealth("bedbath")	=	True
			If Request("bedpan") = 1 Then tblHealth("bedpan")	=	True
			If Request("ptexer") = 1 Then tblHealth("ptexer")	=	True
			If Request("hoyer") = 1 Then tblHealth("hoyer")	=	True
			If Request("eye") = 1 Then tblHealth("eye")	=	True
			If Request("transferbelt") = 1 Then tblHealth("transferbelt")	=	True
			If Request("alz") = 1 Then tblHealth("alz")	=	True
			If Request("Incontinence") = 1 Then tblHealth("Incontinence")	=	True
			If Request("Hospice") = 1 Then tblHealth("Hospice")	=	True
			If Request("lotion") = 1 Then tblHealth("lotion")	=	True
			tblHealth.Update
		'End If
		tblHealth.Close
		Set tblHealth = Nothing	
		'''''''DIAGNOSIS
		'If Request("Diag") <> "" Then
		'	Set rsDiag = Server.CreateObject("ADODB.RecordSet")
		'	sqlDiag = "SELECT * FROM C_Diagnosis_t"
		'	rsDiag.Open sqlDiag, g_strCONN, 1, 3
		'	rsDiag.AddNew
		'	rsDiag("Medicaid_Number") = Request("MNum")
		'	rsDiag("Diagnosis") = Request("Diag")
		'	rsDiag.Update
		'	rsDiag.Close
		'	Set rsDiag = Nothing
		'End If
		Response.Redirect "A_C_Health.asp?MNum=" & Request("MNum")
	ElseIf Request("page") = 4 Then
		Set tblFiles = Server.CreateObject("ADODB.Recordset")
		sqlFiles = "SELECT * FROM C_Files_t WHERE Medicaid_Number = '" & Request("MNum") & "' "
		tblFiles.Open sqlFiles, g_strCONN, 1, 3
			If Not tblFiles.EOF Then
				If Request("chkDEAS") <> "" Then 
					tblFiles("DEAS_Service_Plan") = True 
				Else
					tblFiles("DEAS_Service_Plan") = False
				End If
				If Request("chkLSS") <> "" Then 
					tblFiles("LSS_Care_Plan") = True
				Else
					tblFiles("LSS_Care_Plan") = False
				End If
				If Request("chkVR") <> "" Then 
					tblFiles("Vehicle_Release") = True 
				Else
					tblFiles("Vehicle_Release") = False
				End If
				If Request("chkPS") <> "" Then 
					tblFiles("Privacy_Statement") = True
				Else
					tblFiles("Privacy_Statement") = False
				End If
				If Request("chkARF") <> "" Then 
					tblFiles("A_Representative_Form") = True
				Else
					tblFiles("A_Representative_Form") = False
				End If
				If Request("chkRRO") <> "" Then 
					tblFiles("Roles_Respon_Outline") = True
				Else
					tblFiles("Roles_Respon_Outline") = False
				End If
				If Request("chkCSSC") <> "" Then 
					tblFiles("C_Site_Safety_Check") = True
				Else
					tblFiles("C_Site_Safety_Check") = False
				End If
				If Request("chkTR") <> "" Then 
					tblFiles("Training_Requirements") = True
				Else
					tblFiles("Training_Requirements") = False
				End If
				If Request("chkPhoto") <> "" Then 
					tblFiles("Photo") = True
				Else
					tblFiles("Photo") = False
				End If
				If Request("chkAS") <> "" Then 
					tblFiles("AckStmt") = True
				Else
					tblFiles("AckStmt") = False
				End If
				If Request("chkARepForm") <> "" Then 
					tblFiles("ARepForm") = True
					If Request("DFax") <> "" Then
					If IsDate(Request("DFax")) Then
						tblFiles("DteFax") = Z_FormatDate(Request("DFax"))
					Else
						Session("MSG") = "Enter Valid Date in Faxed to Case Manager."
						Response.Redirect "A_C_Files.asp?MNum=" & Request("MNum")
					End If
				End If
				Else
					tblFiles("ARepForm") = False
					tblFiles("DteFax") = empty
				End If
				
				tblFiles.Update
			End If
		tblFiles.Close
		Set tblFiles = Nothing
		Response.Redirect "A_C_Files.asp?MNum=" & Request("MNum")
	ElseIf Request("page") = 5 Then
		Set tblEMP = Server.CreateObject("ADODB.Recordset")
		strSQL = "SELECT * FROM [C_Diagnosis_t]"
		'On Error Resume Next
		tblEMP.Open strSQL, g_strCONN, 1, 3
		ctr = Request("ctr")
			For i = 0 to ctr - 1 
				tmpctr = Request("chkDiag" & i)
				If tmpctr <> "" Then
					strTmp = "Index= " & tmpctr 
					tblEMP.Find(strTmp)
					If Not tblEMP.EOF Then
						tblEmp.DELETE
						tblEmp.Update
					End If
				End If
				tblEmp.MoveFirst
				
			Next
			 
		tblEmp.Close
		Set tblEMP = Nothing
		Response.Redirect "A_C_Health.asp?MNum=" & Request("MNum")
	End If	
%>