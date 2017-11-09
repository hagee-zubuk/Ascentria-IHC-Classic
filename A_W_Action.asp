<%language=vbscript%>
<!-- #include file="_Files.asp" -->

<!-- #include file="_Utils.asp" -->
<%
Function GetWIndex(xxx)
	Set rsI = Server.CreateObject("ADODB.RecordSet")
	sqlI = "SELECT [index] FROM Worker_T WHERE Social_Security_Number = '" & xxx & "'"
	rsI.Open sqlI, g_strCONN, 3, 1
	If Not rsI.EOF Then
		GetWIndex = rsI("Index")
	End If
	rsI.Close
	Set rsI = Nothing
End Function
	If Request("page") = 1 Then
		Set tblWorker = Server.CreateObject("ADODB.RecordSet")
		
		If Request("new") <> 1 Then
			If request("WID") = ""  Or Request("SSN") = "" then 
				Session("MSG") = "Sorry. To maintain integrity of database, please Sign-in again."
				Response.Redirect "default.asp"
			end if
			If Request("WID") = Request("SSN") Then
				sqlWorker = "SELECT * FROM Worker_t WHERE Social_Security_Number = '" & Request("WID") & "' "
			Else
				sqlWorker = "SELECT * FROM Worker_t WHERE Social_Security_Number = '" & Request("WID") & "' "
			End If
		Else
			Session("WFiles") = Z_DoEncrypt(Request("SSN") & "|" & Request("lname") & "|" & Request("fname") & "|" & _
				Request("Addr") & "|" & Request("Gen") & "|" & Request("DOB") & "|" & Request("PhoneNo") & "|" & _
				Request("CellNo") & "|" & Request("DateHired") & "|" & Request("Stat") & "|" & Request("SepCode") & "|" & _
				Request("chkDrive") & "|" & Request("chkonFile") & "|" & Request("LisNo") & "|" & Request("LisExpDte") & "|" & _
				Request("chkIns") & "|" & Request("Insdate") & "|" & Request("chkTown") & "|" & Request("Towns") & "|" & _
				Request("cty") & "|" & Request("ste") & "|" & Request("zcode") & "|" & Request("Termd8") & "|" & Request("mAddr") & _
				"|" & Request("mcty") & "|" & Request("mste") & "|" & Request("mzcode") & "|" & Request("txtSal") & "|" & Request("chkMail") & _
				"|" & Request("eMail")& "|" & Request("DateManual")& "|" & Request("txtFN") & "|" & Request("chkpp")) & "|" & Request("langid")
			If Request("SSN") = "" Then
				Session("MSG") = "Social Security Number is blank."
				Response.Redirect "A_New_Worker.asp"
			End If
			sqlWorker = "SELECT * FROM Worker_t WHERE Social_Security_Number = '" & Trim(Request("SSN")) & "' "
		End If
		tblWorker.Open sqlWorker, g_strCONN, 1, 3
		If Request("new") = 1 Then
			If tblWorker.EOF Then
				If Request("SSN") <> "" Then
					tblWorker.AddNew
					tblWorker("Social_Security_Number") = Trim(Request("SSN"))
					Session("WID") = Trim(Request("SSN"))
				End If
			Else
				Session("MSG") = "Social Security Number is either blank or already exists."
				Response.Redirect "A_New_Worker.asp"
				
			End If
		else
			Session("WID") = Trim(Request("SSN"))
		End If
		If Request("DOB") <> "" Then
			If IsDate(Request("DOB")) Then
				If CDate(Request("DOB")) < Date Then 
					tblWorker("DOB") = Z_FormatDate(Request("DOB"))
				Else
					Session("MSG") = Session("MSG") & "<br>" & "Invalid DOB."
					'Response.Redirect "A_Worker.asp"
				End If
			Else
				Session("MSG") = Session("MSG") & "<br>" & "Enter a valid date For DOB."
				'Response.Redirect "A_Worker.asp"
			End If
		Else
			tblWorker("DOB") = Empty
		End If
		If Request("DateHired") <> "" Then
			If IsDate(Request("DateHired")) Then
				tblWorker("Date_Hired") = Z_FormatDate(Request("DateHired"))
			Else
				Session("MSG") = Session("MSG") & "<br>" & "Enter a valid date for Date of Hire."
			End If
		Else
			tblWorker("Date_Hired") = Empty
		End If	
		If Request("DateManual") <> "" Then
			If IsDate(Request("DateManual")) Then
				tblWorker("Manual") = Z_FormatDate(Request("DateManual"))
			Else
				Session("MSG") = Session("MSG") & "<br>" & "Enter a valid date for Employee Manual Received."
			End If
		Else
			tblWorker("Manual") = Empty
		End If	
		If Request("TermD8") <> "" Then
			If IsDate(Request("TermD8")) Then
				tblWorker("Term_date") = Z_FormatDate(Request("TermD8"))
			Else
				Session("MSG") = Session("MSG") & "<br>" & "Enter a valid date for Termination date."
			End If
		Else
			tblWorker("Term_date") = Empty
		End If	
		tblWorker("FileNum") = Request("txtFN")
		If Request("chkpp") <> "" Then
			tblWorker("privatepay") = true
		else
			tblWorker("privatepay") = false
		end if
		If Request("new") <> 1 Then 'check if name changed
			tmplname = tblWorker("Lname")
			tmpfname = tblWorker("Fname")
			if (tmplname <> Request("lname")) or (tmpfname <> Request("fname")) then
				Set fso = CreateObject("Scripting.FileSystemObject")
			'If Not fso.FileExists(WorkerList) Then
				Set NewWorkList = fso.CreateTextFile(WorkerList, true)
				Set tblLWork = Server.CreateObject("ADODB.Recordset")
				strSQLd = "SELECT * FROM [Worker_t] WHERE Status = 'Active' OR Status = 'InActive' ORDER BY [lname], [Fname]"
				tblLWork.Open strSQLd, g_strCONN, 3, 1
				tblLWork.Movefirst
				Do Until tblLWork.EOF
					NewWorkList.WriteLine tblLWork("index") & "|<option value='" & tblLWork("Social_Security_Number")& "'> "& tblLWork("Lname") & ", " & tblLWork("fname") & " </option>"
					tblLWork.Movenext
				Loop
				tblLWork.Close
				Set tblLWork = Nothing
				Set NewWorkList = Nothing
				set fso = nothing	
			end if
		end if
		tblWorker("Lname") = Request("lname")
		tblWorker("Fname") = Request("fname")
		
		tblWorker("Address") = Request("Addr")
		tblWorker("City") = Request("cty")
		tblWorker("State") = UCase(Request("ste"))
		tblWorker("zip") = Request("zcode")
		If Request("chkMail") = "" Then
			tblWorker("mAddress") = Request("mAddr")
			tblWorker("mCity") = Request("mcty")
			tblWorker("mState") = UCase(Request("mste"))
			tblWorker("mzip") = Request("mzcode")
		Else
			tblWorker("mAddress") = Request("Addr")
			tblWorker("mCity") = Request("cty")
			tblWorker("mState") = UCase(Request("ste"))
			tblWorker("mzip") = Request("zcode")
		End If
		tblWorker("Gender") = Request("Gen")
		'tblWorker("Lname") = Request("lname")
		If Request("PhoneNo") <> "" then
			If len(Request("PhoneNo")) = 12 then
				tblWorker("PhoneNo") = Request("PhoneNo")
			else
				Session("MSG") = Session("MSG") & "<br>" & "Enter a valid format for Phone Number (xxx-xxx-xxxx)."
			end if
		else
			tblWorker("PhoneNo") = ""
		end if
		If Request("cellno") <> "" then
			if len(Request("cellno")) = 12 then
				tblWorker("CellNo") = Request("CellNo")
			else
				Session("MSG") = Session("MSG") & "<br>" & "Enter a valid format for Cell Number (xxx-xxx-xxxx)."
			end if
		else
			tblWorker("CellNo") = ""
		end if
		tblWorker("Email") = Request("eMail")
		tblWorker("pm1") = Request("selpm1")
		tblWorker("pm2") = Request("selpm2")
		tblWorker("badge") = Request("txtbadge")
		tblWorker("ubadge") = Request("txtuBadge")
		tblWorker("umid") = Request("txtumid")
		tblWorker("comment") = Request("Pcomments")
		tblWorker("misdemeanor") = Request("misd")
		tblWorker("warning") = Request("warn")
		tblWorker("flt") = False
		if Request("chkflt") = 1 then tblWorker("flt") = true
		If Request("Stat") = 1 Then stat = "Active"
		If Request("Stat") = 2 Then stat = "InActive"
		If Request("Stat") = 3 Then stat = "Potential Applicant"
		tblWorker("Status") = stat
		tblWorker("Sep_Code") = ""
		If Request("Stat") = 2 Then tblWorker("Sep_Code") = Request("SepCode")
		If Request("txtSal") <> "" Then
			If IsNumeric(Request("txtSal")) Then
				tblWorker("Salary") = Request("txtSal")
			Else
				Session("MSG") = Session("MSG") & "<br>" & "Enter a valid Salary."
			End If
		Else
			tblWorker("Salary") = 0
		End If
			tblWorker("essentials") = False
			If Request("essentials") <> "" Then tblWorker("essentials") = True
		If Request("chkDrive") <> "" Then
			tblWorker("Driver") = True
			If Request("chkonFile") <> "" Then
				tblWorker("License_File") = True
			Else
				tblWorker("License_File") = False
			End If
			tblWorker("LicenseNo") = Request("LisNo")
			If Request("LisExpDte") <> "" Then 
				If IsDate(Request("LisExpDte")) Then
					If CDate(Request("LisExpDte")) > Date Then
						tblWorker("LicenseExpDate") = Z_FormatDate(Request("LisExpDte"))
					Else
						Session("MSG") = Session("MSG") & "<br>" & "License already expired."
					End If
				Else
					Session("MSG") = Session("MSG") & "<br>" & "Enter a valid date for License Expiration Date."
				End If
			Else
				tblWorker("LicenseExpDate") = Empty
			End If
			If Request("chkIns") <> "" Then
				tblWorker("AutoInsur") = True
			Else
				tblWorker("AutoInsur") = False
			End If
			If Request("Insdate") <> "" Then 
				If IsDate(Request("Insdate")) Then
					If CDate(Request("Insdate")) > Date Then
						tblWorker("InsuranceExpdate") = Z_FormatDate(Request("Insdate"))
					Else
						Session("MSG") = Session("MSG") & "<br>" & "Insurance already expired."
					End If
				Else
					Session("MSG") = Session("MSG") & "<br>" & "Enter a valid date for Insurance Expiration Date."
				End If
			Else
				tblWorker("InsuranceExpdate") = Empty
			End If
		
		Else
			tblWorker("Driver") = False
			tblWorker("License_File") = False
			tblWorker("LicenseNo") = ""
			tblWorker("LicenseExpDate") = Empty
			tblWorker("AutoInsur") = False
			tblWorker("InsuranceExpdate") = Empty
		End If
		If Request("chkTown") <> "" Then 
			tblWorker("More_Towns") = True
		Else
			tblWorker("More_Towns") = False
			Set rsTowns = Server.CreateObject("ADODB.RecordSet")
			sqlT = "DELETE FROM W_Towns_t WHERE SSN = '" & Request("WID") & "' "
			rsTowns.Open sqlT, g_strCONN, 1, 3
			Set rsTowns = Nothing
		End If
		tblWorker("prefcom") = Request("selpref")
		tblWorker("countid") = Request("selcount")
		' new -- 2017-10-27
		tblWorker("langid") = Z_CLng(Request("langid"))
		tblWorker.Update
		tblWorker.Close
		Set tblWorker = Nothing
		Set tblTowns = Server.CreateObject("ADODB.RecordSet")
		sqlTowns = "SELECT * FROM W_Towns_t WHERE SSN = '" & Request("WID") & "' "
		tblTowns.Open sqlTowns, g_strCONN, 1, 3
		If Not tblTowns.EOF Then
			ctr = Request("ctr")
			For i = 0 to ctr 
				tmpctr = Request("chk" & i)
				If tmpctr <> "" Then
					strTmp = "index= " & tmpctr 
					tblTowns.Movefirst
					tblTowns.Find(strTmp)
					If Not tblTowns.EOF Then
						tblTowns("Towns") = Request("Town" & i)
						tblTowns.Update
					End If
				End If
			Next 
		End If
		On Error Resume Next
		If Request("Towns") <> "" Then
			x = 0
			tblTowns.Movefirst
			Do Until tblTowns.EOF
				If tblTowns("Towns") = Request("Towns") Then
					x = 1
				End If
				tblTowns.MoveNext
			Loop
			If x <> 1 then
				tblTowns.AddNew
				tblTowns("SSN") = Request("WID")
				tblTowns("Towns") = Request("Towns")
				tblTowns.Update
			Else
				Session("MSG") = Session("MSG") & "<br>" & Request("Towns") & " already exist!"
			End If
		End If
		tblTowns.Close
		Set tblTowns = Nothing
		
		If Request("new") = 1 Then
			're-create worker list
		Set fso = CreateObject("Scripting.FileSystemObject")
		'If Not fso.FileExists(WorkerList) Then
			Set NewWorkList = fso.CreateTextFile(WorkerList, true)
			Set tblLWork = Server.CreateObject("ADODB.Recordset")
			strSQLd = "SELECT * FROM [Worker_t] WHERE Status = 'Active' OR Status = 'InActive' ORDER BY [lname], [Fname]"
			tblLWork.Open strSQLd, g_strCONN, 3, 1
			tblLWork.Movefirst
			Do Until tblLWork.EOF
				NewWorkList.WriteLine tblLWork("index") & "|<option value='" & tblLWork("Social_Security_Number")& "'> "& tblLWork("Lname") & ", " & tblLWork("fname") & " </option>"
				tblLWork.Movenext
			Loop
			tblLWork.Close
			Set tblLWork = Nothing
			Set NewWorkList = Nothing
			set fso = nothing
		end if
		wid = Request("WID")
		If Request("WID") = "" Then wid = Session("WID")
		Response.Redirect "A_Worker.asp?WID=" & wid
			
	ElseIf Request("page") = 2 Then
		if request("del") = 1 then
			Set tblSite = Server.CreateObject("ADODB.RecordSet")
			sqlSite = "SELECT * FROM W_log_t WHERE ssn ='" & Request("WID") & "' "
			tblSite.Open sqlSite, g_strCONN, 1, 3
			If Not tblSite.EOF Then
				If Request("ctr") <> "" Then 
					ctr = Request("ctr")
					For i = 0 to ctr 
						tmpctr = Request("chktrain" & i)
						If tmpctr <> "" Then
							strTmp = "index= " & tmpctr 
							tblSite.Movefirst
							tblSite.Find(strTmp)
							If Not tblSite.EOF Then
								tblSite.Delete
								tblSite.Update
								Session("MSG") = "Training deleted." 
							End If
						End If
					Next
				End If 
			End If
			tblSite.Close
			Set tblSite = Nothing
	else
		
		Set tblFiles = Server.CreateObject("ADODB.RecordSet")
		sqlFiles = "SELECT * FROM W_Files_t WHERE SSN = '" & Request("WID") & "' "
		tblFiles.Open sqlFiles, g_strCONN, 1, 3
		If Not tblFiles.EOF Then
			If Request("chkt1") <> "" Then 
					tblFiles("t1") = True 
			Else
				tblFiles("t1") = False
			End If
			If Request("chkt2") <> "" Then 
					tblFiles("t2") = True 
			Else
				tblFiles("t2") = False
			End If
			If Request("chkt3") <> "" Then 
					tblFiles("t3") = True 
			Else
				tblFiles("t3") = False
			End If
			If Request("chktb") <> "" Then 
					tblFiles("tb") = True 
			Else
				tblFiles("tb") = False
			End If
			If Request("tbdate") <> "" Then
				If IsDate(Request("tbdate")) Then
					tblFiles("tbdate") = Z_FormatDate(Request("tbdate"))
				Else
					Session("MSG") = Session("MSG") & "<br>" & "Enter a valid date for TB Test 1 Datestamp."
				End If
			Else
				tblFiles("tbdate") = Empty
			End If
			If Request("chktb2") <> "" Then 
					tblFiles("tb2") = True 
			Else
				tblFiles("tb2") = False
			End If
			If Request("tbdate2") <> "" Then
				If IsDate(Request("tbdate2")) Then
					tblFiles("tb2date") = Z_FormatDate(Request("tbdate2"))
				Else
					Session("MSG") = Session("MSG") & "<br>" & "Enter a valid date for TB Test 2 Datestamp."
				End If
			Else
				tblFiles("tb2date") = Empty
			End If
			If Request("chkphy") <> "" Then 
					tblFiles("phy") = True 
			Else
				tblFiles("phy") = False
			End If
			If Request("orientdate") <> "" Then
				If IsDate(Request("orientdate")) Then
					tblFiles("orientdate") = Z_FormatDate(Request("orientdate"))
				Else
					Session("MSG") = Session("MSG") & "<br>" & "Enter a valid date for Orientation Datestamp."
				End If
			Else
				tblFiles("orientdate") = Empty
			End If
			tblFiles("orient") = False
			If Request("chkorient") <> "" Then tblFiles("orient") = True
			tblFiles("pptrain") = False
			If Request("chkpp") <> "" Then tblFiles("pptrain") = True
			tblFiles("lnaactive") = False
			If Request("chklnaa") <> "" Then tblFiles("lnaactive") = True
			tblFiles("lnainactive") = False
			If Request("chklnai") <> "" Then tblFiles("lnainactive") = True
			If Request("essentialsdate") <> "" Then
				If IsDate(Request("essentialsdate")) Then
					tblFiles("essentialsdate") = Z_FormatDate(Request("essentialsdate"))
				Else
					Session("MSG") = Session("MSG") & "<br>" & "Enter a valid date for Essentials Training Datestamp."
				End If
			Else
				tblFiles("essentialsdate") = Empty
			End If
			tblFiles("essentials") = False
			If Request("essentials") <> "" Then tblFiles("essentials") = True
			tblFiles.Update
		End If
		tblFiles.Close
		Set tblFiles = Nothing	
		'train
		Set tblEMP = Server.CreateObject("ADODB.Recordset")
		strSQL = "SELECT * FROM [W_log_T]"
		'On Error Resume Next
		tblEMP.Open strSQL, g_strCONN, 1, 3
		ctr = Request("ctr")
			For i = 0 to ctr - 1 
				tmpctr = Request("chktrain" & i)
				If tmpctr <> "" Then
							strTmp = "Index= " & tmpctr
					tblEMP.Find(strTmp)
					If Not tblEMP.EOF Then
						If Request("txttrain" & i) <> "" Then 
							tblEmp("train") = Request("txttrain" & i)
						Else
							Session("MSG") = "Training value is required. " & tblEmp("train") & " changed to " & Request("txttrain" & i) & ".<br>"
						End If
						If isnumeric(Request("txthrs" & i)) <> "" Then 
							tblEmp("thrs") = Request("txthrs" & i)
						Else
							Session("MSG") = "Invalid hours. " & tblEmp("thrs") & " changed to " & Request("txthrs" & i) & ".<br>"
						End If
						tblEmp("tcom") = Request("trainnote" & i)
						tblEmp.Update
					End If
				End If
				tblEmp.MoveFirst
				
			Next
			 
		tblEmp.Close
		Set tblEMP = Nothing
				
	If Request("txttrain") <> "" And Request("txthrs") <> "" Then
		If Isnumeric(Request("txthrs")) Then
			Set tblSV = Server.CreateObject("ADODB.RecordSet")
			sqlSV = "SELECT * FROM [w_log_t]"
			tblSV.Open sqlSV, g_strCONN, 1, 3
			tblSV.AddNew
			tblSV("ssn") = Request("wID")
			tblSV("train") = Request("txttrain")
			tblSV("thrs") = Request("txthrs")
			tblSV("tcom") = Request("trainnote")
			tblSV.Update
			tblSV.Close
			Set tblSV = Nothing
		Else
			Session("MSG") = "Invalid hours<br>" 
		End If
	End If
end if	
		
		Response.Redirect "A_W_Files.asp?WID=" & Request("WID")
	ElseIf Request("page") = 3 Then
		If Request("SelCon") <> "" Then
			Set rsWork = Server.CreateObject("ADODB.RecordSet")
			sqlWork = "SELECT * FROM [ConWork_t]"
			rsWork.Open sqlWork, g_strCONN, 1, 3
			rsWork.AddNew
			rsWork("CID") = Request("SelCon")
			rsWork("WID") = GetWIndex(Request("WID"))
			rsWork.Update
			rsWork.close
			Set rsWork = Nothing
			End If	
		Response.Redirect "WorkCon.asp?WID=" & Request("WID")
		ElseIf Request("page") = 4 Then
			Set rsSkill = Server.CreateObject("ADODB.RecordSet")
			rsSkill.Open "SELECT * FROM w_skills_T WHERE Wid = '" & Request("WID") & "' ", g_strCONN, 1, 3
			If Not rsSkill.EOF Then
				rsSkill("housekeep") = False
				rsSkill("laundry") = False
				rsSkill("meal") = False
				rsSkill("grocery") = False
				rsSkill("dress") = False
				rsSkill("eat") = False
				rsSkill("asstwalk") = False
				rsSkill("asstwheel") = False
				rsSkill("asstmotor") = False
				rsSkill("commeal") = False
				rsSkill("medical") = False
				rsSkill("shower") = False
				rsSkill("tub") = False
				rsSkill("oral") = False
				rsSkill("commode") = False
				rsSkill("sit") = False
				rsSkill("medication") = False
				rsSkill("undress") = False
				rsSkill("shampoosink") = False
				rsSkill("oralcare") = False
				rsSkill("massage") = False
				rsSkill("shampoobed") = False
				rsSkill("shave") = False
				rsSkill("bedbath") = False
				rsSkill("bedpan") = False
				rsSkill("ptexer") = False
				rsSkill("hoyer") = False
				rsSkill("eye") = False
				rsSkill("transferbelt") = False
				rsSkill("alz") = False
				rsSkill("Incontinence") = False
				rsSkill("Hospice") = False
				rsSkill("lotion") = False
				If Request("housekeep") = 1 Then rsSkill("housekeep")	=	True
				If Request("laundry") = 1 Then rsSkill("laundry")	=	True
				If Request("meal") = 1 Then rsSkill("meal")	=	True
				If Request("grocery") = 1 Then rsSkill("grocery")	=	True
				If Request("dress") = 1 Then rsSkill("dress")	=	True
				If Request("eat") = 1 Then rsSkill("eat")	=	True
				If Request("asstwalk") = 1 Then rsSkill("asstwalk")	=	True
				If Request("asstwheel") = 1 Then rsSkill("asstwheel")	=	True
				If Request("asstmotor") = 1 Then rsSkill("asstmotor")	=	True
				If Request("commeal") = 1 Then rsSkill("commeal")	=	True
				If Request("medical") = 1 Then rsSkill("medical")	=	True
				If Request("shower") = 1 Then rsSkill("shower")	=	True
				If Request("tub") = 1 Then rsSkill("tub")	=	True
				If Request("oral") = 1 Then rsSkill("oral")	=	True
				If Request("commode") = 1 Then rsSkill("commode")	=	True
				If Request("sit") = 1 Then rsSkill("sit")	=	True
				If Request("medication") = 1 Then rsSkill("medication")	=	True
				If Request("undress") = 1 Then rsSkill("undress")	=	True
				If Request("shampoosink") = 1 Then rsSkill("shampoosink")	=	True
				If Request("oralcare") = 1 Then rsSkill("oralcare")	=	True
				If Request("massage") = 1 Then rsSkill("massage")	=	True
				If Request("shampoobed") = 1 Then rsSkill("shampoobed")	=	True
				If Request("shave") = 1 Then rsSkill("shave")	=	True
				If Request("bedbath") = 1 Then rsSkill("bedbath")	=	True
				If Request("bedpan") = 1 Then rsSkill("bedpan")	=	True
				If Request("ptexer") = 1 Then rsSkill("ptexer")	=	True
				If Request("hoyer") = 1 Then rsSkill("hoyer")	=	True
				If Request("eye") = 1 Then rsSkill("eye")	=	True
				If Request("transferbelt") = 1 Then rsSkill("transferbelt")	=	True
				If Request("alz") = 1 Then rsSkill("alz")	=	True
				If Request("Incontinence") = 1 Then rsSkill("Incontinence")	=	True
				If Request("Hospice") = 1 Then rsSkill("Hospice")	=	True
				If Request("lotion") = 1 Then rsSkill("lotion")	=	True
				rsSkill.Update
			End If
			rsSkill.Close
			Set rsSkill = Nothing
			Response.Redirect "a_w_skills.asp?WID=" & Request("WID")
	End If
%>