<%@Language=VBScript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_security.asp" -->
<%
DIM		strtype, tbltype, strSQLt, sbmtDATE, lngI, tblEMP, strSQL, strTableScript, strComHrs
DIM 		tbldept, strComMon, strComTue, strComWed, strComThu, strComFri, strComSat, strComSun
DIM		strSQLd, tmpDATE, myDATE, monDATE, tueDATE, wedDATE, thuDATE, friDATE, satDATE, sunDATE
DIM		strdept, name, Ename, Edate, Eid, difDATE, myDATE2, finDATE, lname, fname, strLINK, mlMail
DIM		strd8, strstat, strdesc, strDisabled, tblstat, strSQLstat, tblLOC, tblJOB, strSQLl, strSQLj
DIM		tbltdes, strSQLtd, tblOthers, strSQLo ,strComputeCode, ctr, email

If SEssion("userid") = "" Then 
	session("MSG") = "Session expired. Please Login again."
	response.redirect "default.asp"
End If
editme = "disabled"
If UCase(Session("lngType")) = "2" or UCase(Session("lngType")) = "1" Then
	editme = ""
End If
function AllowMileage(xxx)
	AllowMileage = False
	Set rsAllow = Server.CreateObject("ADODB.RecordSet")
	sqlAllow = "SELECT * FROM Worker_T WHERE Social_Security_Number = '" & xxx  & "'"
	rsAllow.Open sqlAllow, g_strCONN, 1, 3
	If not rsAllow.EOF Then
		if rsAllow("driver") = true then AllowMileage = true
	end if
	rsAllow.Close
	set rsAllow = Nothing
end function

Set tblWork = Server.CreateObject("ADODB.RecordSet")
sqlWork = "SELECT * FROM Worker_t WHERE Status = 'Active' Order by Lname, Fname ASC"
tblWork.Open sqlWork, g_strCONN, 1, 3
	Do Until tblWork.EOF
		If Session("idemp") = tblWork("Social_Security_Number") Then Sel = "selected"
		flt = ""
		If tblWork("flt") Then flt = " (float)"
		strWork = strWork & "<option " & Sel & " value='" & tblWork("Social_Security_Number") & "' >" & Right(tblWork("Social_Security_Number"),4) & " - " & _
				tblWork("lname") & ", "  & tblWork("fname") & flt & "</option>"
		tblWork.MoveNext
		Sel = ""
	Loop
tblWork.Close
set tblWork = Nothing

tmpDATE = Session("dtDate")
'tmpd8 = DatePart("ww", Cdate(tmpDATE))
'If tmpd8 = 53 then tmpd8 = tmpd8 - 1
difwk = DateDiff("ww", wk1, tmpDATE)

If Not Z_IsOdd2(difwk) Then
	myDATE = tmpDATE
Else
	myDATE = DateAdd("d", -7, tmpDATE)
End If

''''FIRST WEEK
If WeekdayName(Weekday(myDATE), true) = "Sun" Then
	finDATE = myDATE
	sunDATE = myDATE
	monDATE = DateAdd("d", 1, sunDATE)
	tueDATE = DateAdd("d", 1, monDATE)
	wedDATE = DateAdd("d", 1, tueDATE)
	thuDATE = DateAdd("d", 1, wedDATE)
	friDATE = DateAdd("d", 1, thuDATE)
	satDATE = DateAdd("d", 1, friDATE)
Else
	difDATE = DatePart("w", myDATE)
	sunDATE = DateAdd("d", -Cint(difDATE - 1), myDATE)
	myDATE2 = sunDATE
	finDATE = myDATE2
	monDATE = DateAdd("d", 1, sunDATE)
	tueDATE = DateAdd("d", 1, monDATE)
	wedDATE = DateAdd("d", 1, tueDATE)
	thuDATE = DateAdd("d", 1, wedDATE)
	friDATE = DateAdd("d", 1, thuDATE)
	satDATE = DateAdd("d", 1, friDATE)
End If
Session("sundate") = sunDATE
'''''2ND WEEK
If WeekdayName(Weekday(myDATE), true) = "Sun" Then
	finDATE2 = myDATE
	sunDATE2 = DateAdd("d", 7, myDATE)
	monDATE2 = DateAdd("d", 1, sunDATE2)
	tueDATE2 = DateAdd("d", 1, monDATE2)
	wedDATE2 = DateAdd("d", 1, tueDATE2)
	thuDATE2 = DateAdd("d", 1, wedDATE2)
	friDATE2 = DateAdd("d", 1, thuDATE2)
	satDATE2 = DateAdd("d", 1, friDATE2)
Else
	difDATE2 = DatePart("w", myDATE)
	tmpDates = DateAdd("d", -Cint(difDATE - 1), myDATE)
	sunDATE2 = DateAdd("d", 7, tmpDates)
	myDATE2s = sunDATE2
	finDATE2 = myDATE2s
	monDATE2 = DateAdd("d", 1, sunDATE2)
	tueDATE2 = DateAdd("d", 1, monDATE2)
	wedDATE2 = DateAdd("d", 1, tueDATE2)
	thuDATE2 = DateAdd("d", 1, wedDATE2)
	friDATE2 = DateAdd("d", 1, thuDATE2)
	satDATE2 = DateAdd("d", 1, friDATE2)
End If
'response.write "1:" & sundate & "<br> 2:" & sunDATE2
	Session("d8") = finDATE

'disable mileage
mileageAllow = "disabled"
if AllowMileage(Session("idemp")) Then mileageAllow = ""

''''''''''''''1ST WEEK
Set tblEMP = Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT * FROM [tsheets_t],[Worker_t] WHERE tsheets_t.emp_id = Worker_t.Social_security_Number AND " & _
		"[emp_id] = '" & Session("idemp") & "' AND " & _
		"[date] = '" & finDATE &  "' AND EXT = 0 ORDER BY client, timestamp"
tblEMP.Open strSQL, g_strCONN, 3, 1

Session("name") = Session("namef") & " " & Session("namel")
if session("name") = "" then
	session("MSG") = "Operation Timed Out."
	response.redirect "default.asp"
end if
'email = tblEMP("email")
Eid = Session("idemp")
Ename = name

lngI = 0
ctr = 0
Do While Not tblEMP.EOF
	'If tblEMP("verify") Then 
	'	strDisabled = " disabled"
	'	ctr = ctr + 1 
	'Else 
	'	strDisabled = ""
	'End if
	
	if Z_IsOdd(lngI) = true then 
		kulay = "#FFFAF0" 
	else 
		kulay = "#FFFFFF"
	end if
	''''''GET NAME OF CONSUMER''''''
	Set tblCon = Server.CreateObject("ADODB.RecordSet")
	sqlCon = "SELECT * FROM Consumer_t WHERE Medicaid_Number ='" & tblEMP("Client") & "' "
	tblCon.Open sqlCon, g_strCONN, 1, 3
		If Not tblCon.EOF Then
			If tblcon("code") = "P" or tblcon("code") = "C" then mycode = "*"
			pangalan = mycode & tblCon("Lname") & ", " & tblCon("Fname")
		Else
			pangalan = "* " & tblEMP("Client")
		End If
	tblCon.Close
	Set tblCon = Nothing
	''''''''''''''''''''''''''''' GET EXT. HRS
	'Set rsEXT = Server.CreateObject("ADODB.RecordSet")
	EXT = ""
	xmon = ""
	xtue = ""
	xwed = ""
	xthu = ""
	xfri = ""
	xsat = ""
	xsun = ""
	xmonv = 0
	xtuev = 0
	xwedv = 0
	xthuv = 0
	xfriv = 0
	xsatv = 0
	xsunv = 0
	XstrComMon = ""
	XstrComTue = ""
	XstrComWed = ""
	XstrComThu = ""
	XstrComFri = ""
	XstrComSat= ""
	XstrComSun = ""
	myIndex = tblEMP("ID")
	Set rsEXT = Server.CreateObject("ADODB.RecordSet")
	'sqlEXT = "SELECT * " & _
	'	"FROM [tsheets_t],[Worker_t] " & _
	'	"WHERE [tsheets_t.emp_id] = [Worker_t.Social_security_Number] " & _
	'	"AND [emp_id] = '" & Session("idemp") & "' " & _
	'	"AND [date] = #" & finDATE & "# " & _
	'	"AND [EXT] = True AND [client] = '" & tblEMP("Client") & "' "
	'If tblEMP("timestamp") <> Empty Then sqlEXT = sqlEXT & "AND [timestamp] = #" & tblEMP("timestamp") & "# "
	'sqlEXT = sqlEXT & "ORDER BY [client], [timestamp]"
	sqlEXT = "SELECT * " & _
		"FROM [tsheets_t],[Worker_t] " & _
		"WHERE [emp_id] = [Social_security_Number] " & _
		"AND [emp_id] = '" & Session("idemp") & "' " & _
		"AND [date] = '" & finDATE & "' " & _
		"AND [EXT] = 1 AND [client] = '" & tblEMP("Client") & "' AND ID = " & myIndex + 1
	'response.write sqlext
	'rsEXT.Open sqlEXT, g_strCONN, 1, 3 
	rsEXT.OPEN sqlEXT, g_strCONN, 1, 3
	If Not rsEXT.EOF Then
		strDisabled = ""
		If rsEXT("ProcPay") <> "" OR rsEXT("ProcMed") <> "" OR rsEXT("ProcPriv") <> "" Then strDisabled = "readOnly"
		XmonV = rsEXT("mon")
		XtueV = rsEXT("tue")
		XwedV = rsEXT("wed")
		XthuV = rsEXT("thu")
		XfriV = rsEXT("fri")
		XsatV = rsEXT("sat")
		XsunV = rsEXT("sun")
		XstrComputeCode = ""
		XhrsV = Z_CDbl(XmonV) + Z_CDbl(XtueV) + Z_CDbl(XwedV) + Z_CDbl(XthuV) + Z_CDbl(XfriV) + Z_CDbl(XsatV) + Z_CDbl(XsunV)
		If XhrsV <> 0 Then EXT = "checked"
		Xmon = "<br><input style='font-size: 8pt;' type='text' size= '3' class='rightjust' onblur='numericVal(this.value, this); ComputeRows(); Compute(this, 1);' name='hmonX1" & _
			lngI & "' value='" & XmonV & "' class='textbox' " & vbCrLf & _
			strDisabled & ">"
		Xtue = "<br><input style='font-size: 8pt;' type='text' size= '3' class='rightjust' onblur='numericVal(this.value, this); ComputeRows(); Compute(this, 1);' name='htueX1" & _
			lngI & "' value='" & XtueV & "' class='textbox' " & vbCrLf & _
			strDisabled & ">"
		Xwed = "<br><input style='font-size: 8pt;' type='text' size= '3' class='rightjust' onblur='numericVal(this.value, this); ComputeRows(); Compute(this, 1);' name='hwedX1" & _
			lngI & "' value='" & XwedV & "' class='textbox' " & vbCrLf & _
			strDisabled & ">"
		Xthu = "<br><input style='font-size: 8pt;' type='text' size= '3' class='rightjust' onblur='numericVal(this.value, this); ComputeRows(); Compute(this, 1);' name='hthuX1" & _
			lngI & "' value='" & XthuV & "' class='textbox' " & vbCrLf & _
			strDisabled & ">"
		Xfri = "<br><input style='font-size: 8pt;' type='text' size= '3' class='rightjust' onblur='numericVal(this.value, this); ComputeRows(); Compute(this, 1);' name='hfriX1" & _
			lngI & "' value='" & XfriV & "' class='textbox' " & vbCrLf & _
			strDisabled & ">"
		Xsat = "<br><input style='font-size: 8pt;' type='text' size= '3' class='rightjust' onblur='numericVal(this.value, this); ComputeRows(); Compute(this, 1);' name='hsatX1" & _
			lngI & "' value='" & XsatV & "' class='textbox' " & vbCrLf & _
			strDisabled & ">"
		Xsun = "<br><input style='font-size: 8pt;' type='text' size= '3' class='rightjust' onblur='numericVal(this.value, this); ComputeRows(); Compute(this, 1);' name='hsunX1" & _
			lngI & "' value='" & XsunV  & "' class='textbox' " & vbCrLf & _
			strDisabled & ">"
		
		XstrComputeCode = "+ parseFloat(document.frmTime.hmonX1" & lngI & ".value) + " & _
					"parseFloat(document.frmTime.htueX1" & lngI & ".value) + " & _
					"parseFloat(document.frmTime.hwedX1" & lngI & ".value) + " & _
					"parseFloat(document.frmTime.hthuX1" & lngI & ".value) + " & _
					"parseFloat(document.frmTime.hfriX1" & lngI & ".value) + " & _
					"parseFloat(document.frmTime.hsatX1" & lngI & ".value) + " & _
					"parseFloat(document.frmTime.hsunX1" & lngI & ".value);" & vbCrLf
					
		XstrComMon = "parseFloat(document.frmTime.hmonX1" & lngI & ".value) +"
		XstrComTue = "parseFloat(document.frmTime.htueX1" & lngI & ".value) +"
		XstrComWed = "parseFloat(document.frmTime.hwedX1" & lngI & ".value) +"
		XstrComThu = "parseFloat(document.frmTime.hthuX1" & lngI & ".value) +"
		XstrComFri= "parseFloat(document.frmTime.hfriX1" & lngI & ".value) +"
		XstrComSat = "parseFloat(document.frmTime.hsatX1" & lngI & ".value) +"
		XstrComSun = "parseFloat(document.frmTime.hsunX1" & lngI & ".value) +"
		
		strEXT = strEXT & "if (document.frmTime.chkEXT1" & lngI & ".checked == true) " & _
			"{document.frmTime.hsunX1" & lngI & ".style.visibility = 'visible'; " & _
			"document.frmTime.hmonX1" & lngI & ".style.visibility = 'visible'; " & _
			"document.frmTime.htueX1" & lngI & ".style.visibility = 'visible'; " & _
			"document.frmTime.hwedX1" & lngI & ".style.visibility = 'visible'; " & _
			 "document.frmTime.hthuX1" & lngI & ".style.visibility = 'visible'; " & _
			 "document.frmTime.hfriX1" & lngI & ".style.visibility = 'visible'; " & _
			 "document.frmTime.hsatX1" & lngI & ".style.visibility = 'visible';} " & _
		"else " &_
			"{document.frmTime.hsunX1" & lngI & ".style.visibility = 'hidden'; " & _
			 "document.frmTime.hsunX1" & lngI & ".value = 0; " & _
			 "document.frmTime.hmonX1" & lngI & ".style.visibility = 'hidden'; " & _
			 "document.frmTime.hmonX1" & lngI & ".value = 0; " & _
			 "document.frmTime.htueX1" & lngI & ".style.visibility = 'hidden'; " & _
			 "document.frmTime.htueX1" & lngI & ".value = 0; " & _
			 "document.frmTime.hwedX1" & lngI & ".style.visibility = 'hidden'; " & _
			 "document.frmTime.hwedX1" & lngI & ".value = 0; " & _
			 "document.frmTime.hthuX1" & lngI & ".style.visibility = 'hidden'; " & _
			 "document.frmTime.hthuX1" & lngI & ".value = 0; " & _
			 "document.frmTime.hfriX1" & lngI & ".style.visibility = 'hidden'; " & _
			 "document.frmTime.hfriX1" & lngI & ".value = 0; " & _
			 "document.frmTime.hsatX1" & lngI & ".value = 0; " & _
			 "document.frmTime.hsatX1" & lngI & ".style.visibility = 'hidden';} " & vbCrLf
		
	End If	
	rsEXT.Close
	Set rsEXT = Nothing
	strDisabled = ""
	If tblEMP("ProcPay") <> "" OR tblEMP("ProcMed") <> "" OR tblEMP("ProcPriv") <> "" Then strDisabled = "readOnly"
	noedit = ""
	'If tblEMP("ProcPay") <> "" OR tblEMP("ProcMed") <> "" OR tblEMP("ProcPriv") <> "" Then noedit = "disabled"	
	strTableScript = strTableScript & "<tr bgcolor='" & kulay & "'><td bgcolor='#d4d0c8' align='center' rowspan='2'>" & vbCrLf & _
			"<input style='font-size: 8pt;' type='checkbox' size= '5' name='chk" & lngI & "' value='" & _
			tblEMP("id") & "'" & noedit & " onclick='RowDis();'>" & _
			"<input type='hidden' name='tmpID" & lngI & "' value='" & tblEMP("id") & "'></td>" & _
			"<td align='center' rowspan='2'><input style='font-size: 8pt;' type='text' readOnly name='hdept1" & lngI & "' size= '12' value=""" & pangalan & """ " & strDisabled & "" & _
			"><input type='hidden' name='conid1" & lngI & "' value='" & tblEMP("Client") & "'><br>" & _
			"<font face='trebuchet MS' size='1'>Extd Hrs.</font><input type='checkbox' size= '5' " & EXT & "  value='" & tblEMP("id") & "' name='chkEXT1" & lngI & "' onclick='GetEXT();ComputeRows();Compute();'>" & _
			"</td>" & vbCrLf

	strTableScript = strTableScript & "<td align='center' rowspan='2'>" & _
			"<input style='font-size: 8pt;' type='text' size= '3' class='rightjust' onblur='numericVal(this.value, this); ComputeRows(); Compute(this, 1);' name='hsun1" & _
			lngI & "' value='" & tblEMP("sun") & "' class='textbox' " & vbCrLf & _
			strDisabled & ">" & xsun & "</td>" & vbCrLf & _
			"<td align='center' rowspan='2'><input style='font-size: 8pt;' type='text' size='3' name='hmon1" & _
			lngI & "' value='" & tblEMP("mon") & "' onblur='numericVal(this.value, this); ComputeRows(); Compute(this, 1);' class='rightjust'  " & _
			strDisabled & ">" & xmon & "</td" & vbCrLf & _
			"><td align='center' rowspan='2'><input style='font-size: 8pt;' type='text' onblur='numericVal(this.value, this); ComputeRows(); Compute(this, 1);' class='rightjust' size= '3' name='htue1" & _
			lngI & "' class='textbox' value='" & tblEMP("tue") & "' " & _
			strDisabled & ">" & xtue & "</td" & vbCrLf & _
			"><td align='center' rowspan='2'><input style='font-size: 8pt;' type='text' class='rightjust' size= '3' onblur='numericVal(this.value, this); ComputeRows(); Compute(this, 1);' name='hwed1" & _
			lngI & "' class='textbox' value='" & tblEMP("wed") & "' " & _
			strDisabled & ">" & xwed & "</td" & vbCrlf & _
			"><td align='center' rowspan='2'><input style='font-size: 8pt;' type='text' class='rightjust' size= '3' onblur='numericVal(this.value, this); ComputeRows(); Compute(this, 1);' name='hthu1" & _
			lngI & "' class='textbox' value='" & tblEMP("thu") &"' " & _
			strDisabled & ">" & xthu & "</td" & vbCrlf & _
			"><td align='center' rowspan='2'><input style='font-size: 8pt;' type='text' size= '3' class='rightjust' onblur='numericVal(this.value, this); ComputeRows(); Compute(this, 1);' name='hfri1" & _
			lngI & "' class='textbox' value='" & tblEMP ("fri") & "' " & _
			strDisabled & ">" & xfri & "</td" & vbCrlf & _
			"><td align='center' rowspan='2'><input style='font-size: 8pt;' type='text' size= '3' class='rightjust' onblur='numericVal(this.value, this); ComputeRows(); Compute(this, 1);' name='hsat1" & _
			lngI & "' value='" & tblEMP("sat") & "' class='textbox' " & _
			strDisabled & ">" & xsat & "</td" 
			
	
			strTableScript = strTableScript & "></td><td align='center' rowspan='2'><input style='font-size: 8pt;' name='htot1" & lngI & "' " & _
					" readOnly disabled class='rightjust'  size='3' value='" & ( Z_CDbl(tblEmp("mon")) + Z_CDbl(tblEmp("tue")) + _
						Z_CDbl(tblEmp("wed")) + Z_CDbl(tblEmp("thu")) + _
						Z_CDbl(tblEmp("fri")) + Z_CDbl(tblEmp("sat")) + _
						Z_CDbl(tblEmp("sun")) + _
						Z_CDbl(XsunV) + Z_CDbl(XmonV) + Z_CDbl(XtueV) + Z_CDbl(XwedV) + Z_CDbl(XthuV) + Z_CDbl(XfriV) + Z_CDbl(XsatV)) & "' " & strDisabled & "></td" & vbCrLf
							
			'if Ucase(strDisabled) = "DISABLED" Then mileageAllow = ""
				
			procmile = ""	
			myMile = ""
			if z_fixnull(tblEMP("ProcMile")) <> "" then 
				procmile = "Readonly"
				myMile = "*"
			End If	
			strTableScript = strTableScript & "><td align='center' rowspan='2'><input style='font-size: 8pt;' type='text' size= '3' class='rightjust' onblur='numericVal(this.value, this); ComputeRows(); Compute(this, 1);' name='txtmile" & _
			lngI & "' value='" & tblEMP("mile") & "' class='textbox' " & " " & mileageAllow & " " & procmile
			
			strTableScript = strTableScript & ">" & myMile & "</td><td align='center' rowspan='2'><input style='font-size: 8pt;' type='text' size= '3' class='rightjust' onblur='numericVal(this.value, this); ComputeRows(); Compute(this, 1);' name='txtamile" & _
			lngI & "' value='" & tblEMP("amile") & "' class='textbox' " & " " & procmile
			
			strComputeCode = strComputeCode & "document.frmTime.htot1" & lngI & _
					".value = parseFloat(document.frmTime.hmon1" & lngI & ".value) + " & _
					"parseFloat(document.frmTime.htue1" & lngI & ".value) + " & _
					"parseFloat(document.frmTime.hwed1" & lngI & ".value) + " & _
					"parseFloat(document.frmTime.hthu1" & lngI & ".value) + " & _
					"parseFloat(document.frmTime.hfri1" & lngI & ".value) + " & _
					"parseFloat(document.frmTime.hsat1" & lngI & ".value) + " & _
					"parseFloat(document.frmTime.hsun1" & lngI & ".value) " & XstrComputeCode & ";" & vbCrLf
	
	strTableScript = strTableScript & "></td><td align='center' rowspan='2'><textarea style='font-size: 8pt;' readonly rows='2' name='Mnotes1" & _
			lngI & "' cols='15'>" & tblEMP("misc_notes") & _
			"</textarea><td align='center' rowspan='2'><textarea style='font-size: 8pt;' readonly rows='2' name='com1" & _
			lngI & "' cols='15'>" & tblEMP("actcode") & _
			"</textarea"  & vbCrLf 
	
	PayChk2 = ""
	medchk2 = ""
	'If tblEMP("ProcPay") <> "" Then PayChk2 = "checked"
	If tblEMP("ProcMed") <> "" Then MedChk2 = "checked"
	If tblEMP("ProcPriv") <> "" Then MedChk2 = "checked"
	If tblEMP("ProcVA") <> "" Then MedChk2 = "checked"
		
	strTableScript = strTableScript & " ></td><td bgcolor='#d4d0c8' align='center' rowspan='2'>" & vbCrLf & _
			"<input style='font-size: 8pt;' type='text' readOnly name='fon1" & lngI & "' size= '12' value=""" & tblEMP("CallerID") & """ " & strDisabled & "" & _
			"></td><td bgcolor='#d4d0c8' align='center' rowspan='2'>" & vbCrLf & _
			"<input style='font-size: 8pt;' type='checkbox' size= '5' name='chkM1" & lngI & "' value='" & _
			tblEMP("ProcMed") & "' " & MedChk2 & " disabled ></td>" & vbCrLf 
			
	If (Session("lngType")) = 2 AND (Paychk2 = "checked" OR Medchk2 = "checked") Then
		strTableScript = strTableScript	& "<td valign='middle' bgcolor='#d4d0c8' rowspan='2'><input type='button' value='<< Untag' " & _
		"onclick='Untag(" & tblEMP("id") & ");' class='btn' onmouseover=""this.className='hovbtn'"" onmouseout=""this.className='btn'""></td>" & vbCrLf 
	End If		
	
	strTableScript = strTableScript	& "<tr><td>&nbsp;</td></tr>" & vbCrLf 
	
	
	
	strComMon = strComMon & "parseFloat(document.frmTime.hmon1" & lngI & ".value) + " & XstrComMon
	strComTue = strComTue & "parseFloat(document.frmTime.htue1" & lngI & ".value) + " & XstrComTue
	strComWed = strComWed & "parseFloat(document.frmTime.hwed1" & lngI & ".value) + " & XstrComWed
	strComThu = strComThu & "parseFloat(document.frmTime.hthu1" & lngI & ".value) + " & XstrComThu
	strComFri = strComFri & "parseFloat(document.frmTime.hfri1" & lngI & ".value) + " & XstrComFri
	strComSat = strComSat & "parseFloat(document.frmTime.hsat1" & lngI & ".value) + " & XstrComSat
	strComSun = strComSun & "parseFloat(document.frmTime.hsun1" & lngI & ".value) + " & XstrComSun
	strComMile = strComMile & "parseFloat(document.frmTime.txtmile" & lngI & ".value) + "
	strComaMile = strComaMile & "parseFloat(document.frmTime.txtamile" & lngI & ".value) + "
	
	If procmile <> "Readonly" Then
		If tblEMP("ProcPay") <> "" OR tblEMP("ProcMed") <> "" OR tblEMP("ProcPriv") <> "" Then 
			strDis = strDis & "if (document.frmTime.chk" & lngI & ".checked == true) " & _
					"{document.frmTime.chkEXT1" & lngI & ".disabled = true; " & vbCrlf & _
					"document.frmTime.hsun1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.hmon1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.htue1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.hwed1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.hthu1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.hfri1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.hsat1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.hsunX1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.hmonX1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.htueX1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.hwedX1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.hthuX1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.hfriX1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.hsatX1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.hsatX1" & lngI & ".readOnly  = true; " & vbCrlf
					if AllowMileage(Session("idemp")) Then 
						strDis = strDis & "document.frmTime.txtmile" & lngI & ".readOnly  = false; "
					else
						strDis = strDis & "document.frmTime.txtmile" & lngI & ".readOnly  = true; "
					end if
					strDis = strDis & "document.frmTime.txtamile" & lngI & ".readOnly  = false; " & vbCrlf &_
					"document.frmTime.Mnotes1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.fon1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.com1" & lngI & ".readOnly  = true;} " & vbCrlf &_
				"else " & vbCrLf & _
					"{document.frmTime.chkEXT1" & lngI & ".disabled = true; " & vbCrlf &_
					"document.frmTime.hsun1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.hmon1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.htue1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.hwed1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.hthu1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.hfri1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.hsat1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.hsunX1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.hmonX1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.htueX1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.hwedX1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.hthuX1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.hfriX1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.hsatX1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.txtmile" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.txtamile" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.Mnotes1" & lngI & ".readOnly  = true; " & vbCRlf &_
					"document.frmTime.fon1" & lngI & ".readOnly  = true; " & vbCRlf &_
					"document.frmTime.com1" & lngI & ".readOnly  = true;} " & vbCrlf
			Else
				strDis = strDis & "if (document.frmTime.chk" & lngI & ".checked == true) " & _
					"{document.frmTime.chkEXT1" & lngI & ".disabled = false; " & vbCrlf & _
					"document.frmTime.hsun1" & lngI & ".readOnly  = false; " & vbCrlf &_
					"document.frmTime.hmon1" & lngI & ".readOnly  = false; " & vbCrlf &_
					"document.frmTime.htue1" & lngI & ".readOnly  = false; " & vbCrlf &_
					"document.frmTime.hwed1" & lngI & ".readOnly  = false; " & vbCrlf &_
					"document.frmTime.hthu1" & lngI & ".readOnly  = false; " & vbCrlf &_
					"document.frmTime.hfri1" & lngI & ".readOnly  = false; " & vbCrlf &_
					"document.frmTime.hsat1" & lngI & ".readOnly  = false; " & vbCrlf &_
					"document.frmTime.hsunX1" & lngI & ".readOnly  = false; " & vbCrlf &_
					"document.frmTime.hmonX1" & lngI & ".readOnly  = false; " & vbCrlf &_
					"document.frmTime.htueX1" & lngI & ".readOnly  = false; " & vbCrlf &_
					"document.frmTime.hwedX1" & lngI & ".readOnly  = false; " & vbCrlf &_
					"document.frmTime.hthuX1" & lngI & ".readOnly  = false; " & vbCrlf &_
					"document.frmTime.hfriX1" & lngI & ".readOnly  = false; " & vbCrlf &_
					"document.frmTime.hsatX1" & lngI & ".readOnly  = false; " & vbCrlf &_
					"document.frmTime.hsatX1" & lngI & ".readOnly  = false; " & vbCrlf
					if AllowMileage(Session("idemp")) Then 
						strDis = strDis & "document.frmTime.txtmile" & lngI & ".readOnly  = false; "
					else
						strDis = strDis & "document.frmTime.txtmile" & lngI & ".readOnly  = true; "
					end if
					strDis = strDis & "document.frmTime.txtamile" & lngI & ".readOnly  = false; " & vbCrlf &_
					"document.frmTime.Mnotes1" & lngI & ".readOnly  = false; " & vbCrlf &_
						"document.frmTime.fon1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.com1" & lngI & ".readOnly  = false;} " & vbCrlf &_
				"else " & vbCrLf & _
					"{document.frmTime.chkEXT1" & lngI & ".disabled = true; " & vbCrlf &_
					"document.frmTime.hsun1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.hmon1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.htue1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.hwed1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.hthu1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.hfri1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.hsat1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.hsunX1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.hmonX1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.htueX1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.hwedX1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.hthuX1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.hfriX1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.hsatX1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.txtmile" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.txtamile" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.Mnotes1" & lngI & ".readOnly  = true; " & vbCRlf &_
					"document.frmTime.fon1" & lngI & ".readOnly  = true; " & vbCRlf &_
					"document.frmTime.com1" & lngI & ".readOnly  = true;} " & vbCRlf
			End If
	Else
		If tblEMP("ProcPay") <> "" OR tblEMP("ProcMed") <> "" OR tblEMP("ProcPriv") <> "" Then 
			strDis = strDis & "if (document.frmTime.chk" & lngI & ".checked == true) " & vbCrlf &_
					"{document.frmTime.chkEXT1" & lngI & ".disabled = true; " & vbCrlf &_
					"document.frmTime.hsun1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.hmon1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.htue1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.hwed1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.hthu1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.hfri1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.hsat1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.hsunX1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.hmonX1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.htueX1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.hwedX1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.hthuX1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.hfriX1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.txtmile" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.txtamile" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.hsatX1" & lngI & ".readOnly  = true; " & vbCrlf
					if AllowMileage(Session("idemp")) Then 
						strDis = strDis & "document.frmTime.txtmile" & lngI & ".readOnly  = true; "
					else
						strDis = strDis & "document.frmTime.txtmile" & lngI & ".readOnly  = true; "
					end if
					strDis = strDis & "document.frmTime.txtamile" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.Mnotes1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.fon1" & lngI & ".readOnly  = true; " & vbCrlf &_
						"document.frmTime.com1" & lngI & ".readOnly  = true;} " & vbCrlf &_
				"else " & vbCrLf & _
					"{document.frmTime.chkEXT1" & lngI & ".disabled = true; " & vbCrlf &_
					"document.frmTime.hsun1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.hmon1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.htue1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.hwed1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.hthu1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.hfri1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.hsat1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.hsunX1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.hmonX1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.htueX1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.hwedX1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.hthuX1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.hfriX1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.hsatX1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.Mnotes1" & lngI & ".readOnly  = true; " & vbCRlf &_
						"document.frmTime.fon1" & lngI & ".readOnly  = true; " & vbCRlf &_
					"document.frmTime.com1" & lngI & ".readOnly  = true;} " & vbCrlf
		Else
			strDis = strDis & "if (document.frmTime.chk" & lngI & ".checked == true) " & vbCrlf &_
					"{document.frmTime.chkEXT1" & lngI & ".disabled = true; " & vbCrlf &_
					"document.frmTime.hsun1" & lngI & ".readOnly  = false; " & vbCrlf &_
					"document.frmTime.hmon1" & lngI & ".readOnly  = false; " & vbCrlf &_
					"document.frmTime.htue1" & lngI & ".readOnly  = false; " & vbCrlf &_
					"document.frmTime.hwed1" & lngI & ".readOnly  = false; " & vbCrlf &_
					"document.frmTime.hthu1" & lngI & ".readOnly  = false; " & vbCrlf &_
					"document.frmTime.hfri1" & lngI & ".readOnly  = false; " & vbCrlf &_
					"document.frmTime.hsat1" & lngI & ".readOnly  = false; " & vbCrlf &_
					"document.frmTime.hsunX1" & lngI & ".readOnly  = false; " & vbCrlf &_
					"document.frmTime.hmonX1" & lngI & ".readOnly  = false; " & vbCrlf &_
					"document.frmTime.htueX1" & lngI & ".readOnly  = false; " & vbCrlf &_
					"document.frmTime.hwedX1" & lngI & ".readOnly  = false; " & vbCrlf &_
					"document.frmTime.hthuX1" & lngI & ".readOnly  = false; " & vbCrlf &_
					"document.frmTime.hfriX1" & lngI & ".readOnly  = false; " & vbCrlf &_
					"document.frmTime.hsatX1" & lngI & ".readOnly  = false; " & vbCrlf &_
					"document.frmTime.hsatX1" & lngI & ".readOnly  = false; " & vbCrlf
					if AllowMileage(Session("idemp")) Then 
						strDis = strDis & "document.frmTime.txtmile" & lngI & ".readOnly  = true; "
					else
						strDis = strDis & "document.frmTime.txtmile" & lngI & ".readOnly  = true; "
					end if
					strDis = strDis & "document.frmTime.txtamile" & lngI & ".readOnly  = false; " & vbCrlf &_
					"document.frmTime.Mnotes1" & lngI & ".readOnly  = true; " & vbCrlf &_
						"document.frmTime.fon1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.com1" & lngI & ".readOnly  = true;} " & vbCrlf &_
				"else " & vbCrLf & _
					"{document.frmTime.chkEXT1" & lngI & ".disabled = true; " & vbCrlf &_
					"document.frmTime.hsun1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.hmon1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.htue1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.hwed1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.hthu1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.hfri1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.hsat1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.hsunX1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.hmonX1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.htueX1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.hwedX1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.hthuX1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.hfriX1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.hsatX1" & lngI & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.Mnotes1" & lngI & ".readOnly  = true; " & vbCRlf &_
						"document.frmTime.fon1" & lngI & ".readOnly  = true; " & vbCRlf &_
					"document.frmTime.com1" & lngI & ".readOnly  = true;} " & vbCrlf
		End If
	End If
	
	lngI = lngI + 1
	tblEMP.MoveNext
Loop

tblEmp.Close
Set tblEMP = Nothing
'PTO
tmpPTO1 = 0
Set rsPTO = Server.CreateObject("ADODB.RecordSet")
sqlPTO = "SELECT * FROM W_PTO_T WHERE WorkerID = '" & Session("idemp") & "' AND date = '" & sunDATE & "'"
rsPTO.Open sqlPTO, g_strCONN, 1, 3
If Not rsPTO.EOF Then
	tmpPTO1 = rsPTO("PTO")
	PTO1 = ""
	If Not IsNull(rsPTO("procitem")) Then PTO1 = "readOnly"
End If
rsPTO.Close
Set rsPTO = Nothing
''''''''''''''''''2ND WEEK
Set tblEMP = Server.CreateObject("ADODB.Recordset")
strSQL2 = "SELECT * FROM [tsheets_t],[Worker_t] WHERE tsheets_t.emp_id = Worker_t.Social_security_Number AND " & _
		"[emp_id] = '" & Session("idemp") & "' AND " & _
		"[date] = '" & sunDATE2 &  "' AND EXT = 0 ORDER BY client, timestamp"
'on error resume next
tblEMP.Open strSQL2, g_strCONN, 3, 1
'response.write "SQL1 =" & strsql & "<BR> SQL2 = " & strsql2
Session("name") = Session("namef") & " " & Session("namel")
if session("name") = "" then
	session("MSG") = "Operation Timed Out."
	response.redirect "default.asp"
end if
'email = tblEMP("email")
Eid = Session("idemp")
Ename = name

lngI2 = 0
ctr = 0
Do While Not tblEMP.EOF
	'If tblEMP("verify") Then 
	'	strDisabled = " disabled"
	'	ctr = ctr + 1 
	'Else 
	'	strDisabled = ""
	'End if
	
	if Z_IsOdd(lngI2) = true then 
		kulay = "#FFFAF0" 
	else 
		kulay = "#FFFFFF"
	end if
	''''''GET NAME OF CONSUMER''''''
	Set tblCon = Server.CreateObject("ADODB.RecordSet")
	sqlCon = "SELECT * FROM Consumer_t WHERE Medicaid_Number ='" & tblEMP("Client") & "' "
	tblCon.Open sqlCon, g_strCONN, 1, 3
		If Not tblCon.EOF Then
			mycode=""
			If tblcon("code") = "P" or tblcon("code") = "C" then mycode = "*"
			pangalan = mycode & tblCon("Lname") & ", " & tblCon("Fname")
		Else
			pangalan = "**" & tblEMP("Client")
		End If
	tblCon.Close
	Set tblCon = Nothing
''''''''''''''''''''''''''''' GET EXT. HRS2
myIndex = tblEMP("ID")
	Set rsEXT2 = Server.CreateObject("ADODB.RecordSet")
	'sqlEXT2 = "SELECT * FROM [tsheets_t],[Worker_t] WHERE tsheets_t.emp_id = Worker_t.Social_security_Number AND " & _
	'	"[emp_id] = '" & Session("idemp") & "' AND " & _
	'	"[date] = #" & sunDATE2 &  "# AND EXT = True AND client = '" & tblEMP("Client") & "' "
	'If tblEMP("timestamp") <> Empty Then sqlEXT2 = sqlEXT2 & "AND timestamp = #" & tblEMP("timestamp") & "# "
	'sqlEXT2 = sqlEXT2 & "ORDER BY client, timestamp"
	Set rsEXT2 = Server.CreateObject("ADODB.RecordSet")
	sqlEXT2 = "SELECT * FROM [tsheets_t],[Worker_t] WHERE tsheets_t.emp_id = Worker_t.Social_security_Number AND " & _
		"[emp_id] = '" & Session("idemp") & "' AND " & _
		"[date] = '" & sunDATE2 &  "' AND EXT = 1 AND client = '" & tblEMP("Client") & "' AND ID = " & myIndex + 1
	'response.write sqlext2 & "<br>"
	rsEXT2.Open sqlEXT2, g_strCONN, 3, 1
	EXT2 = ""
	xmon2 = ""
	xtue2 = ""
	xwed2= ""
	xthu2 = ""
	xfri2 = ""
	xsat2 = ""
	xsun2 = ""
	xmonv2 = 0
	xtuev2 = 0
	xwedv2 = 0
	xthuv2 = 0
	xfriv2 = 0
	xsatv2 = 0
	xsunv2 = 0
	XstrComMon2 = ""
	XstrComTue2 = ""
	XstrComWed2 = ""
	XstrComThu2 = ""
	XstrComFri2 = ""
	XstrComSat2 = ""
	XstrComSun2 = ""
	If Not rsEXT2.EOF Then
		'response.write "pasok!"
		strDisabled = ""
		If rsEXT2("ProcPay") <> "" OR rsEXT2("ProcMed") <> "" OR rsEXT2("ProcPriv") <> "" Then strDisabled = "readOnly"
		XmonV2 = rsEXT2("mon")
		XtueV2 = rsEXT2("tue")
		XwedV2 = rsEXT2("wed")
		XthuV2 = rsEXT2("thu")
		XfriV2 = rsEXT2("fri")
		XsatV2 = rsEXT2("sat")
		XsunV2 = rsEXT2("sun")
		XstrComputeCode2 = ""
		'response.write xmonv2 
		XhrsV2 = Z_CDbl(XmonV2) + Z_CDbl(XtueV2) + Z_CDbl(XwedV2) + Z_CDbl(XthuV2) + Z_CDbl(XfriV2) + Z_CDbl(XsatV2) + Z_CDbl(XsunV2)
		If XhrsV2 <> 0 Then EXT2 = "checked"
			strDisabled = ""
		If rsEXT2("ProcPay") <> "" OR rsEXT2("ProcMed") <> "" OR rsEXT2("ProcPriv") <> "" Then strDisabled = "readOnly"
		'response.write "EXT2 = " & sqlEXT2
	End If
		Xmon2 = "<br><input style='font-size: 8pt;' type='text' size= '3' class='rightjust' onblur='numericVal(this.value, this); ComputeRows2(); Compute2(this, 1);' name='hmonX2" & _
			lngI2 & "' value='" & XmonV2 & "' class='textbox' " & vbCrLf & _
			strDisabled & ">"
		Xtue2 = "<br><input style='font-size: 8pt;' type='text' size= '3' class='rightjust' onblur='numericVal(this.value, this); ComputeRows2(); Compute2(this, 1);' name='htueX2" & _
			lngI2 & "' value='" & XtueV2 & "' class='textbox' " & vbCrLf & _
			strDisabled & ">"
		Xwed2 = "<br><input style='font-size: 8pt;' type='text' size= '3' class='rightjust' onblur='numericVal(this.value, this); ComputeRows2(); Compute2(this, 1);' name='hwedX2" & _
			lngI2 & "' value='" & XwedV2 & "' class='textbox' " & vbCrLf & _
			strDisabled & ">"
		Xthu2 = "<br><input style='font-size: 8pt;' type='text' size= '3' class='rightjust' onblur='numericVal(this.value, this); ComputeRows2(); Compute2(this, 1);' name='hthuX2" & _
			lngI2 & "' value='" & XthuV2 & "' class='textbox' " & vbCrLf & _
			strDisabled & ">"
		Xfri2 = "<br><input style='font-size: 8pt;' type='text' size= '3' class='rightjust' onblur='numericVal(this.value, this); ComputeRows2(); Compute2(this, 1);' name='hfriX2" & _
			lngI2 & "' value='" & XfriV2 & "' class='textbox' " & vbCrLf & _
			strDisabled & ">"
		Xsat2 = "<br><input style='font-size: 8pt;' type='text' size= '3' class='rightjust' onblur='numericVal(this.value, this); ComputeRows2(); Compute2(this, 1);' name='hsatX2" & _
			lngI2 & "' value='" & XsatV2 & "' class='textbox' " & vbCrLf & _
			strDisabled & ">"
		Xsun2 = "<br><input style='font-size: 8pt;' type='text' size= '3' class='rightjust' onblur='numericVal(this.value, this); ComputeRows2(); Compute2(this, 1);' name='hsunX2" & _
			lngI2 & "' value='" & XsunV2  & "' class='textbox' " & vbCrLf & _
			strDisabled & ">"
		'response.write Xsun2
		XstrComputeCode2 = "+ parseFloat(document.frmTime.hmonX2" & lngI2 & ".value) + " & _
					"parseFloat(document.frmTime.htueX2" & lngI2 & ".value) + " & _
					"parseFloat(document.frmTime.hwedX2" & lngI2 & ".value) + " & _
					"parseFloat(document.frmTime.hthuX2" & lngI2 & ".value) + " & _
					"parseFloat(document.frmTime.hfriX2" & lngI2 & ".value) + " & _
					"parseFloat(document.frmTime.hsatX2" & lngI2 & ".value) + " & _
					"parseFloat(document.frmTime.hsunX2" & lngI2 & ".value);" & vbCrLf
					
		XstrComMon2 = "parseFloat(document.frmTime.hmonX2" & lngI2 & ".value) +"
		XstrComTue2 = "parseFloat(document.frmTime.htueX2" & lngI2 & ".value) +"
		XstrComWed2 = "parseFloat(document.frmTime.hwedX2" & lngI2 & ".value) +"
		XstrComThu2 = "parseFloat(document.frmTime.hthuX2" & lngI2 & ".value) +"
		XstrComFri2 = "parseFloat(document.frmTime.hfriX2" & lngI2 & ".value) +"
		XstrComSat2 = "parseFloat(document.frmTime.hsatX2" & lngI2 & ".value) +"
		XstrComSun2 = "parseFloat(document.frmTime.hsunX2" & lngI2 & ".value) +"
		
		strEXT2 = strEXT2 & "if (document.frmTime.chkEXT2" & lngI2 & ".checked == true) " & _
			"{document.frmTime.hsunX2" & lngI2 & ".style.visibility = 'visible'; " & _
			"document.frmTime.hmonX2" & lngI2 & ".style.visibility = 'visible'; " & _
			"document.frmTime.htueX2" & lngI2 & ".style.visibility = 'visible'; " & _
			"document.frmTime.hwedX2" & lngI2 & ".style.visibility = 'visible'; " & _
			 "document.frmTime.hthuX2" & lngI2 & ".style.visibility = 'visible'; " & _
			 "document.frmTime.hfriX2" & lngI2 & ".style.visibility = 'visible'; " & _
			 "document.frmTime.hsatX2" & lngI2 & ".style.visibility = 'visible';} " & _
		"else " &_
			"{document.frmTime.hsunX2" & lngI2 & ".style.visibility = 'hidden'; " & _
			 "document.frmTime.hsunX2" & lngI2 & ".value = 0; " & _
			 "document.frmTime.hmonX2" & lngI2 & ".style.visibility = 'hidden'; " & _
			 "document.frmTime.hmonX2" & lngI2 & ".value = 0; " & _
			 "document.frmTime.htueX2" & lngI2 & ".style.visibility = 'hidden'; " & _
			 "document.frmTime.htueX2" & lngI2 & ".value = 0; " & _
			 "document.frmTime.hwedX2" & lngI2 & ".style.visibility = 'hidden'; " & _
			 "document.frmTime.hwedX2" & lngI2 & ".value = 0; " & _
			 "document.frmTime.hthuX2" & lngI2 & ".style.visibility = 'hidden'; " & _
			 "document.frmTime.hthuX2" & lngI2 & ".value = 0; " & _
			 "document.frmTime.hfriX2" & lngI2 & ".style.visibility = 'hidden'; " & _
			 "document.frmTime.hfriX2" & lngI2 & ".value = 0; " & _
			 "document.frmTime.hsatX2" & lngI2 & ".value = 0; " & _
			 "document.frmTime.hsatX2" & lngI2 & ".style.visibility = 'hidden';} " & vbCrLf
		
	'End If	
	rsEXT2.Close
	Set rsEXT2 = Nothing
	strDisabled = ""
	If tblEMP("ProcPay") <> "" OR tblEMP("ProcMed") <> "" OR tblEMP("ProcPriv") <> "" Then strDisabled = "readOnly"
	noedit2 = ""
	'If tblEMP("ProcPay") <> "" OR tblEMP("ProcMed") <> "" OR tblEMP("ProcPriv") <> "" Then noedit2 = "disabled"
	strTableScript2 = strTableScript2 & "<tr bgcolor='" & kulay & "'><td align='center' rowspan='2' bgcolor='#C4B464'>" & vbCrLf & _
			"<input type='checkbox' size= '5' style='font-size: 8pt;' name='chkS" & lngI2 & "' value='" & _
			tblEMP("id") & "'" & noedit2 & "  onclick='RowDis2();'>" & _
			"<input type='hidden' name='tmpID2" & lngI2 & "' value='" & tblEMP("id") & "'></td>" & _
			"<td align='center' rowspan='2'><input style='font-size: 8pt;' type='text' readOnly name='hdept2" & lngI2 & "' size= '12' value=""" & pangalan & """ " & strDisabled & "" & _
			"><input type='hidden' name='conid2" & lngI2 & "' value='" & tblEMP("Client") & "'><br>" & _
			"<font face='trebuchet MS' size='1'>Extd Hrs.</font><input type='checkbox' size= '5' " & EXT2 & "  value='" & tblEMP("id") & "' name='chkEXT2" & lngI2 & "' onclick='GetEXT2();ComputeRows2();Compute2();'>" & _
			"</td" & vbCrLf

	strTableScript2 = strTableScript2 & "><td align='center' rowspan='2'><input style='font-size: 8pt;' type='text' size= '3' class='rightjust' onblur='numericVal(this.value, this); ComputeRows2(); Compute2(this, 1);' name='hsun2" & _
			lngI2 & "' value='" & tblEMP("sun") & "' class='textbox' " & vbCrLf & _
			strDisabled & ">" & xsun2 & "</td>" & vbCrLf & _
			"<td align='center' rowspan='2'><input style='font-size: 8pt;' type='text' size='3' name='hmon2" & _
			lngI2 & "' value='" & tblEMP("mon") & "' onblur='numericVal(this.value, this); ComputeRows2(); Compute2(this, 1);' class='rightjust'  " & _
			strDisabled & ">" & xmon2 & "</td" & vbCrLf & _
			"><td align='center' rowspan='2'><input style='font-size: 8pt;' type='text' onblur='numericVal(this.value, this); ComputeRows2(); Compute2(this, 1);' class='rightjust' size= '3' name='htue2" & _
			lngI2 & "' class='textbox' value='" & tblEMP("tue") & "' " & _
			strDisabled & ">" & xtue2 & "</td" & vbCrLf & _
			"><td align='center' rowspan='2'><input style='font-size: 8pt;' type='text' class='rightjust' size= '3' onblur='numericVal(this.value, this); ComputeRows2(); Compute2(this, 1);' name='hwed2" & _
			lngI2 & "' class='textbox' value='" & tblEMP("wed") & "' " & _
			strDisabled & ">" & xwed2 & "</td" & vbCrlf & _
			"><td align='center' rowspan='2'><input style='font-size: 8pt;' type='text' class='rightjust' size= '3' onblur='numericVal(this.value, this); ComputeRows2(); Compute2(this, 1);' name='hthu2" & _
			lngI2 & "' class='textbox' value='" & tblEMP("thu") &"' " & _
			strDisabled & ">" & xthu2 & "</td" & vbCrlf & _
			"><td align='center' rowspan='2'><input style='font-size: 8pt;' type='text' size= '3' class='rightjust' onblur='numericVal(this.value, this); ComputeRows2(); Compute2(this, 1);' name='hfri2" & _
			lngI2 & "' class='textbox' value='" & tblEMP ("fri") & "' " & _
			strDisabled & ">" & xfri2 & "</td" & vbCrlf & _
			"><td align='center' rowspan='2'><input style='font-size: 8pt;' type='text' size= '3' class='rightjust' onblur='numericVal(this.value, this); ComputeRows2(); Compute2(this, 1);' name='hsat2" & _
			lngI2 & "' value='" & tblEMP("sat") & "' class='textbox' " & _
			strDisabled & ">" & xsat2 & "</td" 
			
	
			strTableScript2 = strTableScript2 & "></td><td align='center' rowspan='2'><input style='font-size: 8pt;' name='htot2" & lngI2 & "' " & _
					" readOnly disabled class='rightjust'  size='3' value='" & ( Z_CDbl(tblEmp("mon")) + Z_CDbl(tblEmp("tue")) + _
						Z_CDbl(tblEmp("wed")) + Z_CDbl(tblEmp("thu")) + _
						Z_CDbl(tblEmp("fri")) + Z_CDbl(tblEmp("sat")) + _
						Z_CDbl(tblEmp("sun")) + _
						Z_CDbl(XsunV2) + Z_CDbl(XmonV2) + Z_CDbl(XtueV2) + Z_CDbl(XwedV2) + Z_CDbl(XthuV2) + Z_CDbl(XfriV2) + Z_CDbl(XsatV2)) & "' " & strDisabled & "></td" & vbCrLf
							
			'if Ucase(strDisabled) = "DISABLED" Then mileageAllow = ""
				
			procmile = ""	
			myMile = ""
			if z_fixnull(tblEMP("ProcMile")) <> "" then 
				procmile = "Readonly"
				myMile = "*"
			End If	
				
			strTableScript2 = strTableScript2 & "><td align='center' rowspan='2'><input style='font-size: 8pt;' type='text' size= '3' class='rightjust' onblur='numericVal(this.value, this); ComputeRows2(); Compute2(this, 1);' name='txt2mile" & _
			lngI2 & "' value='" & tblEMP("mile") & "' class='textbox' " & _
			mileageAllow & " " & procmile
			
			strTableScript2 = strTableScript2 & ">" & myMile & "</td><td align='center' rowspan='2'><input style='font-size: 8pt;' type='text' size= '3' class='rightjust' onblur='numericVal(this.value, this); ComputeRows2(); Compute2(this, 1);' name='txt2amile" & _
			lngI2 & "' value='" & tblEMP("amile") & "' class='textbox' " & _
			procmile
			
			strComputeCode2 = strComputeCode2 & "document.frmTime.htot2" & lngI2 & _
					".value = parseFloat(document.frmTime.hmon2" & lngI2 & ".value) + " & _
					"parseFloat(document.frmTime.htue2" & lngI2 & ".value) + " & _
					"parseFloat(document.frmTime.hwed2" & lngI2 & ".value) + " & _
					"parseFloat(document.frmTime.hthu2" & lngI2 & ".value) + " & _
					"parseFloat(document.frmTime.hfri2" & lngI2 & ".value) + " & _
					"parseFloat(document.frmTime.hsat2" & lngI2 & ".value) + " & _
					"parseFloat(document.frmTime.hsun2" & lngI2 & ".value) " & XstrComputeCode2 & ";" & vbCrLf
	
	strTableScript2 = strTableScript2 & "></td><td align='center' rowspan='2'><textarea style='font-size: 8pt;' readonly rows='2' name='Mnotes2" & _
			lngI2 & "' cols='15'>" & tblEMP("misc_notes") & _
		"</textarea><td align='center' rowspan='2'><textarea style='font-size: 8pt;' readonly rows='2' name='com2" & _
			lngI2 & "' cols='15'>" & tblEMP("actcode") & _
			"</textarea"  & vbCrLf 
		
	PayChk2 = ""
	medchk2 = ""
	'If tblEMP("ProcPay") <> "" Then PayChk2 = "checked"
	If tblEMP("ProcMed") <> "" Then MedChk2 = "checked"
	If tblEMP("ProcPriv") <> "" Then MedChk2 = "checked"
	If tblEMP("ProcVA") <> "" Then MedChk2 = "checked"
			
	strTableScript2 = strTableScript2 & " ></td><td bgcolor='#C4B464' align='center' rowspan='2'>" & vbCrLf & _
			"<input style='font-size: 8pt;' type='text' readOnly name='fon2" & lngI2 & "' size= '12' value=""" & tblEMP("CallerID") & """ " & strDisabled & "" & _
			"></td><td bgcolor='#C4B464' align='center' rowspan='2'>" & vbCrLf & _
			"<input style='font-size: 8pt;' type='checkbox' size= '5' name='chkM2" & lngI2 & "' value='" & _
			tblEMP("ProcMed") & "' " & MedChk2 & " disabled ></td>" & vbCrLf 
			
	If (Session("lngType")) = 2 AND (Paychk2 = "checked" OR Medchk2 = "checked") Then
		strTableScript2 = strTableScript2	& "<td valign='middle' bgcolor='#C4B464' rowspan='2'><input type='button' value='<< Untag' " & _
		"onclick='Untag2(" & tblEMP("id") & ");' class='btn' onmouseover=""this.className='hovbtn'"" onmouseout=""this.className='btn'""></td>" & vbCrLf 
	End If		
	
	strTableScript2 = strTableScript2	& "<tr><td>&nbsp;</td></tr>" & vbCrLf 
	
	
	strComMon2 = strComMon2 & "parseFloat(document.frmTime.hmon2" & lngI2 & ".value) + " & XstrComMon2
	strComTue2 = strComTue2 & "parseFloat(document.frmTime.htue2" & lngI2 & ".value) + " & XstrComTue2
	strComWed2 = strComWed2 & "parseFloat(document.frmTime.hwed2" & lngI2 & ".value) + " & XstrComWed2
	strComThu2 = strComThu2 & "parseFloat(document.frmTime.hthu2" & lngI2 & ".value) + " & XstrComThu2
	strComFri2 = strComFri2 & "parseFloat(document.frmTime.hfri2" & lngI2 & ".value) + " & XstrComFri2
	strComSat2 = strComSat2 & "parseFloat(document.frmTime.hsat2" & lngI2 & ".value) + " & XstrComSat2
	strComSun2 = strComSun2 & "parseFloat(document.frmTime.hsun2" & lngI2 & ".value) + " & XstrComSun2
	strComMile2 = strComMile2 & "parseFloat(document.frmTime.txt2mile" & lngI2 & ".value) + "
	strComaMile2 = strComaMile2 & "parseFloat(document.frmTime.txt2amile" & lngI2 & ".value) + "
	
	
	If procmile <> "Readonly" Then
		If tblEMP("ProcPay") <> "" OR tblEMP("ProcMed") <> "" OR tblEMP("ProcPriv") <> "" Then 
			strDis2 = strDis2 & "if (document.frmTime.chkS" & lngI2 & ".checked == true) " & vbCrlf &_
				"{document.frmTime.chkEXT2" & lngI2 & ".disabled = true; " & vbCrlf &_
				"document.frmTime.hsun2" & lngI2 & ".readOnly = true; " & vbCrlf &_
				"document.frmTime.hmon2" & lngI2 & ".readOnly = true; " & vbCrlf &_
				"document.frmTime.htue2" & lngI2 & ".readOnly = true; " & vbCrlf &_
				"document.frmTime.hwed2" & lngI2 & ".readOnly = true; " & vbCrlf &_
				"document.frmTime.hthu2" & lngI2 & ".readOnly = true; " & vbCrlf &_
				"document.frmTime.hfri2" & lngI2 & ".readOnly = true; " & vbCrlf &_
				"document.frmTime.hsat2" & lngI2 & ".readOnly = true; " & vbCrlf &_
				"document.frmTime.hsunX2" & lngI2 & ".readOnly = true; " & vbCrlf &_
				"document.frmTime.hmonX2" & lngI2 & ".readOnly = true; " & vbCrlf &_
				"document.frmTime.htueX2" & lngI2 & ".readOnly = true; " & vbCrlf &_
				"document.frmTime.hwedX2" & lngI2 & ".readOnly = true; " & vbCrlf &_
				"document.frmTime.hthuX2" & lngI2 & ".readOnly = true; " & vbCrlf &_
				"document.frmTime.hfriX2" & lngI2 & ".readOnly = true; " & vbCrlf &_
				"document.frmTime.hsatX2" & lngI2 & ".readOnly = true; " & vbCrlf
				if AllowMileage(Session("idemp")) Then 
					strDis2 = strDis2 & "document.frmTime.txt2mile" & lngI2 & ".readOnly = false; "
				else
					strDis2 = strDis2 & "document.frmTime.txt2mile" & lngI2 & ".readOnly = true; "
				end if
				strDis2 = strDis2 & "document.frmTime.txt2amile" & lngI2 & ".readOnly = false; " & vbCrlf &_
					"document.frmTime.Mnotes2" & lngI2 & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.fon2" & lngI2 & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.com2" & lngI2 & ".readOnly  = true;} " & vbCrlf &_
			"else " & vbCrLf & vbCrlf &_
				"{document.frmTime.chkEXT2" & lngI2 & ".disabled = true; " & vbCrlf &_
				"document.frmTime.hsun2" & lngI2 & ".readOnly = true; " & vbCrlf &_
				"document.frmTime.hmon2" & lngI2 & ".readOnly = true; " & vbCrlf &_
				"document.frmTime.htue2" & lngI2 & ".readOnly = true; " & vbCrlf &_
				"document.frmTime.hwed2" & lngI2 & ".readOnly = true; " & vbCrlf &_
				"document.frmTime.hthu2" & lngI2 & ".readOnly = true; " & vbCrlf &_
				"document.frmTime.hfri2" & lngI2 & ".readOnly = true; " & vbCrlf &_
				"document.frmTime.hsat2" & lngI2 & ".readOnly = true; " & vbCrlf &_
				"document.frmTime.hsunX2" & lngI2 & ".readOnly = true; " & vbCrlf &_
				"document.frmTime.hmonX2" & lngI2 & ".readOnly = true; " & vbCrlf &_
				"document.frmTime.htueX2" & lngI2 & ".readOnly = true; " & vbCrlf &_
				"document.frmTime.hwedX2" & lngI2 & ".readOnly = true; " & vbCrlf &_
				"document.frmTime.hthuX2" & lngI2 & ".readOnly = true; " & vbCrlf &_
				"document.frmTime.hfriX2" & lngI2 & ".readOnly = true; " & vbCrlf &_
				"document.frmTime.hsatX2" & lngI2 & ".readOnly = true; " & vbCrlf &_
				"document.frmTime.txt2mile" & lngI2 & ".readOnly  = true; " & vbCrlf &_
				"document.frmTime.txt2amile" & lngI2 & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.Mnotes2" & lngI2 & ".readOnly  = true; " & vbCRlf &_
					"document.frmTime.fon2" & lngI2 & ".readOnly  = true; " & vbCRlf &_
					"document.frmTime.com2" & lngI2 & ".readOnly  = true;} " & vbCrlf
			Else
				'response.write "2"
				strDis2 = strDis2 & "if (document.frmTime.chkS" & lngI2 & ".checked == true) " & vbCrlf &_
				"{document.frmTime.chkEXT2" & lngI2 & ".disabled = false; " & vbCrlf &_
				"document.frmTime.hsun2" & lngI2 & ".readOnly = false; " & vbCrlf &_
				"document.frmTime.hmon2" & lngI2 & ".readOnly = false; " & vbCrlf &_
				"document.frmTime.htue2" & lngI2 & ".readOnly = false; " & vbCrlf &_
				"document.frmTime.hwed2" & lngI2 & ".readOnly = false; " & vbCrlf &_
				"document.frmTime.hthu2" & lngI2 & ".readOnly = false; " & vbCrlf &_
				"document.frmTime.hfri2" & lngI2 & ".readOnly = false; " & vbCrlf &_
				"document.frmTime.hsat2" & lngI2 & ".readOnly = false; " & vbCrlf &_
				"document.frmTime.hsunX2" & lngI2 & ".readOnly = false; " & vbCrlf &_
				"document.frmTime.hmonX2" & lngI2 & ".readOnly = false; " & vbCrlf &_
				"document.frmTime.htueX2" & lngI2 & ".readOnly = false; " & vbCrlf &_
				"document.frmTime.hwedX2" & lngI2 & ".readOnly = false; " & vbCrlf &_
				"document.frmTime.hthuX2" & lngI2 & ".readOnly = false; " & vbCrlf &_
				"document.frmTime.hfriX2" & lngI2 & ".readOnly = false; " & vbCrlf &_
				"document.frmTime.hsatX2" & lngI2 & ".readOnly = false; " & vbCrlf
				if AllowMileage(Session("idemp")) Then 
					strDis2 = strDis2 & "document.frmTime.txt2mile" & lngI2 & ".readOnly = false; "
				else
					strDis2 = strDis2 & "document.frmTime.txt2mile" & lngI2 & ".readOnly = true; "
				end if
				strDis2 = strDis2 & "document.frmTime.txt2amile" & lngI2 & ".readOnly = false; " & vbCrlf &_
					"document.frmTime.Mnotes2" & lngI2 & ".readOnly  = false; " & vbCrlf &_
						"document.frmTime.fon2" & lngI2 & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.com2" & lngI2 & ".readOnly  = false;} " & vbCrlf &_
			"else " & vbCrLf & _
				"{document.frmTime.chkEXT2" & lngI2 & ".disabled = true; " & vbCrlf &_
				"document.frmTime.hsun2" & lngI2 & ".readOnly = true; " & vbCrlf &_
				"document.frmTime.hmon2" & lngI2 & ".readOnly = true; " & vbCrlf &_
				"document.frmTime.htue2" & lngI2 & ".readOnly = true; " & vbCrlf &_
				"document.frmTime.hwed2" & lngI2 & ".readOnly = true; " & vbCrlf &_
				"document.frmTime.hthu2" & lngI2 & ".readOnly = true; " & vbCrlf &_
				"document.frmTime.hfri2" & lngI2 & ".readOnly = true; " & vbCrlf &_
				"document.frmTime.hsat2" & lngI2 & ".readOnly = true; " & vbCrlf &_
				"document.frmTime.hsunX2" & lngI2 & ".readOnly = true; " & vbCrlf &_
				"document.frmTime.hmonX2" & lngI2 & ".readOnly = true; " & vbCrlf &_
				"document.frmTime.htueX2" & lngI2 & ".readOnly = true; " & vbCrlf &_
				"document.frmTime.hwedX2" & lngI2 & ".readOnly = true; " & vbCrlf &_
				"document.frmTime.hthuX2" & lngI2 & ".readOnly = true; " & vbCrlf &_
				"document.frmTime.hfriX2" & lngI2 & ".readOnly = true; " & vbCrlf &_
				"document.frmTime.hsatX2" & lngI2 & ".readOnly = true; " & vbCrlf &_
				"document.frmTime.txt2mile" & lngI2 & ".readOnly  = true; " & vbCrlf &_
				"document.frmTime.txt2amile" & lngI2 & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.Mnotes2" & lngI2 & ".readOnly  = true; " & vbCRlf &_
					"document.frmTime.fon2" & lngI2 & ".readOnly  = true; " & vbCRlf &_
					"document.frmTime.com2" & lngI2 & ".readOnly  = true;} " & vbCRlf
			end if
		Else
			If tblEMP("ProcPay") <> "" OR tblEMP("ProcMed") <> "" OR tblEMP("ProcPriv") <> "" Then
				'response.write "3"
				strDis2 = strDis2 & "if (document.frmTime.chkS" & lngI2 & ".checked == true) " & _
				"{document.frmTime.chkEXT2" & lngI2 & ".disabled = true; " & vbCrlf &_
				"document.frmTime.hsun2" & lngI2 & ".readOnly = true; " & vbCrlf &_
				"document.frmTime.hmon2" & lngI2 & ".readOnly = true; " & vbCrlf &_
				"document.frmTime.htue2" & lngI2 & ".readOnly = true; " & vbCrlf &_
				"document.frmTime.hwed2" & lngI2 & ".readOnly = true; " & vbCrlf &_
				"document.frmTime.hthu2" & lngI2 & ".readOnly = true; " & vbCrlf &_
				"document.frmTime.hfri2" & lngI2 & ".readOnly = true; " & vbCrlf &_
				"document.frmTime.hsat2" & lngI2 & ".readOnly = true; " & vbCrlf &_
				"document.frmTime.hsunX2" & lngI2 & ".readOnly = true; " & vbCrlf &_
				"document.frmTime.hmonX2" & lngI2 & ".readOnly = true; " & vbCrlf &_
				"document.frmTime.htueX2" & lngI2 & ".readOnly = true; " & vbCrlf &_
				"document.frmTime.hwedX2" & lngI2 & ".readOnly = true; " & vbCrlf &_
				"document.frmTime.hthuX2" & lngI2 & ".readOnly = true; " & vbCrlf &_
				"document.frmTime.hfriX2" & lngI2 & ".readOnly = true; " & vbCrlf &_
				"document.frmTime.hsatX2" & lngI2 & ".readOnly = true; " & vbCrlf
				if AllowMileage(Session("idemp")) Then 
					strDis2 = strDis2 & "document.frmTime.txt2mile" & lngI2 & ".readOnly = true; "
				else
					strDis2 = strDis2 & "document.frmTime.txt2mile" & lngI2 & ".readOnly = true; "
				end if
				strDis2 = strDis2 & "document.frmTime.txt2amile" & lngI2 & ".readOnly = true; " & _
				"document.frmTime.Mnotes2" & lngI2 & ".readOnly  = false; " & vbCrlf &_
						"document.frmTime.fon2" & lngI2 & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.com2" & lngI2 & ".readOnly  = false;} " & vbCrlf &_
			"else " & vbCrLf & _
				"{document.frmTime.chkEXT2" & lngI2 & ".disabled = true; " & vbCrlf &_
				"document.frmTime.hsun2" & lngI2 & ".readOnly = true; " & vbCrlf &_
				"document.frmTime.hmon2" & lngI2 & ".readOnly = true; " & vbCrlf &_
				"document.frmTime.htue2" & lngI2 & ".readOnly = true; " & vbCrlf &_
				"document.frmTime.hwed2" & lngI2 & ".readOnly = true; " & vbCrlf &_
				"document.frmTime.hthu2" & lngI2 & ".readOnly = true; " & vbCrlf &_
				"document.frmTime.hfri2" & lngI2 & ".readOnly = true; " & vbCrlf &_
				"document.frmTime.hsat2" & lngI2 & ".readOnly = true; " & vbCrlf &_
				"document.frmTime.hsunX2" & lngI2 & ".readOnly = true; " & vbCrlf &_
				"document.frmTime.hmonX2" & lngI2 & ".readOnly = true; " & vbCrlf &_
				"document.frmTime.htueX2" & lngI2 & ".readOnly = true; " & vbCrlf &_
				"document.frmTime.hwedX2" & lngI2 & ".readOnly = true; " & vbCrlf &_
				"document.frmTime.hthuX2" & lngI2 & ".readOnly = true; " & vbCrlf &_
				"document.frmTime.hfriX2" & lngI2 & ".readOnly = true; " & vbCrlf &_
				"document.frmTime.hsatX2" & lngI2 & ".readOnly = true; " & vbCrlf &_
				"document.frmTime.txt2mile" & lngI2 & ".readOnly  = true; " & vbCrlf &_
				"document.frmTime.txt2amile" & lngI2 & ".readOnly  = true; " & vbCrlf &_
				"document.frmTime.Mnotes2" & lngI2 & ".readOnly = true;} " & vbCRlf
			else
				'response.write "4"
				strDis2 = strDis2 & "if (document.frmTime.chkS" & lngI2 & ".checked == true) " & vbCrlf &_
					"{document.frmTime.chkEXT2" & lngI2 & ".disabled = false; " & vbCrlf &_
					"document.frmTime.hsun2" & lngI2 & ".readOnly  = false; " & vbCrlf &_
					"document.frmTime.hmon2" & lngI2 & ".readOnly  = false; " & vbCrlf &_
					"document.frmTime.htue2" & lngI2 & ".readOnly  = false; " & vbCrlf &_
					"document.frmTime.hwed2" & lngI2 & ".readOnly  = false; " & vbCrlf &_
					"document.frmTime.hthu2" & lngI2 & ".readOnly  = false; " & vbCrlf &_
					"document.frmTime.hfri2" & lngI2 & ".readOnly  = false; " & vbCrlf &_
					"document.frmTime.hsat2" & lngI2 & ".readOnly  = false; " & vbCrlf &_
					"document.frmTime.hsunX2" & lngI2 & ".readOnly  = false; " & vbCrlf &_
					"document.frmTime.hmonX2" & lngI2 & ".readOnly  = false; " & vbCrlf &_
					"document.frmTime.htueX2" & lngI2 & ".readOnly  = false; " & vbCrlf &_
					"document.frmTime.hwedX2" & lngI2 & ".readOnly  = false; " & vbCrlf &_
					"document.frmTime.hthuX2" & lngI2 & ".readOnly  = false; " & vbCrlf &_
					"document.frmTime.hfriX2" & lngI2 & ".readOnly  = false; " & vbCrlf &_
					"document.frmTime.hsatX2" & lngI2 & ".readOnly  = false; " & vbCrlf
					if AllowMileage(Session("idemp")) Then 
						strDis2 = strDis2 & "document.frmTime.txt2mile" & lngI2 & ".readOnly  = true; "
					else
						strDis2 = strDis2 & "document.frmTime.txt2mile" & lngI2 & ".readOnly  = true; "
					end if
					strDis2 = strDis2 & "document.frmTime.txt2amile" & lngI2 & ".readOnly  = false; " & vbCrlf &_
					"document.frmTime.Mnotes2" & lngI2 & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.fon2" & lngI2 & ".readOnly  = true; " & vbCrlf &_
						"document.frmTime.com2" & lngI2 & ".readOnly  = true;} " & vbCrlf &_
				"else " & vbCrLf & _
					"{document.frmTime.chkEXT2" & lngI2 & ".disabled = true; " & vbCrlf &_
					"document.frmTime.hsun2" & lngI2 & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.hmon2" & lngI2 & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.htue2" & lngI2 & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.hthu2" & lngI2 & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.hfri2" & lngI2 & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.hsat2" & lngI2 & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.hsunX2" & lngI2 & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.hmonX2" & lngI2 & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.htueX2" & lngI2 & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.hwedX2" & lngI2 & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.hthuX2" & lngI2 & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.hfriX2" & lngI2 & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.hsatX2" & lngI2 & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.txt2mile" & lngI2 & ".readOnly  = true; " & vbCrlf &_
					"document.frmTime.txt2amile" & lngI2 & ".readOnly  = true; " & vbCrlf &_
						"document.frmTime.Mnotes2" & lngI2 & ".readOnly  = true; " & vbCRlf &_
						"document.frmTime.fon2" & lngI2 & ".readOnly  = true; " & vbCRlf &_
					"document.frmTime.com2" & lngI2 & ".readOnly  = true;} " & vbCrlf
			end if
		End If
	
	lngI2 = lngI2 + 1
	tblEMP.MoveNext
Loop

tblEmp.Close
Set tblEMP = Nothing
'PTO
tmpPTO2 = 0
Set rsPTO = Server.CreateObject("ADODB.RecordSet")
sqlPTO = "SELECT * FROM W_PTO_T WHERE WorkerID = '" & Session("idemp") & "' AND date = '" & sunDATE2 & "'"
rsPTO.Open sqlPTO, g_strCONN, 1, 3
If Not rsPTO.EOF Then
	tmpPTO2 = rsPTO("PTO")
	PTO2 = ""
	If Not IsNull(rsPTO("procItem")) Then PTO2 = "readOnly"
End If
rsPTO.Close
Set rsPTO = Nothing
''''''''''''''''''''consumer dropdown WK 1
Set tblIdx = Server.CreateObject("ADODB.RecordSet")
sqlIdx = "SELECT * FROM [Worker_t] WHERE [Social_Security_Number] = '" & Session("idemp") & "' "
tblIdx.Open sqlIdx, g_strCONN, 3, 1
fltwork = False
If not tblIdx.EOF Then
	tmpWorkID = tblIdx("index")
	If tblIdx("flt") Then fltwork = True
End If
tblIdx.Close
Set tblIDX = Nothing	
If Not fltwork Then 
	Set tblCW = Server.CreateObject("ADODB.RecordSet")
	Set tbldept = Server.CreateObject("ADODB.Recordset")
	'Set tblChkCon = Server.CreateObject("ADODB.Recordset")
	sqlCW = "SELECT * FROM [ConWork_t] WHERE WID = '" & tmpWorkID & "' " 
	tblCW.Open sqlCW, g_strCONN, 3, 1
	Do Until tblCW.EOF
		strSQLd = "SELECT * FROM [Consumer_t], C_Status_t WHERE [Consumer_t].[Medicaid_Number] = '" & tblCW("CID") & "' AND [Consumer_t].[Medicaid_Number] = " & _
			"[C_Status_t].[Medicaid_Number] AND Active = 1 ORDER BY lname, fname"
		tbldept.Open strSQLd, g_strCONN, 1, 3
		If Not tbldept.EOF Then
				Xsel = ""
				If Request("Con") = tbldept("Medicaid_number") Then Xsel = "SELECTED"
					myCode = ""
					If tblDept("Code") = "P" Or tblDept("Code") = "C" Then myCode = "*"
				strdept = strdept & "<option value='" & tbldept("Medicaid_Number")& "' " & Xsel & "> "& myCode & tbldept("Lname") & ", " & tbldept("fname") & " </option>"  
			'End If
			DisNew = ""
			If strdept = "" Then DisNew = "DISABLED"
			'tblChkCon.close
		End If
		tbldept.Close
		tblCW.MoveNext
	Loop
	tblCW.Close
	set tblCW = Nothing
	'Set tblChkCon = Nothing
	Set tbldept = Nothing
Else
	Set tbldept = Server.CreateObject("ADODB.Recordset")
	strSQLd = "SELECT * FROM [Consumer_t], C_Status_t WHERE [Consumer_t].[Medicaid_Number] = " & _
			"[C_Status_t].[Medicaid_Number] AND Active = 1 ORDER BY lname, fname"
	tbldept.Open strSQLd, g_strCONN, 3, 1
	Do Until tblDept.EOF
		Xsel = ""
		If Request("Con") = tbldept("Medicaid_number") Then Xsel = "SELECTED"
		myCode = ""
		If tblDept("Code") = "P" Or tblDept("Code") = "C" Then myCode = "*"
		strdept = strdept & "<option value='" & tbldept("Medicaid_Number")& "' " & Xsel & "> "& myCode & tbldept("Lname") & ", " & tbldept("fname") & " </option>"  
		DisNew = ""
		If strdept = "" Then DisNew = "DISABLED"
		TblDept.MoveNext
	Loop
	TblDept.Close
	Set TblDept = Nothing
End If
'''''''''''''''''''''
Set tblUser = Server.CreateObject("ADODB.Recordset")
sqlUser = "SELECT * FROM [input_t] WHERE [index] = " & session("UserID")
'response.write "<!-- sql" & sqlUser & "-->"
tblUser.Open sqlUser, g_strCONN, 3, 1
	If session("UserID") = "" Then
		session("MSG") = "Session timed out. Sign in again."
		response.redirect "default.asp"
	End IF
	'tmpUser = "ADMINISTRATOR"
	if not tblUser.EOF Then
		tmpUser = UCase(tblUser("lname")) & ", " & UCase(tblUser("fname"))
		'response.write "<!-- name" & tmpUser & "-->"
	else
		session("MSG") = "Session timed out. Sign in again."
		response.redirect "default.asp"
		
	End If
tblUser.Close
set tblUser = Nothing

'''''''''''''''''''consumer dropdown WK 2
Set tblIdx = Server.CreateObject("ADODB.RecordSet")
sqlIdx = "SELECT * FROM [Worker_t] WHERE [Social_Security_Number] = '" & Session("idemp") & "' "
tblIdx.Open sqlIdx, g_strCONN, 3, 1
If not tblIdx.EOF Then
	tmpWorkID = tblIdx("index")
	If tblIdx("flt") Then fltwork = True
End If
tblIdx.Close
Set tblIDX = Nothing	
If Not fltwork Then 
	Set tblCW = Server.CreateObject("ADODB.RecordSet")
	Set tbldept = Server.CreateObject("ADODB.Recordset")
	sqlCW = "SELECT * FROM [ConWork_t] WHERE WID = '" & tmpWorkID & "' " 
	tblCW.Open sqlCW, g_strCONN, 3, 1
	Do Until tblCW.EOF
		strSQLd = "SELECT * FROM [Consumer_t], C_Status_t WHERE [Consumer_t].[Medicaid_Number] = '" & tblCW("CID") & "' AND [Consumer_t].[Medicaid_Number] = " & _
			"[C_Status_t].[Medicaid_Number] AND Active = 1  ORDER BY lname, fname"
		tbldept.Open strSQLd, g_strCONN, 1, 3
		If Not tbldept.EOF Then
			Xsel = ""
			If Request("Con2") = tbldept("Medicaid_number") Then Xsel = "SELECTED"
			myCode = ""
			If tblDept("Code") = "P" Or tblDept("Code") = "C" Then myCode = "*"
			strdept2 = strdept2 & "<option value='" & tbldept("Medicaid_Number")& "' " & Xsel & "> "& mycode & tbldept("Lname") & ", " & tbldept("fname") & " </option>"  
			DisNew2 = ""
			If strdept2 = "" Then DisNew2 = "DISABLED"
		End If
		tbldept.Close
		tblCW.MoveNext
	Loop
	tblCW.Close
	set tblCW = Nothing
	Set tbldept = Nothing
Else
	Set tbldept = Server.CreateObject("ADODB.Recordset")
	strSQLd = "SELECT * FROM [Consumer_t], C_Status_t WHERE [Consumer_t].[Medicaid_Number] = " & _
			"[C_Status_t].[Medicaid_Number] AND Active = 1 ORDER BY lname, fname"
	tbldept.Open strSQLd, g_strCONN, 3, 1
	Do Until tblDept.EOF	
		Xsel = ""
		If Request("Con2") = tbldept("Medicaid_number") Then Xsel = "SELECTED"
		myCode = ""
		If tblDept("Code") = "P" Or tblDept("Code") = "C" Then myCode = "*"
		strdept2 = strdept2 & "<option value='" & tbldept("Medicaid_Number")& "' " & Xsel & "> "& mycode & tbldept("Lname") & ", " & tbldept("fname") & " </option>"  
		DisNew2 = ""
		If strdept2 = "" Then DisNew2 = "DISABLED"
		TblDept.MoveNext
	Loop
	TblDept.Close
	Set TblDept = Nothing
End If
'''''''''''''process 1
'Paychk = ""
'Medchk = ""
'Set tblchk = server.createobject("ADODB.RecordSet")
'sqlchk = "Select * from process_t Where Wor = '"& Session("idemp") &"' and TsDate = #" &  sunDATE & "# "
'tblchk.OPen sqlchk, g_strCONN, 1, 3
'if Not tblchk.EOF then
'	If tblchk("procPay") = true Then Paychk = "checked"
'	If tblchk("procMed") = true Then Medchk = "checked"	
'end if
'tblchk.close
'set tblchk = nothing
'DisAdmin = "Disabled"
'If UCase(Session("lngType")) = "TRUE" Then DisAdmin = ""
'AdminDis = ""
'If Paychk <> "" Or Medchk <> "" Then AdminDis = "DISABLED"

'''''''''''process 2
'Paychk2 = ""
'Medchk2 = ""
'Set tblchk = server.createobject("ADODB.RecordSet")
'sqlchk = "Select * from process_t Where Wor = '"& Session("idemp") &"' and TsDate = #" &  sunDATE2 & "# "
'tblchk.OPen sqlchk, g_strCONN, 1, 3
'if Not tblchk.EOF then
'	If tblchk("procPay") = true Then Paychk2 = "checked"
'	If tblchk("procMed") = true Then Medchk2 = "checked"	
'end if
'tblchk.close
'set tblchk = nothing
'DisAdmin2 = "Disabled"

'''''RESUPPLY INPUTS ON ERROR
	monh = 0
	tueh = 0
	wedh = 0
	thuh = 0
	frih = 0
	sath = 0
	sunh = 0
	mile1 = 0
	amile1 = 0
	notesh = ""
	monhX = 0
	tuehX = 0
	wedhX = 0
	thuhX = 0
	frihX = 0
	sathX = 0
	sunhX = 0
	monh2 = 0
	tueh2 = 0
	wedh2 = 0
	thuh2 = 0
	frih2 = 0
	sath2 = 0
	sunh2 = 0
	mile2 = 0
	amile2 = 0
	notesh2 = ""
	monhX2 = 0
	tuehX2 = 0
	wedhX2 = 0
	thuhX2 = 0
	frihX2 = 0
	sathX2 = 0
	sunhX2 = 0
	If Request("xHrs") <> "" Then
		GetHrs = Split(Z_DoDeCrypt(Request("xHrs")), "|")
		monh = GetHrs(0)
		tueh = GetHrs(1)
		wedh = GetHrs(2)
		thuh = GetHrs(3)
		frih = GetHrs(4)
		sath = GetHrs(5)
		sunh = GetHrs(6)
		notesh = GetHrs(7)
		monhX = GetHrs(8)
		tuehX = GetHrs(9)
		wedhX = GetHrs(10)
		thuhX = GetHrs(11)
		frihX = GetHrs(12)
		sathX = GetHrs(13)
		sunhX = GetHrs(14)
		com1 = GetHrs(15)
		fon1 = GetHrs(16)
		ExtX = ""
		THrsX = monhX + tuehX +	wedhX + thuhX +	frihX + sathX + sunhX 
		If THrsX <> 0 Then ExtX = "CHECKED"
	End If
	If Request("xHrs2") <> "" Then
		GetHrs = Split(Z_DoDeCrypt(Request("xHrs2")), "|")
		monh2 = GetHrs(0)
		tueh2 = GetHrs(1)
		wedh2 = GetHrs(2)
		thuh2 = GetHrs(3)
		frih2 = GetHrs(4)
		sath2 = GetHrs(5)
		sunh2 = GetHrs(6)
		notesh2 = GetHrs(7)
		monhX2 = GetHrs(8)
		tuehX2 = GetHrs(9)
		wedhX2 = GetHrs(10)
		thuhX2 = GetHrs(11)
		frihX2 = GetHrs(12)
		sathX2 = GetHrs(13)
		sunhX2 = GetHrs(14)
		com2 = GetHrs(15)
		fon2 = GetHrs(16)
		ExtX2 = ""
		THrsX2 = monhX2 + tuehX2 + wedhX2 + thuhX2 + frihX2 + sathX2 + sunhX2 
		If THrsX2 <> 0 Then ExtX2 = "CHECKED"
	End If
'''''
%>
<html>
<head>
<title>LSS - In-Home Care - Timesheet</title>
<link href="styles.css" type="text/css" rel="stylesheet" media="print,screen">
<SCRIPT LANGUAGE="JavaScript"><!--
function numericVal(nbr, theCELL)
 {
	var re, strXX, lngXX;
	strXX = new String(nbr);
	var re = new RegExp(',','gi');
	strXX = strXX.replace(re,"");
	lngXX = new Number(strXX);
	if (isNaN(lngXX)) {
	    theCELL.value =  0
	}
    theCELL.value =  lngXX;
}	
function Untag(zzz)
	{
		var ans = window.confirm("This action will untag Week 1 - Timesheet ID: " + zzz + ". Cancel to stop.");
		if (ans){
		document.frmTime.action = "Untag.asp?ID=" + zzz;
		document.frmTime.submit();
		}
	}
function Untag2(zzz)
{
	var ans = window.confirm("This action will untag Week 2 - Timesheet ID: " + zzz + ". Cancel to stop.");
	if (ans){
	document.frmTime.action = "Untag2.asp?ID=" + zzz;
	document.frmTime.submit();
	}
}
function GotoD8()
{
	document.frmTime.action = "GotoD8.asp";
	document.frmTime.submit();
}
function ComputeRows() {
	<%=strComputeCode%>
}
function ComputeRows2() {
	<%=strComputeCode2%>
}
function RowDis() {
	<%=strDis%>
}
function RowDis2() {
	<%=strDis2%>
}
function Compute(theCELL, blnUse)
{
var tmpHRS;
	
 tmpHRS = parseFloat(document.frmTime.hmon1.value) +  parseFloat(document.frmTime.htue1.value) + parseFloat(document.frmTime.hwed1.value) +  parseFloat(document.frmTime.hthu1.value) + parseFloat(document.frmTime.hfri1.value) +  parseFloat(document.frmTime.hsat1.value) + parseFloat(document.frmTime.hsun1.value);
 tmpHRSX = parseFloat(document.frmTime.hmonX1.value) +  parseFloat(document.frmTime.htueX1.value) + parseFloat(document.frmTime.hwedX1.value) +  parseFloat(document.frmTime.hthuX1.value) + parseFloat(document.frmTime.hfriX1.value) +  parseFloat(document.frmTime.hsatX1.value) + parseFloat(document.frmTime.hsunX1.value);
 document.frmTime.htot1.value = tmpHRS + tmpHRSX;

	document.frmTime.mileTOT.value = <%=strComMile%> parseFloat(document.frmTime.txtmile.value);
	document.frmTime.mileaTOT.value = <%=strComaMile%> parseFloat(document.frmTime.txtamile.value);
	
	document.frmTime.thmon.value = <%=strComMon%> parseFloat(document.frmTime.hmonX1.value) + parseFloat(document.frmTime.hmon1.value);
	if (parseFloat(document.frmTime.thmon.value) > 24)
		{
		alert("Total hours for Monday exceeds 24 hours.");
		if (blnUse) {
			theCELL.style.color = 'red';
			theCELL.value = 0;
			theCELL.focus();
	  } else {
	  	  theCELL.style.color = 'red';
		  theCELL.value = 0;
		}
		//Compute();
		return false;
		}
		
	document.frmTime.thtue.value = <%=strComTue%> parseFloat(document.frmTime.htue1.value) + parseFloat(document.frmTime.htueX1.value);
	if (parseFloat(document.frmTime.thtue.value) > 24)
		{
		alert("Total hours for Tuesday exceeds 24 hours.");
		if (blnUse) {
			theCELL.style.color = 'red';
			theCELL.value = 0;
			theCELL.focus();
	  } else {
	  	theCELL.style.color = 'red';
		theCELL.value = 0;
		}
		return false;
		} 
 	document.frmTime.thwed.value = <%=strComWed%> parseFloat(document.frmTime.hwed1.value) + parseFloat(document.frmTime.hwedX1.value);
	if (parseFloat(document.frmTime.thwed.value) > 24)
		{
		alert("Total hours for Wednesday exceeds 24 hours.");
		if (blnUse) {
			theCELL.style.color = 'red';
			theCELL.value = 0;
			theCELL.focus();
	  } else {
	  	theCELL.style.color = 'red';
	  	theCELL.value = 0;
		}
		return false;
		}
 	document.frmTime.ththu.value = <%=strComThu%> parseFloat(document.frmTime.hthu1.value) + parseFloat(document.frmTime.hthuX1.value);
 	if (parseFloat(document.frmTime.ththu.value) > 24)
		{
		alert("Total hours for Thursday exceeds 24 hours.");
		if (blnUse) {
			theCELL.style.color = 'red';
			theCELL.value = 0;
			theCELL.focus();
	  } else {
	  	theCELL.style.color = 'red';
		theCELL.value = 0;
		}
		return false;
		}
 	document.frmTime.thfri.value = <%=strComFri%> parseFloat(document.frmTime.hfri1.value) + parseFloat(document.frmTime.hfriX1.value);
 	if (parseFloat(document.frmTime.thfri.value) > 24)
		{
		alert("Total hours for Friday exceeds 24 hours.");
		if (blnUse) {
			theCELL.style.color = 'red';
			theCELL.value = 0;
			theCELL.focus();
	  } else {
	  	theCELL.style.color = 'red';
		theCELL.hfri.value = 0;
		}
		return false;
		}
 	document.frmTime.thsat.value = <%=strComSat%> parseFloat(document.frmTime.hsat1.value) + parseFloat(document.frmTime.hsatX1.value);
 	if (parseFloat(document.frmTime.thsat.value) > 24)
		{
		alert("Total hours for Saturday exceeds 24 hours.");
		if (blnUse) {
			theCELL.style.color = 'red';
			theCELL.value = 0;
			theCELL.focus();
	  } else {
	  	theCELL.style.color = 'red';
		theCELL.hsat.value = 0;
		}
		return false;
		}
 	document.frmTime.thsun.value = <%=strComSun%> parseFloat(document.frmTime.hsun1.value) + parseFloat(document.frmTime.hsunX1.value);
		if (parseFloat(document.frmTime.thsun.value) > 24)
		{
		alert("Total hours for Sunday exceeds 24 hours.");
		if (blnUse) {
			theCELL.style.color = 'red';
			theCELL.value = 0;
			theCELL.focus();
	  } else {
	  	theCELL.style.color = 'red';
		theCELL.value = 0;
		}
		return false;
		}
 document.frmTime.thtot.value = parseFloat(document.frmTime.thmon.value) +  parseFloat(document.frmTime.thtue.value) + parseFloat(document.frmTime.thwed.value) +  parseFloat(document.frmTime.ththu.value) + parseFloat(document.frmTime.thfri.value) + parseFloat(document.frmTime.thsat.value) + parseFloat(document.frmTime.thsun.value);
 
 tmpHRS = document.frmTime.thtot.value;
 document.frmTime.thrs.value = tmpHRS;
 
 document.frmTime.tmiles.value = parseFloat(document.frmTime.mileTOT.value) + parseFloat(document.frmTime.mileaTOT.value);
 return true;
}
function Compute2(theCELL, blnUse)
{
var tmpHRS2, tmpHRSX2;
tmpHRS2 = parseFloat(document.frmTime.hmon2.value) +  parseFloat(document.frmTime.htue2.value) + parseFloat(document.frmTime.hwed2.value) +  parseFloat(document.frmTime.hthu2.value) + parseFloat(document.frmTime.hfri2.value) +  parseFloat(document.frmTime.hsat2.value) + parseFloat(document.frmTime.hsun2.value);
 tmpHRSX2 = parseFloat(document.frmTime.hmonX2.value) +  parseFloat(document.frmTime.htueX2.value) + parseFloat(document.frmTime.hwedX2.value) +  parseFloat(document.frmTime.hthuX2.value) + parseFloat(document.frmTime.hfriX2.value) +  parseFloat(document.frmTime.hsatX2.value) + parseFloat(document.frmTime.hsunX2.value);
 document.frmTime.htot2.value = parseFloat(tmpHRS2) + parseFloat(tmpHRSX2);

	document.frmTime.mile2TOT.value = <%=strComMile2%> parseFloat(document.frmTime.txt2mile.value);
	document.frmTime.mile2aTOT.value = <%=strComaMile2%> parseFloat(document.frmTime.txt2amile.value);
	
	document.frmTime.thmon2.value = <%=strComMon2%> parseFloat(document.frmTime.hmon2.value) + parseFloat(document.frmTime.hmonX2.value) ;
	if (parseFloat(document.frmTime.thmon2.value) > 24)
		{
		alert("Total hours for Monday exceeds 24 hours.");
		if (blnUse) {
			theCELL.style.color = 'red';
			theCELL.value = 0;
			theCELL.focus();
	  } else {
	  theCELL = 'red';
		theCELL = 0;
		}
		//Compute();
		return false;
		}
		
	document.frmTime.thtue2.value = <%=strComTue2%> parseFloat(document.frmTime.htue2.value) + parseFloat(document.frmTime.htueX2.value);
	if (parseFloat(document.frmTime.thtue2.value) > 24)
		{
		alert("Total hours for Tuesday exceeds 24 hours.");
		if (blnUse) {
			theCELL.style.color = 'red';
			theCELL.value = 0;
			theCELL.focus();
	  } else {
	  	theCELL.style.color = 'red';
		theCELL.value = 0;
		}
		return false;
		} 
 	document.frmTime.thwed2.value = <%=strComWed2%> parseFloat(document.frmTime.hwed2.value) + parseFloat(document.frmTime.hwedX2.value);
	if (parseFloat(document.frmTime.thwed2.value) > 24)
		{
		alert("Total hours for Wednesday exceeds 24 hours.");
		if (blnUse) {
			theCELL.style.color = 'red';
			theCELL.value = 0;
			theCELL.focus();
	  } else {
	  	theCELL.style.color = 'red';
	  	theCELL.value = 0;
		}
		return false;
		}
 	document.frmTime.ththu2.value = <%=strComThu2%> parseFloat(document.frmTime.hthu2.value) + parseFloat(document.frmTime.hthuX2.value);
 	if (parseFloat(document.frmTime.ththu2.value) > 24)
		{
		alert("Total hours for Thursday exceeds 24 hours.");
		if (blnUse) {
			theCELL.style.color = 'red';
			theCELL.value = 0;
			theCELL.focus();
	  } else {
	  	theCELL.style.color = 'red';
		theCELL.value = 0;
		}
		return false;
		}
 	document.frmTime.thfri2.value = <%=strComFri2%> parseFloat(document.frmTime.hfri2.value) + parseFloat(document.frmTime.hfriX2.value);
 	if (parseFloat(document.frmTime.thfri2.value) > 24)
		{
		alert("Total hours for Friday exceeds 24 hours.");
		if (blnUse) {
			theCELL.style.color = 'red';
			theCELL.value = 0;
			theCELL.focus();
	  } else {
	  	theCELL.style.color = 'red';
		theCELL.value = 0;
		}
		return false;
		}
 	document.frmTime.thsat2.value = <%=strComSat2%> parseFloat(document.frmTime.hsat2.value) + parseFloat(document.frmTime.hsatX2.value);
 	if (parseFloat(document.frmTime.thsat2.value) > 24)
		{
		alert("Total hours for Saturday exceeds 24 hours.");
		if (blnUse) {
			theCELL.style.color = 'red';
			theCELL.value = 0;
			theCELL.focus();
	  } else {
	  	theCELL.style.color = 'red';
		theCELL.value = 0;
		}
		return false;
		}
 	document.frmTime.thsun2.value = <%=strComSun2%> parseFloat(document.frmTime.hsun2.value) + parseFloat(document.frmTime.hsunX2.value);
		if (parseFloat(document.frmTime.thsun2.value) > 24)
		{
		alert("Total hours for Sunday exceeds 24 hours.");
		if (blnUse) {
			theCELL.style.color = 'red';
			theCELL.value = 0;
			theCELL.focus();
	  } else {
	  	theCELL.style.color = 'red';
		theCELL.value = 0;
		}
		return false;
		}
 document.frmTime.thtot2.value = parseFloat(document.frmTime.thmon2.value) +  parseFloat(document.frmTime.thtue2.value) + parseFloat(document.frmTime.thwed2.value) +  parseFloat(document.frmTime.ththu2.value) + parseFloat(document.frmTime.thfri2.value) +  parseFloat(document.frmTime.thsat2.value) + parseFloat(document.frmTime.thsun2.value);

 tmpHRS = document.frmTime.thtot2.value;
 document.frmTime.thrs2.value = tmpHRS;
  document.frmTime.tmiles2.value = parseFloat(document.frmTime.mile2TOT.value) + parseFloat(document.frmTime.mile2aTOT.value);
 return true;
}
function delrow()
{
	var ans = window.confirm("Delete checked timesheet rows? Click OK to continue.");
	if (ans){
	document.frmTime.action = "delrow-2.asp";
	document.frmTime.submit();
	}
}
function prevWK(xxx)
{
	if (xxx = 1)
		{
			document.frmdate.action = "prevwk.asp?wk=1";
			document.frmdate.submit();
		}
	else
		{
			document.frmdate.action = "prevwk.asp?";
			document.frmdate.submit();
		}
}
function nextWK(xxx)
{
	if (xxx = 1)
		{
			document.frmdate.action = "nextwk.asp?wk=1";
			document.frmdate.submit();
		}
	else
		{
			document.frmdate.action = "nextwk.asp";
			document.frmdate.submit();
		}
}
function savrow()
{
	Compute(document, 0);
	document.frmTime.action = "savrow2.asp";
	document.frmTime.submit();
}
function savrow2()
{
	var ans = window.confirm("Timesheet will be submitted to the database. \n Week 1 = " + document.frmTime.thrs.value + "hrs. \n Week 2 = " + document.frmTime.thrs2.value + "hrs. \n Click Cancel to stop.");
	if (ans){
	Compute(document, 0);
	document.frmTime.action = "savrow2.asp";
	document.frmTime.submit();
	}
}
function CWorker()
{
	//var ans = window.confirm("Always Save your Work . Click Cancel to stop.");
	//if (ans){
		document.frmmain.action = "cWork.asp";
		document.frmmain.submit();
	//}
}
function hideEXT()
{
	if (document.frmTime.chkEXT.checked == true)
		{document.frmTime.hsunX1.style.visibility = 'visible';
		 document.frmTime.hmonX1.style.visibility = 'visible';
		 document.frmTime.htueX1.style.visibility = 'visible';
		 document.frmTime.hwedX1.style.visibility = 'visible';
		 document.frmTime.hthuX1.style.visibility = 'visible';
		 document.frmTime.hfriX1.style.visibility = 'visible';
		 document.frmTime.hsatX1.style.visibility = 'visible';}
	else
		{document.frmTime.hsunX1.style.visibility = 'hidden';
		 document.frmTime.hsunX1.value = 0;
		 document.frmTime.hmonX1.style.visibility = 'hidden';
		 document.frmTime.hmonX1.value = 0;
		 document.frmTime.htueX1.style.visibility = 'hidden';
		 document.frmTime.htueX1.value = 0;
		 document.frmTime.hwedX1.style.visibility = 'hidden';
		 document.frmTime.hwedX1.value = 0;
		 document.frmTime.hthuX1.style.visibility = 'hidden';
		 document.frmTime.hthuX1.value = 0;
		 document.frmTime.hfriX1.style.visibility = 'hidden';
		 document.frmTime.hfriX1.value = 0;
		 document.frmTime.hsatX1.style.visibility = 'hidden';
		 document.frmTime.hsatX1.value = 0;}
}
function hideEXT2()
{
	if (document.frmTime.chkEXT2.checked == true)
		{document.frmTime.hsunX2.style.visibility = 'visible';
		 document.frmTime.hmonX2.style.visibility = 'visible';
		 document.frmTime.htueX2.style.visibility = 'visible';
		 document.frmTime.hwedX2.style.visibility = 'visible';
		 document.frmTime.hthuX2.style.visibility = 'visible';
		 document.frmTime.hfriX2.style.visibility = 'visible';
		 document.frmTime.hsatX2.style.visibility = 'visible';}
	else
		{document.frmTime.hsunX2.style.visibility = 'hidden';
		 document.frmTime.hsunX2.value = 0;
		 document.frmTime.hmonX2.style.visibility = 'hidden';
		 document.frmTime.hmonX2.value = 0;
		 document.frmTime.htueX2.style.visibility = 'hidden';
		 document.frmTime.htueX2.value = 0;
		 document.frmTime.hwedX2.style.visibility = 'hidden';
		 document.frmTime.hwedX2.value = 0;
		 document.frmTime.hthuX2.style.visibility = 'hidden';
		 document.frmTime.hthuX2.value = 0;
		 document.frmTime.hfriX2.style.visibility = 'hidden';
		 document.frmTime.hfriX2.value = 0;
		 document.frmTime.hsatX2.style.visibility = 'hidden';
		 document.frmTime.hsatX2.value = 0;}
}
function GetEXT() {
	<%=strEXT%>
}
function GetEXT2() {
	<%=strEXT2%>
}
function chkNotes()
{
	if (document.frmTime.Mnotes1.value == "" && document.frmTime.chkEXT.checked == true)
		{ 
			alert("Extended hours requires notes.");
		}
	else
		{
			savrow();
		}
}
function chkNotes2()
{
	if (document.frmTime.Mnotes2.value == "" && document.frmTime.chkEXT2.checked == true)
		{ 
			alert("Extended hours requires notes.");
		}
	else
		{
			savrow();
		}
}
function PopMe(zzz)
	{
		newwindow = window.open(zzz,'name','height=400,width=300,scrollbars=1,directories=0,status=0,toolbar=0,resizable=0');
		if (window.focus) {newwindow.focus()}
	}
-->
</script>
<style>
INPUT, SELECT, TEXTBOX {
	font-size: 7.5pt;
	font-family: arial;
}
.HEADLINK, .HEADLINK A:link, .HEADLINK A:visited, .HEADLINK A:active {
color:	#666666;
text-decoration: none;
}
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
<body bgcolor='white'  LEFTMARGIN='0' TOPMARGIN='0' onload="document.frmmain.worker.focus(); RowDis(); RowDis2(); Compute(document.frmTime, 0); Compute2(document.frmTime, 0); hideEXT(); GetEXT(); GetEXT2(); hideEXT2();">

<table width='100%' height='100%' cellSpacing='0' cellPadding='0' border='0'>

<tr><td colspan='3'>
		<table align='left' border='0' width='100%' cellpadding='0' cellspacing='0'>
				<tr bgcolor='#040C8B' height="30px" valign="center">
					<td width='255px' style="background-color: #040C8B; font-family: Trebuchet MS; font-size: 10pt; font-weight: bold; letter-spacing: 1px; color: white;"
						style="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=0, startColorstr=#c7c9e5, endColorstr=#040C8B);">
						&nbsp;&nbsp;
						TIMESHEET<!-- <img src='images/Welcome.gif' border='0'> -->
						&nbsp;|&nbsp;
						<a href='admin2.asp' class="headlink"><font color='white'>Home<!-- <img src='images/Home.gif' border='0'> --></font></a>
						&nbsp;|&nbsp;
						<a href='default.asp' class="headlink"><font color='white'>Sign Out<!-- <img src='images/SignOut.gif' border='0'> --></font></a>
						&nbsp;|
					</td>
					<td align='left' style="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=0, startColorstr=#c7c9e5, endColorstr=#040C8B);">
					<font face='trebuchet MS' size='2' color='white'>&nbsp;&nbsp;<b>User:</b>&nbsp;<i><%=tmpUser%></i></font>
					</td>
					<td align='right' style="background-color: ##A4CADB; color: white; font-family: Trebuchet MS; font-size: 10pt; font-weight: bold; letter-spacing: 1px;"
						style="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=0, startColorstr=#c7c9e5, endColorstr=#040C8B);">
						&nbsp;|&nbsp;
						<a href='help.asp' target="_blank" class="headlink"><font color='white'>HELP&nbsp;&nbsp; <!-- <img src='images/Help.gif' border='0'> --></font></a></td>
				</tr>
				
		</table>
</td></tr>
<tr><td colspan='3' valign='top' style="height: 20px; background-color: #040C8B;">
	&nbsp;<!--<img border='0' src='images/Topbar.gif' style='width: 100%;' height='25px'> --></td></tr>
<tr>
	<td valign='top' style="height: 100%; width: 18px; background-color: #040C8B;">
	&nbsp;<!-- <img border='0' src='images/Leftbar.gif' style='width: 24px;' height='100%'> --></td>
			<td>
				<table border='0' width=100%>
					<tr><td>
						<table width="100%" border='0' cellspacing='0' cellpadding='0'>
					
		<tr><td width="330px" align='left' style="font-family: Trebuchet MS; font-size: 8pt; font-weight: bold;">
			<nobr><form method='post' name='frmmain'><span>PCSP Worker:&nbsp;</span>
				<select name='worker' style='font-size: 10px; height: 20px; width: 200px;' onchange='JavaScript:CWorker();'>					
					<%=strWork%>
				</select></form></nobr></td>
			<td align="right">
					<table border='0' cellspacing='0' cellpadding='0'>
						<form method='post' name='frmdate' action="GotoD8.asp">
						<tr><td align='center' colspan='8'><input type='submit' value='Go to Date'  class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" 
								style="width: 75px" onclick="javascript:GotoD8()">&nbsp;
								<input type='text' class='rightjust' name='specd8' size='10' maxlength='10'>
							</td></tr>
						<tr align='right'><td align='center'><input type='button' value='<' 
								title='previous 2 weeks' class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'"  style="width: 25px" onclick="javascript:prevWK('<%=sunDATE%>');">&nbsp;</td
							><td><table cellpadding="1" cellspacing="1" bgcolor='#040C8B'>
								<tr valign="center"><td style="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=0, startColorstr=#040C8B, endColorstr=#c7c9e5);">&nbsp;
									<span class="text"><font color='white'><b>From</b></font></span>
									<input type='text'  class='rightjust' name='1day' readOnly disabled size='10' value='<%=sunDATE%>'>
									<input type='hidden' name='2day' value='<%=sunDATE2%>'>
									<span class="text"><font color='white'><b>To</b></font></span>
									<input type='text' class='rightjust' size= '10' name='7day' readOnly  disabled value='<%=satDATE2%>'>
									&nbsp;</td></tr>
								</table>
							</td><td>&nbsp;
								<input type='button' value='>' title='next 2 weeks' class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" style='width: 25px' onclick="javascript:nextWK()" ></td></tr>
					</form></table></td></tr>

						</table><!-- from: (1200) table border='0' cellspacing='0' cellpadding='0'> -->
<center>
<form method='post' name='frmTime' action="response.asp"><span class="error"><%=Session("MSG")%></span>
<table border='0' align='left'>
	<tr>
		<td><input type='button' style='width: 100px;' align='left' value='Save Row' class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick="javascript:savrow();" <%=editme%> ></td>
		<% If SEssion("userid") = 2 Then %>
			<td><input type='button' style='width: 100px;' align='right' value='Delete Row(s)' class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick="javascript:delrow()" <%=editme%> ></td>
		<% End If %>
		<td colspan='3' ><input type='button' style='width: 100px;' align='left' value='Submit' class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick="javascript:savrow2()" <%=editme%>  ></td>
		</tr>
</table>
</td></tr></table>
<table align='center' border='0'>
	<tr align='center'>
	<td rowspan='2'>
		<table bgcolor='#d4d0c8'  border='0' cellpadding='0'>
			<tr class='title'>
				<td align='center' rowspan='2' width='100px'><font size='2'><b>Week 1</b></font></td>
				<td class='title' width= '75px' align='center'><font face='trebuchet MS' size='1'><u>  Total Hours </u></font></td
				><td class='title' width= '75px' align='center'><font face='trebuchet MS' size='1'><u> Total Mileage </u></font></td
				><td class='title' width= '75px' align='center'><font face='trebuchet MS' size='1'><u> PTO </u></font></td
			></tr>
			<tr>
				<td align='center'><input type='text' size='4' readOnly disabled name='thrs' class='rightjust'></td
				><td align='center'><input type='text' size='4' readOnly disabled name='tmiles' class='rightjust'></td
				><td align='center'><input type='text' size='4' name='txtPTO1' class='rightjust' <%=PTO1%> value='<%=tmpPTO1%>'></td
			></tr>
		</table>

	<table id='time_t' align='center' bgcolor='#d4d0c8' border='0' cellpadding='1' cellspacing='0' width='100%'>
<tr class='title' ><td width= '35px'>&nbsp;</td
><td class='title' width= '50px' align='center'><font face='trebuchet MS' size='1'><u>Consumer</u></td
><td class='title' width= '50px' align='center'><font face='trebuchet MS' size='1'><u>Sun<br><%=Z_MDYDate(sunDATE)%></u></font></td
><td class='title' width= '50px' align='center'><font face='trebuchet MS' size='1'><u>Mon<br><%=Z_MDYDate(monDATE)%></u></font></td
><td class='title' width= '50px' align='center'><font face='trebuchet MS' size='1'><u>Tue<br><%=Z_MDYDate(tueDATE)%></u></font></td
><td class='title' width= '50px' align='center'><font face='trebuchet MS' size='1'><u>Wed<br><%=Z_MDYDate(wedDATE)%></u></font></td
><td class='title' width= '50px' align='center'><font face='trebuchet MS' size='1'><u>Thu<br><%=Z_MDYDate(thuDATE)%></u></font></td
><td class='title' width= '50px' align='center'><font face='trebuchet MS' size='1'><u>Fri<br><%=Z_MDYDate(friDATE)%></u></font></td
><td class='title' width= '50px' align='center'><font face='trebuchet MS' size='1'><u>Sat<br><%=Z_MDYDate(satDATE)%></u></font></td
><td class='title' width= '50px' align='center'><font face='trebuchet MS' size='1'><u> Total </u></font></td
><td class='title' width= '50px' align='center'><font face='trebuchet MS' size='1'><u> Mileage </u></font></td
><td class='title' width= '50px' align='center'><font face='trebuchet MS' size='1'><u>Admin Mileage</u></font></td
><td class='title' width= '75px' align='center'><a href="" onclick="PopMe('actcodes.asp');"><font face='trebuchet MS' size='1'><u>Activity Code</u></font></a></td
><td class='title' width= '75px' align='center'><font face='trebuchet MS' size='1'><u>Notes</u></font></td
><td class='title' width= '45px' align='center'><font face='trebuchet MS' size='1'><nobr><u>Phone Num</u></font></td
><td class='title' width= '45px' align='center'><font face='trebuchet MS' size='1'><u>Medicaid/Private Pay/VA</u></font></td
><td class='title'>&nbsp;</td
></tr>
<%=strTableScript%>
<tr>
<td align='center'>&nbsp;<input type='hidden' name='tmpID'
><td align='center'><select name='hdept1' <%=DisNew%> class='textbox' style='width:97px; font-size: 8pt;'>
					<option></option>
					<%=strdept%>
					</select><br>
					
	<font face='trebuchet MS' size='1'>Extd Hrs.</font><input type='checkbox' size= '5'  <%=ExtX%> <%=DisNew%> name='chkEXT' onclick='hideEXT();Compute();'>
</td
><td align='center' >
	<input type='text' size= '3' class='rightjust' name='hsun1'  <%=DisNew%> value='<%=sunh%>'  onblur="JavaScript:numericVal(this.value, this); Compute(this, 0);" style='font-size: 8pt;'><br>
	<input type='text' size= '3' class='rightjust' name='hsunX1'  value='<%=sunhX%>'  onblur="JavaScript:numericVal(this.value, this); Compute(this, 0)" style='font-size: 8pt;'>
</td
><td align='center' >
	<input type='text' size= '3' class='rightjust' name='hmon1'  <%=DisNew%> value='<%=monh%>'  onblur="JavaScript:numericVal(this.value, this); Compute(this, 0);" style='font-size: 8pt;'><br>
	<input type='text' size= '3' class='rightjust' name='hmonX1'  value='<%=monhX%>'  onblur="JavaScript:numericVal(this.value, this); Compute(this, 0);" style='font-size: 8pt;'>
</td
><td align='center' >
	<input type='text' size= '3' class='rightjust' name='htue1'  <%=DisNew%> value='<%=tueh%>'  onblur="JavaScript:numericVal(this.value, this); Compute(this, 0);" style='font-size: 8pt;'><br>
	<input type='text' size= '3' class='rightjust' name='htueX1'  value='<%=tuehX%>'  onblur="JavaScript:numericVal(this.value, this); Compute(this, 0);" style='font-size: 8pt;'>
</td
><td align='center' >
	<input type='text' size= '3' class='rightjust' name='hwed1'  <%=DisNew%> value='<%=wedh%>'  onblur="JavaScript:numericVal(this.value, this); Compute(this, 0);" style='font-size: 8pt;'><br>
	<input type='text' size= '3' class='rightjust' name='hwedX1' value='<%=wedhX%>'  onblur="JavaScript:numericVal(this.value, this); Compute(this, 0);" style='font-size: 8pt;'>
</td
><td align='center' >
	<input type='text' size= '3' class='rightjust' name='hthu1'  <%=DisNew%> value='<%=thuh%>'  onblur="JavaScript:numericVal(this.value, this); Compute(this, 0);" style='font-size: 8pt;'><br>
	<input type='text' size= '3' class='rightjust' name='hthuX1'  value='<%=thuhX%>'  onblur="JavaScript:numericVal(this.value, this); Compute(this, 0);" style='font-size: 8pt;'>
</td
><td align='center' >
	<input type='text' size= '3' class='rightjust' name='hfri1'  <%=DisNew%> value='<%=frih%>'  onblur="JavaScript:numericVal(this.value, this); Compute(this, 0);" style='font-size: 8pt;'><br>
	<input type='text' size= '3' class='rightjust' name='hfriX1'  value='<%=frihX%>'  onblur="JavaScript:numericVal(this.value, this); Compute(this, 0);" style='font-size: 8pt;'>
</td
><td align='center' >
	<input type='text' size= '3' class='rightjust' name='hsat1'  <%=DisNew%> value='<%=sath%>'  onblur="JavaScript:numericVal(this.value, this); Compute(this, 0);" style='font-size: 8pt;'><br>
	<input type='text' size= '3' class='rightjust' name='hsatX1'  value='<%=sathX%>'  onblur="JavaScript:numericVal(this.value, this); Compute(this, 0);" style='font-size: 8pt;'>
</td
><td align='center' >
	<input type='text' size= '3' name='htot1'  readOnly  class='rightjust' style='font-size: 8pt;'>
</td
><td align='center' >
	<% if Ucase(DisNew) = "DISABLED" Then 
		mileageAllow = "" 
	Else
		if AllowMileage(Session("idemp")) Then
			
			mileageAllow = ""
		Else
			
			mileageAllow = "DISABLED" 
		End If
	end if
	%>
	<input type='text' size= '3' name='txtmile' class='rightjust' value='<%=mile1%>' <%=DisNew%> <%=mileageAllow%> style='font-size: 8pt;'  onblur="JavaScript:numericVal(this.value, this); Compute(this, 0);" style='font-size: 8pt;'
>
<td align='center' >
	<input type='text' size= '3' name='txtamile' class='rightjust' value='<%=amile1%>' <%=DisNew%> style='font-size: 8pt;'  onblur="JavaScript:numericVal(this.value, this); Compute(this, 0);" style='font-size: 8pt;'
>
<td align='center' >
	<textarea rows="2" style='font-size: 8pt;' <%=DisNew%> name="Mnotes1" cols="15"  ><%=notesh%></textarea>
</td>
<td align='center' >
	<textarea rows="2" style='font-size: 8pt;' <%=DisNew%> name="com1" cols="15"  ><%=com1%></textarea>
</td>
<td align='center'>
	<input style='font-size: 8pt;' type='text' <%=DisNew%> name='fon1' size= '12' value="<%=fon1%>">
</td>
<td>&nbsp;</td>
<td>&nbsp;</td>
</tr>

	
<tr bgcolor='#b6b3ae'><td border='1' style="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=0, startColorstr=#d4d0c8, endColorstr=#b6b3ae);">&nbsp;</td
><td class='title' align='right' style="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=0, startColorstr=#d4d0c8, endColorstr=#b6b3ae);"><p align='center'><font face='trebuchet MS' size='2'><b>Totals</b></font></p></td
><td align='center' style="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=0, startColorstr=#d4d0c8, endColorstr=#b6b3ae);"><input type='text' size='3' readOnly disabled name='thsun' class='rightjust' style='font-size: 8pt;'></td
><td align='center' style="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=0, startColorstr=#d4d0c8, endColorstr=#b6b3ae);"><input type='text' size='3' readOnly disabled name='thmon' class='rightjust' style='font-size: 8pt;'></td
><td align='center' style="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=0, startColorstr=#d4d0c8, endColorstr=#b6b3ae);"><input type='text' size='3' readOnly disabled name='thtue' class='rightjust' style='font-size: 8pt;'></td
><td align='center' style="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=0, startColorstr=#d4d0c8, endColorstr=#b6b3ae);"><input type='text' size='3' readOnly disabled name='thwed' class='rightjust' style='font-size: 8pt;'></td
><td align='center' style="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=0, startColorstr=#d4d0c8, endColorstr=#b6b3ae);"><input type='text' size='3' readOnly disabled name='ththu' class='rightjust' style='font-size: 8pt;'></td
><td align='center' style="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=0, startColorstr=#d4d0c8, endColorstr=#b6b3ae);"><input type='text' size='3' readOnly disabled name='thfri' class='rightjust' style='font-size: 8pt;'></td
><td align='center' style="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=0, startColorstr=#d4d0c8, endColorstr=#b6b3ae);"><input type='text' size='3' readOnly disabled name='thsat' class='rightjust' style='font-size: 8pt;' ></td
><td align='center' style="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=0, startColorstr=#d4d0c8, endColorstr=#b6b3ae);"><input type='text' size='3' name='thtot' readOnly disabled class='rightjust' style='font-size: 8pt;'></td
><td align='center' style="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=0, startColorstr=#C4B464, endColorstr=#9a882f);"><input type='text' size='3' name='mileTOT' readOnly disabled class='rightjust' style='font-size: 8pt;'></td
><td align='center' style="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=0, startColorstr=#C4B464, endColorstr=#9a882f);"><input type='text' size='3' name='mileaTOT' readOnly disabled class='rightjust' style='font-size: 8pt;'></td
><td style="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=0, startColorstr=#d4d0c8, endColorstr=#b6b3ae);">&nbsp;</td
><td style="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=0, startColorstr=#d4d0c8, endColorstr=#b6b3ae);">&nbsp;</td
><td style="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=0, startColorstr=#d4d0c8, endColorstr=#b6b3ae);">&nbsp;</td
><td style="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=0, startColorstr=#d4d0c8, endColorstr=#b6b3ae);">&nbsp;</td

></tr>
</table>
<br>
<%'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''WEEK 2
%>
<table align='center' border='0' cellpadding='0'>
	<tr align='center'>
	<td rowspan='2'>
		<table bgcolor='#C4B464'  border='0' cellpadding='0'>
			<tr class='title'>
				<td align='center' rowspan='2' width='100px'><font size='2'><b>Week 2</b></font></td>
				<td class='title' width= '75px' align='center'><font face='trebuchet MS' size='1'><u> Total Hours </u></font></td
				><td class='title' width= '75px' align='center'><font face='trebuchet MS' size='1'><u> Total Mileage </u></font></td
				><td class='title' width= '75px' align='center'><font face='trebuchet MS' size='1'><u> PTO </u></font></td
			></tr>
			<tr>
				<td align='center'><input type='text' size='4' readOnly disabled name='thrs2' class='rightjust'></td
				><td align='center'><input type='text' size='4' readOnly disabled name='tmiles2' class='rightjust'></td
				><td align='center'><input type='text' size='4' name='txtPTO2' class='rightjust' <%=PTO2%> value='<%=tmpPTO2%>'></td
			></tr>
	</table
></table>
	<table id='time_t2' align='center' bgcolor='#C4B464' border='0' cellpadding='1' cellspacing='0' width='100%'>
<tr class='title' ><td width= '35px'>&nbsp;</td
><td class='title' width= '50px' align='center'><font face='trebuchet MS' size='1'><u>Consumer</u></td
><td class='title' width= '50px' align='center'><font face='trebuchet MS' size='1'><u>Sun<br><%=Z_MDYDate(sunDATE2)%></u></font></td
><td class='title' width= '50px' align='center'><font face='trebuchet MS' size='1'><u>Mon<br><%=Z_MDYDate(monDATE2)%></u></font></td
><td class='title' width= '50px' align='center'><font face='trebuchet MS' size='1'><u>Tue<br><%=Z_MDYDate(tueDATE2)%></u></font></td
><td class='title' width= '50px' align='center'><font face='trebuchet MS' size='1'><u>Wed<br><%=Z_MDYDate(wedDATE2)%></u></font></td
><td class='title' width= '50px' align='center'><font face='trebuchet MS' size='1'><u>Thu<br><%=Z_MDYDate(thuDATE2)%></u></font></td
><td class='title' width= '50px' align='center'><font face='trebuchet MS' size='1'><u>Fri<br><%=Z_MDYDate(friDATE2)%></u></font></td
><td class='title' width= '50px' align='center'><font face='trebuchet MS' size='1'><u>Sat<br><%=Z_MDYDate(satDATE2)%></u></font></td
><td class='title' width= '50px' align='center'><font face='trebuchet MS' size='1'><u> Total </u></font></td
><td class='title' width= '50px' align='center'><font face='trebuchet MS' size='1'><u> Mileage </u></font></td
><td class='title' width= '50px' align='center'><font face='trebuchet MS' size='1'><u>Admin Mileage</u></font></td
><td class='title' width= '75px' align='center'><a href="" onclick="PopMe('actcodes.asp');"><font face='trebuchet MS' size='1'><u>Activity Code</u></font></a></td
><td class='title' width= '75px' align='center'><font face='trebuchet MS' size='1'><u>Notes</u></font></td
><td class='title' width= '45px' align='center'><font face='trebuchet MS' size='1'><nobr><u>Phone Num.</u></font></td
><td class='title' width= '45px' align='center'><font face='trebuchet MS' size='1'><u>Medicaid/Private Pay/VA</u></font></td
><td class='title'>&nbsp;</td
></tr>
<%=strTableScript2%>
<tr>
<td align='center'>&nbsp;<input type='hidden' name='tmpID2'
><td align='center'><select name='hdept2' <%=DisNew2%> class='textbox' style='width:97px; font-size: 8pt;'>
					<option></option>
					<%=strdept2%>
					</select><br>
					<font face='trebuchet MS' size='1'>Extd Hrs.</font><input type='checkbox' size= '5'  <%=ExtX2%> <%=DisNew2%> name='chkEXT2' onclick='hideEXT2();Compute2();'>
					</td
><td align='center' >
	<input type='text' size= '3' class='rightjust' style='font-size: 8pt;' name='hsun2'  <%=DisNew2%> value='<%=sunh2%>'  onblur="JavaScript:numericVal(this.value, this); Compute2(this, 0);"><br>
	<input type='text' size= '3' class='rightjust' style='font-size: 8pt;' name='hsunX2'  value='<%=sunhX2%>'  onblur="JavaScript:numericVal(this.value, this); Compute2(this, 0);">
</td
><td align='center' >
	<input type='text' size= '3' class='rightjust' style='font-size: 8pt;' name='hmon2'  <%=DisNew2%> value='<%=monh2%>'  onblur="JavaScript:numericVal(this.value, this); Compute2(this, 0);"><br>
	<input type='text' size= '3' class='rightjust' style='font-size: 8pt;' name='hmonX2'  value='<%=monhX2%>'  onblur="JavaScript:numericVal(this.value, this); Compute2(this, 0);">
	</td
><td align='center' >
	<input type='text' size= '3' class='rightjust' style='font-size: 8pt;' name='htue2'  <%=DisNew2%> value='<%=tueh2%>'  onblur="JavaScript:numericVal(this.value, this); Compute2(this, 0);"><br>
	<input type='text' size= '3' class='rightjust' style='font-size: 8pt;' name='htueX2'  value='<%=tuehX2%>'  onblur="JavaScript:numericVal(this.value, this); Compute2(this, 0);">
	</td
><td align='center' >
	<input type='text' size= '3' class='rightjust' style='font-size: 8pt;' name='hwed2'  <%=DisNew2%> value='<%=wedh2%>'  onblur="JavaScript:numericVal(this.value, this); Compute2(this, 0);"><br>
	<input type='text' size= '3' class='rightjust' style='font-size: 8pt;' name='hwedX2'  value='<%=wedhX2%>'  onblur="JavaScript:numericVal(this.value, this); Compute2(this, 0);">
	</td
><td align='center' >
	<input type='text' size= '3' class='rightjust' style='font-size: 8pt;' name='hthu2'  <%=DisNew2%> value='<%=thuh2%>'  onblur="JavaScript:numericVal(this.value, this); Compute2(this, 0);"><br>
	<input type='text' size= '3' class='rightjust' style='font-size: 8pt;' name='hthuX2'  value='<%=thuhX2%>'  onblur="JavaScript:numericVal(this.value, this); Compute2(this, 0);">
	</td
><td align='center' >
	<input type='text' size= '3' class='rightjust' style='font-size: 8pt;' name='hfri2'  <%=DisNew2%> value='<%=frih2%>'  onblur="JavaScript:numericVal(this.value, this); Compute2(this, 0);"><br>
	<input type='text' size= '3' class='rightjust' style='font-size: 8pt;' name='hfriX2'  value='<%=frihX2%>' onblur="JavaScript:numericVal(this.value, this); Compute2(this, 0);">
	</td
><td align='center' >
	<input type='text' size= '3' class='rightjust' style='font-size: 8pt;' name='hsat2'  <%=DisNew2%> value='<%=sath2%>'  onblur="JavaScript:numericVal(this.value, this); Compute2(this, 0);"><br>
	<input type='text' size= '3' class='rightjust' style='font-size: 8pt;' name='hsatX2'  value='<%=sathX2%>'  onblur="JavaScript:numericVal(this.value, this); Compute2(this, 0);">
	</td
><td align='center' >
	<input type='text' size= '3' name='htot2'  readOnly class='rightjust' style='font-size: 8pt;'>
	</td
><td align='center' >
	<% if Ucase(DisNew2) = "DISABLED" Then 
		mileageAllow = ""
		Else
		if AllowMileage(Session("idemp")) Then
			
			mileageAllow = ""
		Else
			
			mileageAllow = "DISABLED" 
		End If
	end if 
		%>
	<input type='text' size= '3' name='txt2mile' class='rightjust' value='<%=mile2%>' <%=DisNew2%> <%=mileageAllow%> style='font-size: 8pt;' onblur="JavaScript:numericVal(this.value, this); Compute2(this, 0);" style='font-size: 8pt;'
>
<td align='center' >
	<input type='text' size= '3' name='txt2amile' class='rightjust' value='<%=amile2%>' <%=DisNew2%> style='font-size: 8pt;' onblur="JavaScript:numericVal(this.value, this); Compute2(this, 0);" style='font-size: 8pt;'
>
<td align='center' >
	<textarea rows="2" name="Mnotes2" cols="15" <%=DisNew2%> style='font-size: 8pt;' ><%=notesh2%></textarea>
</td>
<td align='center' >
	<textarea rows="2" style='font-size: 8pt;' <%=DisNew2%> name="com2" cols="15"  ><%=com2%></textarea>
</td>
<td align='center'>
	<input style='font-size: 8pt;' type='text' <%=DisNew2%> name='fon2' size= '12' value="<%=fon2%>">
</td>
</tr>
	
<tr bgcolor='#9a882f'><td style="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=0, startColorstr=#C4B464, endColorstr=#9a882f);">&nbsp;</td
><td class='title' align='right' style="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=0, startColorstr=#C4B464, endColorstr=#9a882f);"><p align='center'><font face='trebuchet MS' size='2'><b>Totals</b></font></p></td
><td align='center' style="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=0, startColorstr=#C4B464, endColorstr=#9a882f);"><input type='text' size='3' readOnly disabled name='thsun2' class='rightjust' style='font-size: 8pt;'></td
><td align='center' style="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=0, startColorstr=#C4B464, endColorstr=#9a882f);"><input type='text' size='3' readOnly disabled name='thmon2' class='rightjust' style='font-size: 8pt;'></td
><td align='center' style="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=0, startColorstr=#C4B464, endColorstr=#9a882f);"><input type='text' size='3' readOnly disabled name='thtue2' class='rightjust' style='font-size: 8pt;'></td
><td align='center' style="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=0, startColorstr=#C4B464, endColorstr=#9a882f);"><input type='text' size='3' readOnly disabled name='thwed2' class='rightjust' style='font-size: 8pt;'></td
><td align='center' style="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=0, startColorstr=#C4B464, endColorstr=#9a882f);"><input type='text' size='3' readOnly disabled name='ththu2' class='rightjust' style='font-size: 8pt;'></td
><td align='center' style="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=0, startColorstr=#C4B464, endColorstr=#9a882f);"><input type='text' size='3' readOnly disabled name='thfri2' class='rightjust' style='font-size: 8pt;'></td
><td align='center' style="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=0, startColorstr=#C4B464, endColorstr=#9a882f);"><input type='text' size='3' readOnly disabled name='thsat2' class='rightjust' style='font-size: 8pt;'></td
><td align='center' style="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=0, startColorstr=#C4B464, endColorstr=#9a882f);"><input type='text' size='3' name='thtot2' readOnly disabled class='rightjust' style='font-size: 8pt;'></td
><td align='center' style="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=0, startColorstr=#C4B464, endColorstr=#9a882f);"><input type='text' size='3' name='mile2TOT' readOnly disabled class='rightjust' style='font-size: 8pt;'></td
><td align='center' style="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=0, startColorstr=#C4B464, endColorstr=#9a882f);"><input type='text' size='3' name='mile2aTOT' readOnly disabled class='rightjust' style='font-size: 8pt;'></td
><td style="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=0, startColorstr=#C4B464, endColorstr=#9a882f);">&nbsp;</td
><td style="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=0, startColorstr=#C4B464, endColorstr=#9a882f);">&nbsp;</td
><td style="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=0, startColorstr=#C4B464, endColorstr=#9a882f);">&nbsp;</td
><td style="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=0, startColorstr=#C4B464, endColorstr=#9a882f);">&nbsp;</td
></tr>
<tr bgcolor='white'><td colspan='17'>&nbsp;</td></tr>
</table>


<input type='hidden' name='count' value='<%=lngI%>'>
<input type='hidden' name='count2' value='<%=lngI2%>'>
<input type='hidden' name='eid' value='<%=Session("idemp")%>'>
<input type='hidden' name='ename' value='<%=Session("name")%>'>
<input type='hidden' name='User' value='<%=tmpUser%>'>
<input type='hidden' name='Paychk' value='<%=Paychk%>'>
<input type='hidden' name='Paychk2' value='<%=Paychk2%>'>
<input type='hidden' name='Medchk' value='<%=Medchk%>'>
<input type='hidden' name='Medchk2' value='<%=Medchk2%>'>
<input type='hidden' name='2day' value='<%=sunDATE2%>'>
<input type='hidden' name='1day' value='<%=sunDATE%>'>

						</td></tr>
					
				</table>
			</td>
			<td align='right' style='height: 100%; width: 18px; background-color: #040C8B;'>
				&nbsp;<!-- <img border='0' src='images/Rightbar.gif' style='width: 25px;' height='100%'> --></td>	
		</tr>
	</form> 
</td></tr>
<tr><td colspan='3' valign='top' style='height: 20px; background-color: #040C8B;'>
		&nbsp;<!-- <img border='0' src='images/Butbar.gif' style='width: 100%;' height='25px'> --></td></tr>
<tr vAlign=center
				><td bgcolor='#040C8B' colspan='3' align='center' 
					style="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=0, startColorstr=#040C8B, endColorstr=#c7c9e5);
					color: #FFFFFF; font-size: 10pt; font-family: Trebuchet MS; font-weight: bold; letter-spacing: 5px;">
					Lutheran&nbsp;&nbsp;Social&nbsp;&nbsp;Services&nbsp;&nbsp;of&nbsp;&nbsp;New&nbsp;&nbsp;England
					<!-- <img alt="Lutheran Social Services of New England" src="images/BotBanner.Gif" style='width: 100%;' border=0> --></td
			></tr>
</table>
</body>
</html>
<%
If Session("MSG") <> "" Then
tmpMSG = Replace(Session("MSG"), "<br>", " \n ")
%>
<script><!--
	alert("<%=tmpMSG%>");
--></script>
<%
End If
Session("MSG") = ""
%>
