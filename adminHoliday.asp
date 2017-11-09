<%Language=VBScript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_security.asp" -->
<%
If Request("ctrl") = 1 Then
	If Request("selMonth") <> 0 And Request("selDay") <> 0 Then
		Set rsReg = Server.CreateObject("ADODB.RecordSet")
		sqlReg = "SELECT * FROM RegHoliday_T WHERE month = " & Request("selMonth") & " AND day = " & Request("selDay")
		rsReg.Open sqlReg, g_strCONN, 1, 3
		If rsReg.EOF Then
			rsReg.AddNew
			rsReg("month") = Request("selMonth")
			rsReg("day") = Request("selDay")
			rsReg.Update
			Set fso = CreateObject("Scripting.FileSystemObject")
			Set ALog = fso.OpenTextFile(AdminLog, 8, True)
			Alog.WriteLine Now & ":: Holiday Added ( " & Request("selMonth1") & " / " & Request("selDay1") & " ) BY: UID: " & Session("UserID") & vbCrLf
			Set Alog = Nothing
			Set fso = Nothing 
		Else
			Session("MSG") = Session("MSG") & "Holiday Date " & Request("selMonth") & " / " & Request("selDay") & " already exists.<br>"
		End If
		rsReg.Close
		Set rsReg = Nothing
	Else
		If Not (Request("selMonth") = 0 And Request("selDay") = 0) Then
			Session("MSG") = Session("MSG") & "Month or Date cannot be blank.<br>"
		End If
	End If
	If Request("selMonth1") <> 0 And Request("selDay1") <> 0 And Request("txtYear") <> "" Then
		If IsNumeric(Request("txtYear")) Then
			Set rsReg = Server.CreateObject("ADODB.RecordSet")
			sqlReg = "SELECT * FROM SpecHoliday_T WHERE month = " & Request("selMonth1") & " AND day = " & Request("selDay1") & _
				" AND year = " & Request("txtYear")
			rsReg.Open sqlReg, g_strCONN, 1, 3
			If rsReg.EOF Then
				rsReg.AddNew
				rsReg("month") = Request("selMonth1")
				rsReg("day") = Request("selDay1")
				rsReg("year") = Z_CZero(Request("txtYear"))
				rsReg.Update
				Set fso = CreateObject("Scripting.FileSystemObject")
				Set ALog = fso.OpenTextFile(AdminLog, 8, True)
				Alog.WriteLine Now & ":: Holiday Added ( " & Request("selMonth1") & " / " & Request("selDay1") & " / " & Request("txtYear") & " ) BY: UID: " & Session("UserID") & vbCrLf
				Set Alog = Nothing
				Set fso = Nothing 
			Else
				Session("MSG") = Session("MSG") & "Holiday Date " & Request("selMonth1") & " / " & Request("selDay1") & " / " & Request("txtYear") & " already exists.<br>"
			End If
			rsReg.Close
			Set rsReq = Nothing
		Else
			Session("MSG") = Session("MSG") & "Please input a valid year.<br>"
		End If
	Else
		If Not (Request("selMonth1") <> 0 And Request("selDay1") <> 0 And Request("txtYear") <> "") Then
			Session("MSG") = Session("MSG") & "Month or Date or Year cannot be blank.<br>"
		End If
	End If
ElseIf Request("ctrl") = 2 Then
	Set tblSite = Server.CreateObject("ADODB.RecordSet")
	sqlSite = "SELECT * FROM RegHoliday_T "
	tblSite.Open sqlSite, g_strCONN, 1, 3
	If Not tblSite.EOF Then
		If Request("ctr1") <> "" Then 
			'response.write Request("ctr1")
			ctr = Request("ctr1")
			For i = 0 to ctr 
			'response.write Request("chkReg" & i)
				tmpctr = Request("chkReg" & i)
				If tmpctr <> "" Then
					strTmp = "index= " & tmpctr 
					tblSite.Movefirst
					tblSite.Find(strTmp)
					'response.write strTmp
					If Not tblSite.EOF Then
						tmpHol = tblSite("month") & " / " & tblSite("day")
						tblSite.Delete
						tblSite.Update
						Session("MSG") = Session("MSG") & "Holiday deleted.<br>"
						Set fso = CreateObject("Scripting.FileSystemObject")
						Set ALog = fso.OpenTextFile(AdminLog, 8, True)
						Alog.WriteLine Now & ":: Holiday DELETED ( " & tmpHol & " ) BY: UID: " & Session("UserID") & vbCrLf
						Set Alog = Nothing
						Set fso = Nothing 
					End If
				End If
			Next
		End If 
	End If
	tblSite.Close
	Set tblSite = Nothing
	Set tblSite = Server.CreateObject("ADODB.RecordSet")
	sqlSite = "SELECT * FROM SpecHoliday_T "
	tblSite.Open sqlSite, g_strCONN, 1, 3
	If Not tblSite.EOF Then
		If Request("ctr2") <> "" Then 
			ctr = Request("ctr2")
			For i = 0 to ctr 
				tmpctr = Request("chkSpec" & i)
				If tmpctr <> "" Then
					strTmp = "index= " & tmpctr 
					tblSite.Movefirst
					tblSite.Find(strTmp)
					If Not tblSite.EOF Then
						tmpHol = tblSite("month") & " / " & tblSite("day") & " / " & tblSite("year")
						tblSite.Delete
						tblSite.Update
						Session("MSG") = Session("MSG") & "Holiday deleted.<br>"
						Set fso = CreateObject("Scripting.FileSystemObject")
						Set ALog = fso.OpenTextFile(AdminLog, 8, True)
						Alog.WriteLine Now & ":: Holiday DELETED ( " & tmpHol & " ) BY: UID: " & Session("UserID") & vbCrLf
						Set Alog = Nothing
						Set fso = Nothing 
					End If
				End If
			Next
		End If 
	End If
	tblSite.Close
	Set tblSite = Nothing
End If
Response.Redirect "holiday.asp"
%>