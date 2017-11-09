<%Language=VBScript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_security.asp" -->
<%
If Request("ctrl") = 1 Then
	If Not IsNumeric(Request("txtcode")) Then
		Session("MSG") = "ERROR: Activity Code needs to be a 2 digit number."
		response.redirect "Activity.asp"
	Else
		Set rsAct = Server.CreateObject("ADODB.RecordSet")
		sqlAct = "SELECT * FROM Activity_T WHERE code = " & Request("txtcode")
		rsAct.Open sqlAct, g_strCONN, 1, 3
		If Not rsAct.EOF Then
			Session("MSG") = "ERROR: Activity Code already exists in the database."
			rsAct.Close
			Set rsAct = Nothing
			response.redirect "Activity.asp"
		Else
			rsAct.AddNew
			rsAct("code") = Request("txtcode")
			rsACt("desc") = Request("txtdesc")
			rsAct.Update
		End If
		rsAct.Close
		Set rsAct = Nothing
		Session("MSG") = "Activity Saved."
		response.redirect "Activity.asp"
	End If
ElseIf Request("ctrl") = 2 Then
	Set tblSite = Server.CreateObject("ADODB.RecordSet")
	sqlSite = "SELECT * FROM activity_T "
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
						tmpHol = tblSite("code") 
						tblSite.Delete
						tblSite.Update
						Session("MSG") = Session("MSG") & "Acitvity deleted.<br>"
						Set fso = CreateObject("Scripting.FileSystemObject")
						Set ALog = fso.OpenTextFile(AdminLog, 8, True)
						Alog.WriteLine Now & ":: Activity DELETED ( " & tmpHol & " ) BY: UID: " & Session("UserID") & vbCrLf
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
Response.Redirect "activity.asp"
%>