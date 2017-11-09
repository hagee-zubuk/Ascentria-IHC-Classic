<%@Language=VBScript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<%

DIM		lngI, tblEMP, strSQL, strTableScript, tmpID, ctrI, tmpctr
'If Request("paychk") = "" AND Request("medchk") = "" Then
	ctrI = Request("count")
		For i = 0 to ctrI 
			tmpctr = Request("chk" & i)
			If tmpctr <> "" Then
				Set tblEMP = Server.CreateObject("ADODB.Recordset")
				strSQL = "SELECT * FROM [tsheets_t] where procmed is null and procpriv is null and procpay is null and date = '" & Request("1day") & "' "
				tblEMP.Open strSQL, g_strCONN, 1, 3
				tblEMP.movefirst
				strTmp1 = "ID=" & tmpctr 
				tblEMP.Find(strTmp1)
				If Not tblEMP.EOF Then
					extfound = false
					If Z_FixNull(tblEMP("procMed")) = "" And Z_FixNull(tblEMP("procPay")) = "" And Z_FixNull(tblEMP("procPriv")) = "" And Z_FixNull(tblEMP("procMile")) = "" Then
						tmpWor = tblEMP("emp_id")
						tmpCli = tblEMP("client")
						tmpDate = tblEMP("date")
						tmpTS = tblEMP("timestamp") 
						Set rsEXT = Server.CreateObject("ADODB.RecordSet")
						'sqlEXT = "DELETE * FROM tsheets_t WHERE ID = " & tmpctr + 1
						sqlEXT = "SELECT * FROM tsheets_T WHERE emp_ID = '" & tmpWor & "' AND client = '" & tmpCli & "' AND date = '" & tmpDate & "' " & _
							"AND timestamp = '" & tmpTS & "' AND EXT = 1 AND ID = " & tmpctr + 1
						rsEXT.Open sqlEXT, g_strCONN, 1, 3
						If rsEXT.EOF Then 
							Set rsEXT2 = Server.CreateObject("ADODB.RecordSet")
							sqlEXT2 = "SELECT * FROM tsheets_t WHERE ID = " & tmpctr + 1
							rsEXT2.Open sqlEXT2, g_strCONN, 1, 3
							If rsEXT2("emp_ID") = tmpWor And rsEXT2("client") = tmpCli And rsEXT2("date") = tmpDate Then
								Set fso = CreateObject("Scripting.FileSystemObject")
								Set ALog = fso.OpenTextFile(AdminLog, 8, True)
								Alog.WriteLine Now & ":: Inital EXT SQL NOT FOUND -- (" & tmpTS & " <> " & rsEXT2("timestamp") & ")Timesheet ID: " &  tmpctr + 1 & " was deleted . (week 1) -- UID: " & Session("UserID") & vbCrLf
								Set Alog = Nothing
								Set fso = Nothing 
								rsEXT2.Delete 1
								extfound = true
							Else
								Set fso = CreateObject("Scripting.FileSystemObject")
								Set ALog = fso.OpenTextFile(AdminLog, 8, True)
								Alog.WriteLine Now & ":: Extended timesheet not found for " & tmpctr & " (week 1) -- UID: " & Session("UserID") & vbCrLf
								Set Alog = Nothing
								Set fso = Nothing 
								extfound = false
							End If
							Set rsEXT2 = Nothing
						Else
							rsEXT.Delete 1
							extfound = true
						End If
						Set rsEXT = Nothing
						if extfound then 
							tblEmp.DELETE 1
						Else
							Session("MSG") = Session("MSG") & "<br>Unable to delete checked timesheet in week 1. Please make a screen shot of the row in the timesheet and email to patrick@zubuk.com"
						end if
						Set fso = CreateObject("Scripting.FileSystemObject")
						Set ALog = fso.OpenTextFile(AdminLog, 8, True)
						Alog.WriteLine Now & ":: Timesheet ID: " &  tmpctr & " was deleted. (week 1) -- UID: " & Session("UserID") & vbCrLf
						Set Alog = Nothing
						Set fso = Nothing 
					Else
						Session("MSG") = Session("MSG") & "<br>Unable to delete checked timesheet in week 1. Timesheet may have been billed already."
					End If
				End If
				tblEmp.Close
				Set tblEMP = Nothing
			End If
		Next 
'End If
''''''2 week
'If Request("paychk2") = "" AND Request("medchk2") = "" Then
		ctrI = Request("count2")
		For i = 0 to ctrI - 1
			tmpctr = Request("chkS" & i)
			If tmpctr <> "" Then
				Set tblEMP = Server.CreateObject("ADODB.Recordset")
				strSQL = "SELECT * FROM [tsheets_t] where procmed is null and procpriv is null and procpay is null and date = '" & Request("2day") & "' "
				tblEMP.Open strSQL, g_strCONN, 1, 3
				tblEMP.movefirst
				strTmp3 = "ID=" & tmpctr 
				tblEMP.Find(strTmp3)
				If Not tblEMP.EOF Then
					extfound = false
					If Z_FixNull(tblEMP("procMed")) = "" And Z_FixNull(tblEMP("procPay")) = "" And Z_FixNull(tblEMP("procPriv")) = "" And Z_FixNull(tblEMP("procMile")) = "" Then
						tmpWor2 = tblEMP("emp_id")
						tmpCli2 = tblEMP("client")
						tmpDate2 = tblEMP("date")
						tmpTS2 = tblEMP("timestamp") 
						Set rsEXT = Server.CreateObject("ADODB.RecordSet")
						'sqlEXT = "DELETE * FROM tsheets_t WHERE ID = " & tmpctr + 1
						sqlEXT = "SELECT * FROM tsheets_T WHERE emp_ID = '" & tmpWor2 & "' AND client = '" & tmpCli2 & "' AND date = '" & tmpDate2 & "' " & _
							"AND timestamp = '" & tmpTS2 & "' AND EXT = 1 AND ID = " & tmpctr + 1
						rsEXT.Open sqlEXT, g_strCONN, 1, 3
						If rsEXT.EOF Then
							Set rsEXT2 = Server.CreateObject("ADODB.RecordSet")
							sqlEXT2 = "SELECT * FROM tsheets_t WHERE ID = " & tmpctr + 1
							rsEXT2.Open sqlEXT2, g_strCONN, 1, 3
							If rsEXT2("emp_ID") = tmpWor2 And rsEXT2("client") = tmpCli2 And rsEXT2("date") = tmpDate2 Then
								Set fso = CreateObject("Scripting.FileSystemObject")
								Set ALog = fso.OpenTextFile(AdminLog, 8, True)
								Alog.WriteLine Now & ":: Inital EXT SQL NOT FOUND -- (" & tmpTS2 & " <> " & rsEXT2("timestamp") & ")Timesheet ID: " &  tmpctr + 1 & " was deleted . (week 2) -- UID: " & Session("UserID") & vbCrLf
								Set Alog = Nothing
								Set fso = Nothing 
								rsEXT2.Delete 1
								extfound = true
							Else
								Set fso = CreateObject("Scripting.FileSystemObject")
								Set ALog = fso.OpenTextFile(AdminLog, 8, True)
								Alog.WriteLine Now & ":: Extended timesheet not found for " & tmpctr & " (week 2) -- UID: " & Session("UserID") & vbCrLf
								Set Alog = Nothing
								Set fso = Nothing 
								extfound = false
							End If
							Set rsEXT2 = Nothing
						Else
							rsEXT.Delete 1
							extfound = true
						End If
						Set rsEXT = Nothing
						Set fso = CreateObject("Scripting.FileSystemObject")
						Set ALog = fso.OpenTextFile(AdminLog, 8, True)
						Alog.WriteLine Now & ":: Timesheet ID: " &  tmpctr & " was deleted. (week 2) -- UID: " & Session("UserID") & vbCrLf
						Set Alog = Nothing
						Set fso = Nothing 
						if extfound then 
							tblEmp.DELETE 1
						Else
							Session("MSG") = Session("MSG") & "<br>Unable to delete checked timesheet in week 2. Please make a screen shot of the row in the timesheet and email to patrick@zubuk.com"
						End if
					Else
						Session("MSG") = Session("MSG") & "<br>Unable to delete checked timesheet in week 2. Timesheet may have been billed already."
					End If
				End If
				tblEmp.Close
				Set tblEMP = Nothing
			End If
		Next 
'	End If
'Log
Response.Redirect "view.asp"
%>
