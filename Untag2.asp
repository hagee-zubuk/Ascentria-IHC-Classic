<%@Language=VBScript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<%
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set ALog = fso.OpenTextFile(AdminLog, 8, True)
	Alog.WriteLine Now & ":: Timesheet (Date: " & Request("2day") & ")" & " of PCS Worker: (" & Right(Request("eid"), 4) & _
		 ") was untagged by " & Session("UserID") & " - " & Request("User") & ". "  & vbCrLf
	Set Alog = Nothing
	Set fso = Nothing 
	'NORMAL HOURS
	Set rsAdmin = Server.CreateObject("ADODB.RecordSet")
	sqlAdmin = "SELECT * FROM Tsheets_t WHERE id = " & Request("id") 
	rsAdmin.Open sqlAdmin, g_strCONN, 1, 3
	If Not rsAdmin.EOF Then
		Do Until rsAdmin.EOF
			rsAdmin("ProcPay") = Empty
			rsAdmin("ProcMed") = Empty
			rsAdmin("ProcPriv") = Empty	
			rsAdmin("ProcVA") = Empty	
			rsAdmin.Update 
			rsAdmin.MoveNext
		Loop
	End If
	rsAdmin.Close
	Set rsAdmin = Nothing
	'EXTENDED HOURS
	Set rsAdmin = Server.CreateObject("ADODB.RecordSet")
	sqlAdmin = "SELECT * FROM Tsheets_t WHERE id = " & Request("id") + 1
	rsAdmin.Open sqlAdmin, g_strCONN, 1, 3
	If Not rsAdmin.EOF Then
		Do Until rsAdmin.EOF
			rsAdmin("ProcPay") = Empty
			rsAdmin("ProcMed") = Empty
			rsAdmin("ProcPriv") = Empty	
			rsAdmin("ProcVA") = Empty		
			rsAdmin.Update 
			rsAdmin.MoveNext
		Loop
	End If
	rsAdmin.Close
	Set rsAdmin = Nothing
	Response.Redirect "View.asp"
%>