<%Language=VBScript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<%
	Set tblOpen = Server.CreateObject("ADODB.RecordSet")
		sqlOpen = "SELECT * FROM Process_t"
		tblOpen.Open sqlOpen, g_strCONN, 1, 3
		If Not tblOpen.EOF Then
			ctr = Request("ctr")
			For i = 0 to ctr 
				tmpctr = Request("chk" & i)
				If tmpctr <> "" Then
					strTmp = "index= " & tmpctr 
					tblOpen.Movefirst
					tblOpen.Find(strTmp)
					If Not tblOpen.EOF Then
						tblOpen("proc") = true
						tblOpen.Update
					End If
				End If
			Next 
		End If
		tblOpen.Close
	Set tblOpen = Nothing
	Response.Redirect "Process.asp"
%>