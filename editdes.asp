<%@Language=VBScript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<%

DIM		lngI, tbldes, strSQL, tmpID, ctrI, tmpctr	

Set tbldes = Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT * FROM [tdes_t]"

tbldes.Open strSQL, g_strCONN, 1, 3
ctrI = Request("countb")
	For i = 0 to ctrI 
		tmpctr = Request("dchk" & i)
		If tmpctr <> "" Then
			strTmp = "index='" & tmpctr & "'"
			tblDes.Movefirst
			tbldes.Find(strTmp)
			If Not tbldes.EOF Then
				tbldes("tdes") = Request("des" & i)
			tbldes.Update
			End If
		End If

	Next 
	On Error Resume Next
	x = 0
	Do until tbldes.EOF
		if tbldes("tdes") = Request("des") then
			x = 1
		end if
		tbldes.movenext
	loop
	if x <> 1 then
		tbldes.AddNew
		tbldes("tdes") = Request("des")
		tbldes.Update
	else
		Session("MSG") = Request("des") & " Already Exist!"
	end if
tbldes.Close
Set tbldes = Nothing
Response.Redirect "admin.asp"
%>

