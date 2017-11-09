<!-- #include file="_Utils.asp" -->
<%
 
DIM	m_User, m_Pass

m_User = Request.Cookies("TSHEET")("User")

If m_User = "" Then 
	Session("MSG") = "Please Log-in first/again."
	'Response.Write "default.asp --> Please Log-in first." & 
	Response.Redirect "Default.asp"
	
End if

If Session("lngType") = "" Then
	Session("MSG") = "Please Log-in first/again."
	'Response.Write "default.asp --> Please Log-in first." & 
	Response.Redirect "Default.asp"
End If

If session("UserID") = "" Then
	Session("MSG") = "Please Log-in first/again."
	'Response.Write "default.asp --> Please Log-in first." & 
	Response.Redirect "Default.asp"
End If

%>

