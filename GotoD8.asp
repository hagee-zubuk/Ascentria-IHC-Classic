<%@Language=VBScript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<%
If Request("specd8") = "" OR Not IsDate(Request("specd8")) Then
	Session("dtDate") = date
Else
	Session("dtDate") = Request("specd8") 
End If
Response.Redirect "view.asp"
%>