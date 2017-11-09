<%@Language=VBScript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<%
DIM tmpDATE, newDATE
If request("wk") <> 1 Then 
	dayadd = 7
Else
	dayadd = 14
End if
tmpDATE = Session("sundate")
newDATE = DateAdd("d", dayadd, tmpDATE)
Session("dtDATE") = newDATE
Response.Redirect "view.asp"
%>			