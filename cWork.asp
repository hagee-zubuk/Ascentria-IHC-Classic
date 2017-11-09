<%@Language=VBScript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<%
	
		tmpID = split(Request("worker"), " - ") 
		Response.Redirect "chkdata2.asp?indate=" & Session("d8") & " &tmpID='" & tmpID(0) & "' &cwork=1" 
		
	
%>