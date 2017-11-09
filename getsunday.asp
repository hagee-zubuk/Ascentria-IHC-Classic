<%@Language=VBScript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<%
GoToSunday = DateSerial(Year(keyDate) _
        , Month(keyDate) _
        , Day(keyDate) - DatePart("w", keyDate) + 1)
%>