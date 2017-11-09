<%@Language=VBScript%>
<!-- #include file="_Files.asp" -->
<%
	oFile = ouputPath & Request("fname")
	Set objStream = Server.CreateObject("ADODB.Stream")
  objStream.Type = 1 'adTypeBinary
  objStream.Open
  objStream.LoadFromFile(oFile)
  Response.ContentType = "application/x-unknown"
  Response.Addheader "Content-Disposition", "attachment; filename=" & Request("fname")
  Response.BinaryWrite objStream.Read
  objStream.SaveToFile Request("fname"), 2
  objStream.Close
  Set objStream = Nothing
%>