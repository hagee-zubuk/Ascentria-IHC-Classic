<%Language=VBScript%>
<!-- #include file="_Utils.asp" -->
<%
On Error Resume Next
If Session("PrintPrev") <> "" Then
	 tmpPrint = Session("PrintPrev")
	 PrintPrev = Split(tmpPrint, "|")	
	 If PrintPrev(0) <> "" Then strHead = PrintPrev(0)
	 If PrintPrev(1) <> "" Then strBody = PrintPrev(1)
	 If PrintPrev(2) <> "" Then MSG = PrintPrev(2)
	 If Ubound(PrintPrev) > 2 Then
	 	strHead2 = PrintPrev(4)
	  strBody2 = PrintPrev(3)	
	 End If
End If

%>
<html>
	<head>
		<title>LSS - Print2</title>
		<link href='style.css' type='text/css' rel='stylesheet'>
		<style>
			Input.btn{
			font-size: 7.5pt;
			font-family: arial;
			color:#000000;
			font-weight:bolder;
			background-color:#d4d0c8;
			border:2px solid;
			text-align: center;
			border-top-color:#d4d0c8;
			border-left-color:#d4d0c8;
			border-right-color:#b6b3ae;
			border-bottom-color:#b6b3ae;
			filter:progid:DXImageTransform.Microsoft.Gradient(GradientType=0,StartColorStr='#ffffffff',EndColorStr='#d4d0c8');
		}
		INPUT.hovbtn{
			font-size: 7.5pt;
			font-family: arial;
			color:#000000;
			font-weight:bolder;
			background-color:#b6b3ae;
			border:2px solid;
			text-align: center;
			border-top-color:#b6b3ae;
			border-left-color:#b6b3ae;
			border-right-color:#d4d0c8;
			border-bottom-color:#d4d0c8;
			filter:progid:DXImageTransform.Microsoft.Gradient(GradientType=0,StartColorStr='#ffffffff',EndColorStr='#b6b3ae');
		}  
		@media print
			{
			  .page-break  { display:block; page-break-before:always; }
			}
		</style>
	</head>
	<body>
		<% If Session("PrintPrev") <> "" Then %>
			&nbsp;<input type='button' value='back' class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick="document.location='specrep.asp';">
		<% ElseIf Session("PrintPrevPRoc") <> "" Then %>
			&nbsp;<input type='button' value='back' class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick="document.location='process.asp';">
		<% ElseIf Session("PrintPrevRep") <> "" Then %>
			&nbsp;<input type='button' value='back' class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick="document.location='report.asp';">
		<%End If%>
		&nbsp;<input type='button' value='print' class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='print();'>
		<br>
			<center>
		<%= Session("MSG")%>
		<br>
		
		<% If Session("PrintPrev") <> "" THen %>
			
				<%=strHead%>
				<%=strBody%>
				
		<% End If %>
		

	</body>
</html>
<%
Session("PrintPrev") = ""
tmpPrint = ""
strHEAD = ""
strBODY = ""
strTBL = ""
strHEADer = ""
strBODYer = ""
strTBLer = ""
MSG = ""
sMSG = ""
HX = ""
BX = ""
%>
