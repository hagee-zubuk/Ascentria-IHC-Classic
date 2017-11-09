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
If	Session("PrintPrevPRoc") <> "" Then
	 tmpPrint = Session("PrintPrevPRoc")
	 PrintPrev = Split(tmpPrint, "|")
	 If PrintPrev(0) <> "" Then strHeader = PrintPrev(0)
	 If PrintPrev(1) <> "" Then strTBLer = PrintPrev(1)
   If PrintPrev(2) <> "" Then sMSG = PrintPrev(2)
   If PrintPrev(3) <> "" Then HX = PrintPrev(3)
   If PrintPrev(4) <> "" Then BX = PrintPrev(4)
	 If PrintPrev(5) <> "" Then MSG = PrintPrev(5)	
	If PrintPrev(6) <> "" Then MSG2 = PrintPrev(6)	
	If PrintPrev(7) <> "" Then HX2 = PrintPrev(7)	
	If PrintPrev(8) <> "" Then BX2 = PrintPrev(8)	
	If PrintPrev(9) <> "" Then MSG3 = PrintPrev(9)	
	If PrintPrev(10) <> "" Then HX3 = PrintPrev(10)	
	If PrintPrev(11) <> "" Then BX3 = PrintPrev(11)	
End If
If	Session("PrintPrevRep") <> "" Then
	 tmpPrint = Session("PrintPrevRep")
	 PrintPrev = Split(tmpPrint, "|")
	 If PrintPrev(0) <> "" Then strTBL = PrintPrev(0)
	 If PrintPrev(1) <> "" Then MSG = PrintPrev(1)
End If	

%>
<html>
	<head>
		<title>LSS - Print</title>
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
				<img src='images/heartmaster1in.gif' border='0'?
			</center>
		<br>
		<table align='center' cellspacing='2' cellpadding='0' border='1' bgcolor='#FFFFFF' >
			<tr><td valign='top'>
		<% If Session("PrintPrev") <> "" THen %>
			<table align='left' cellspacing='2' cellpadding='0' border='1' bgcolor='#FFFFFF' >
				<tr><td align='center' colspan='10'><font size='2'><%=MSG%></font></td></tr>
				<%=strHead%>
				<%=strBody%>
				<% If strHEad2 <> "" Then %>
					<tr><td colspan='11'>&nbsp;</td></tr>
					<%=strHead2%>
					<%=strBody2%>
				<% End If %>
			</table>
		<% End If %>
		<% If Session("PrintPrevPRoc") <> "" THen %>	
			<table align='left' cellspacing='2' cellpadding='0' border='1' bgcolor='#FFFFFF' >
				<tr><td align='center' colspan='10' width='100%'><font size='2'><%=MSG%></font></td></tr>
					<%=strHeader%>	
					<%=strTBLer%>
			</table>
			</td></tr>
			<tr><td colspan='4'>&nbsp;</td></tr>
			<tr><td>
			<table border='1' width='100%'>
					<%=sMSG%>
					<%=HX%>
					<%=BX%>	
			</table>
			<br>
				<table border='1' width='100%'>
					<%=MSG2%>
					<%=HX2%>
					<%=BX2%>	
			</table>
			<br>
				<table border='1' width='100%'>
					<%=MSG3%>
					<%=HX3%>
					<%=BX3%>	
			</table>
		<%End If %>
		<% If Session("PrintPrevRep") <> "" THen %>	
			<table align='left' cellspacing='2' cellpadding='0' border='1' bgcolor='#FFFFFF' >
				<tr><td align='center' colspan='10'><font size='2'><%=MSG%></font></td></tr>
				<tr><td>
					<%=strTBL%>
				</td></tr>
			</table>
		<% End If %>
		</td></tr>
	</table>
	</body>
</html>
<%
Session("PrintPrev") = ""
Session("PrintPrevRep") = ""
Session("PrintPrevPRoc") = ""
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
