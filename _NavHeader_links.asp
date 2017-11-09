<style type="text/css">
<!--
.style3 {font-family: Verdana; font-size: 11px; }
.style5 {font-family: Verdana; font-size: 11px; color: #FFFFFF; }
.style6 {color: #FFFFFF}
-->
</style>
<table width="100%" border='0' cellpadding='0' cellspacing='0' bgcolor='white'>
  <tr valign="center">
    <td width="195" align='left' bgcolor="#FFFFFF" style="width: 100px; height: 100px;;"
				title="Lutheran Social Services of New England"><img src="Images/smartcarelogo.jpg" alt="logo" width="167" height="99" /></td>
    <td width="100" align='center' valign="middle" bgcolor="#FFFFFF" style="width: 50px; height: 100px;;"
				title="Lutheran Social Services of New England">&nbsp;</td>
    <td width="575" align='center' valign="middle" bgcolor="#FFFFFF" style="width: 100px; height: 100px;;"
				title="Lutheran Social Services of New England"><p align="center"><img src="Images/topbannerblankwork.jpg" alt="top" width="475" height="76" align="absbottom" /></p></td>
    <td width="89" align='center' valign="middle" bgcolor="#FFFFFF" style="width: 75px; height: 99px;;"
				title="Lutheran Social Services of New England"><p>&nbsp;</p>
        <p><a href="financeview.asp"></a></p></td>
  </tr>
</table>
<table cellSpacing='0' cellPadding='0' width="100%" bgColor='white' border='0'>		
	<tr bgcolor='#A4CADB'>
	  <td width="83" bgcolor="#040C8B">&nbsp;</td>
	  <td width="154" bgcolor="#040C8B"><div align="center"><a href='default.asp' ></a><img src="Images/Home.gif" alt="home" width="68" height="32" /></div></td>
	  <td width="138" bgcolor="#040C8B"><div align="center"><img src="Images/Reports.gif" alt="reports" width="68" height="32" /></div></td>
	  <td width="151" bgcolor="#040C8B"><div align="center"><img src="Images/timesheet.gif" alt="timesheet" width="68" height="32" /></div></td>
	  <td width="147" bgcolor="#040C8B"><div align="center"><img src="Images/Help.gif" alt="help" width="68" height="32" /></div></td>
	  <td width="134" bgcolor="#040C8B"><div align="center"><a href='default.asp' ><img src='images/SignOut.gif' alt="Sign Out" border='0' /></a></div></td>
	  <td width="114" bgcolor="#040C8B">&nbsp;</td>
		<td bgcolor="#040C8B">&nbsp;</td>
	</tr>		
</table>

<table align='left' bgcolor='#C4B464' border='0' width='131' height='100%' leftmargin="0">
	<tr>
		<td width="125" height='30px' align='left' bordercolor="#C4B464" bgcolor="#C4B464"><span class="style5">&nbsp;<b>DATABASE</b></span></td>
	</tr>
	<tr>
		<td height='30px' align='left' bordercolor="#C4B464" bgcolor="#C4B464"><span class="style5">&nbsp;&nbsp;<font size='1' >+&nbsp;<a href="A_Choice.asp?choice=Consumer" style='color: #000000; text-decoration: none;'>Consumer</a></span></td>
	</tr>
	<tr>
		<td height='30px' align='left' bordercolor="#C4B464" bgcolor="#C4B464"><span class="style5">&nbsp;&nbsp;<font size='1' >+&nbsp;<a href="A_Choice.asp?choice=Worker" style='color: #000000; text-decoration: none;'>PCSP Worker</a></span></td>
	</tr>
	<tr>
		<td height='30px' align='left' bordercolor="#C4B464" bgcolor="#C4B464"><span class="style5">&nbsp;&nbsp;<font size='1' >+&nbsp;<a href="A_Choice.asp?choice=Case" style='color: #000000; text-decoration: none;'>Case Manager</a></span></td>
	</tr>
	<tr>
		<td height='30px' align='left' bordercolor="#C4B464" bgcolor="#C4B464"><span class="style5">&nbsp;&nbsp;<font size='1' >+&nbsp;<a href="A_Choice.asp?choice=Rep" style='color: #000000; text-decoration: none;'>Representative</a></span></td>
	</tr>
	<tr>
		<td height='30px' bordercolor="#C4B464" bgcolor="#C4B464"><span class="style6"></span></td>
	</tr>
	<tr>
		<td height='30px' align='left' bordercolor="#C4B464" bgcolor="#C4B464"><a href='chkdata2.asp?cwork=1' class="style3" style='color: #FFFFFF; text-decoration: none;'>&nbsp;<b>TIMESHEET</b></a></td>
	</tr>
	<tr>
		<td height='30px' align='left' bordercolor="#C4B464" bgcolor="#C4B464"><a href='SpecRep.asp' class="style3" style='color: #FFFFFF; text-decoration: none;'>&nbsp;<b>REPORTS</b></a></td>
	</tr>
	<% If UCase(Session("lngType")) = "1" or  UCase(Session("lngType")) = "2" Then %>
		<tr>
			<td height='30px' align='left' bordercolor="#C4B464" bgcolor="#C4B464"><a href='Process.asp' class="style3" style='color: #FFFFFF; text-decoration: none;'>&nbsp;<b>PROCESS ITEMS</b></a></td>
		</tr>
	<% Else %>
		<tr>
			<td bordercolor="#C4B464" bgcolor="#C4B464"><span class="style6"></span></td>
		</tr>
	<% End If%>
	<% If UCase(Session("lngType")) = "2" Then %>
		<tr>
			<td height='30px' align='left' bordercolor="#C4B464" bgcolor="#C4B464"><a href="adminchoice.asp" class="style3" style='color: #FFFFFF; text-decoration: none;'>&nbsp;<b>ADMIN TOOLS</b></a></td>
		</tr>
	<% Else %>
		<tr>
			<td bordercolor="#C4B464" bgcolor="#C4B464"><span class="style6"></span></td>
		</tr>
	<% End If %>
	<tr>
			<td bordercolor="#C4B464" bgcolor="#C4B464"><span class="style6"></span></td>
  </tr>
</table>

