<style type="text/css">
<!--
.style2 {color: #FFFFFF}
.style4 {
	font-family: Verdana;
	font-weight: bold;
}
.style5 {font-family: Verdana}
.style9 {font-size: 11px; }
.style10 {font-family: Verdana; font-weight: bold; font-size: 11px; }
-->
</style>
<table width="100%" border='0' cellpadding='0' cellspacing='0' bgcolor='white'>
  <tr valign="center">
    <td width="195" align='left' bgcolor="#FFFFFF" style="width: 100px; height: 100px;;"
				title="Lutheran Social Services of New England"><img src="Images/In-Home Carelogo.jpg" alt="logo" width="167" height="99" /></td>
    <td width="100" align='center' valign="middle" bgcolor="#FFFFFF" style="width: 50px; height: 100px;;"
				title="Lutheran Social Services of New England">&nbsp;</td>
    <td width="575" align='center' valign="middle" bgcolor="#FFFFFF" style="width: 100px; height: 100px;;"
				title="Lutheran Social Services of New England"><p align="center"><img src="Images/topbannerblankwork.jpg" alt="banner" width="475" height="76" align="absbottom" /></p></td>
    <td width="89" align='center' valign="middle" bgcolor="#FFFFFF" style="width: 75px; height: 99px;;"
				title="Lutheran Social Services of New England"><p>&nbsp;</p>
        <p><a href="timesheet.asp">Timesheet</a></p></td>
  </tr>
</table>
<table cellSpacing='0' cellPadding='0' width="100%" bgColor='white' border='0'>		
	<tr bgcolor='#A4CADB'>
	  <td width="951" bgcolor="#000C8C"><a href='default.asp' ></a>
          <table width="100%" border="0">
            <tr>
              <td width="83">&nbsp;</td>
              <td width="154"><div align="center"><a href='default.asp' ></a><img src="Images/Home.gif" alt="home" width="68" height="32" /></div></td>
              <td width="138"><div align="center"><img src="Images/Reports.gif" alt="reports" width="68" height="32" /></div></td>
              <td width="151"><div align="center"><img src="Images/timesheet.gif" alt="timesheet" width="68" height="32" /></div></td>
              <td width="147"><div align="center"><img src="Images/Help.gif" alt="help" width="68" height="32" /></div></td>
              <td width="134"><div align="center"><a href='default.asp' ><img src='images/SignOut.gif' alt="Sign Out" border='0' /></a></div></td>
              <td width="114">&nbsp;</td>
            </tr>
        </table></td>
	</tr>		
</table>

<table align='left' bgcolor='#D6D3CE' border='0' width='124' height='100%' leftmargin="0">
	<tr>
		<td width="118" height='30px' align='left' bgcolor="#C4B464">&nbsp;</td>
	</tr>
	<tr>
		<td height='30px' align='left' bgcolor="#C4B464"><span class="style2 style5 style9">&nbsp;&nbsp;<font size='1' ></span></td>
	</tr>
	<tr>
		<td height='30px' align='left' bgcolor="#C4B464"><span class="style2 style5 style9">&nbsp;&nbsp;<font size='1' ></span></td>
	</tr>
	<tr>
		<td height='30px' align='left' bgcolor="#C4B464"><span class="style2 style5 style9">&nbsp;&nbsp;<font size='1' ></span></td>
	</tr>
	<tr>
		<td height='30px' align='left' bgcolor="#C4B464"><span class="style2 style5 style9">&nbsp;&nbsp;</span></td>
	</tr>
	<tr>
		<td height='30px' bgcolor="#C4B464">&nbsp;</td>
	</tr>
	<tr>
		<td height='30px' align='left' bgcolor="#C4B464" class="style4"><a href='chkdata2.asp?cwork=1' class="style9" style='color: #FFFFFF; text-decoration: none;'>&nbsp;</a></td>
	</tr>
	<tr>
		<td height='30px' align='left' bgcolor="#C4B464"><a href='SpecRep.asp' style='color: #FFFFFF; text-decoration: none;'><span class="style10">&nbsp;</span></a></td>
	</tr>
	<% If UCase(Session("lngType")) = "1" or  UCase(Session("lngType")) = "2" Then %>
		<tr>
			<td height='30px' align='left' bgcolor="#C4B464"><a href='Process.asp' class="style9" style='color: #FFFFFF; text-decoration: none;'>&nbsp;</a></td>
		</tr>
	<% Else %>
		<tr>
			<td height="30" bgcolor="#C4B464">&nbsp;</td>
		</tr>
	<% End If%>
	<% If UCase(Session("lngType")) = "2" Then %>
		<tr>
			<td height='30px' align='left' bgcolor="#C4B464"><a href="adminchoice.asp" class="style9" style='color: #FFFFFF; text-decoration: none;'>&nbsp;</a></td>
		</tr>
	<% Else %>
		<tr>
			<td height="30" bgcolor="#C4B464">&nbsp;</td>
		</tr>
	<% End If %>
	<tr>
			<td height="30" bgcolor="#C4B464">&nbsp;</td>
  </tr>
</table>

<div align="center">
  <p>
    <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,19,0" width="309" height="400" title="help">
      <param name="movie" value="Microsoft Word - LSS Database Manual latest.swf" />
      <param name="quality" value="high" />
      <embed src="Microsoft Word - LSS Database Manual latest.swf" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" width="309" height="400"></embed>
    </object>
  </p>
  <p>&nbsp;</p>
</div>
