<style type="text/css">
<!--
.style3 {font-family: Verdana; font-size: 11px; }
.style5 {font-family: Verdana; font-size: 11px; color: #FFFFFF; }
.style6 {color: #FFFFFF}
-->
</style>
<script language="JavaScript">
function PopMe(zzz)
	{
		newwindow = window.open(zzz,'name','height=600,width=550,scrollbars=1,directories=0,status=0,toolbar=0,resizable=0');
		if (window.focus) {newwindow.focus()}
	}
</script>
<table width='100%' border='0' cellpadding='0' cellspacing='0' bgcolor='white'>
	<tr>
		<td align='left' style="height: 103px;">
			
				<img src="images/printlogo.jpg" alt="In Home Care" border='0' />
			
		</td>
		<td align='center'>
			<img src="Images/topbannerblankwork.jpg" alt="top" width="520px" height="76px" align="absbottom" >
		</td>
		<td align='center'>
			<a href="default.asp"><img src="Images/SignOutDoor.gif" alt="Log Out" width="75" height="75" border="0" ></a>
		</td>
	</tr>
	<tr height="32" border='0' bgcolor='#040C8B' >		
		<td width="100%" align='right' colspan='3' style="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=0, startColorstr=#c7c9e5, endColorstr=#040C8B);">
			<a href="help.asp" target="_blank" class="style3" style='color: #FFFFFF; text-decoration: none;'><b>Help</b>&nbsp;&nbsp;</a>
		</td>
	</tr>		
</table>
</td></tr>
<tr>
<td>
<table align='left' bgcolor='#C4B464' border='0' width='131' height='100%' leftmargin="0" style="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=1, startColorstr=#C4B464, endColorstr=#f7efc7);">
	<tr>
		<td width="125" height='30px' align='left'><span class="style5">&nbsp;<b>DATABASE</b></span></td>
	</tr>
	<tr>
		<td height='30px' align='left'><span class="style5">&nbsp;&nbsp;<font size='1' >+&nbsp;<a href="A_Choice.asp?choice=Consumer" style='color: #000000; text-decoration: none;'>Consumer</a></span></td>
	</tr>
	<tr>
		<td height='30px' align='left'><span class="style5">&nbsp;&nbsp;<font size='1' >+&nbsp;<a href="A_Choice.asp?choice=Worker" style='color: #000000; text-decoration: none;'>PCSP Worker</a></span></td>
	</tr>
	<tr>
		<td height='30px' align='left'><span class="style5">&nbsp;&nbsp;<font size='1' >+&nbsp;<a href="A_Choice.asp?choice=Case" style='color: #000000; text-decoration: none;'>Case Manager</a></span></td>
	</tr>
	<tr>
		<td height='30px' align='left'><span class="style5">&nbsp;&nbsp;<font size='1' >+&nbsp;<a href="A_Choice.asp?choice=Rep" style='color: #000000; text-decoration: none;'>Representative</a></span></td>
	</tr>
	<tr>
		<td height='30px' align='left'><span class="style5">&nbsp;&nbsp;<font size='1' >+&nbsp;<a href="PM.asp" style='color: #000000; text-decoration: none;'>RIHCC</a></span></td>
	</tr>
	<tr>
		<td height='30px' align='left'><span class="style5">&nbsp;&nbsp;<font size='1' >+&nbsp;<a href="mcc.asp" style='color: #000000; text-decoration: none;'>Managed Care<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Company</a></span></td>
	</tr>
	<tr>
		<td height='30px' align='left'><span class="style5">&nbsp;&nbsp;<font size='1' >+&nbsp;<a href="cmc.asp" style='color: #000000; text-decoration: none;'>Case Management<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Company</a></span></td>
	</tr>
	<tr>
		<td height='30px' ><span class="style6"></span></td>
	</tr>
	<tr>
		<td height='30px' align='left' ><a href='chkdata2.asp?cwork=1' class="style3" style='color: #FFFFFF; text-decoration: none;'>&nbsp;<b>TIMESHEET</b></a></td>
	</tr>
	<tr>
		<td height='30px' align='left' ><a href='SpecRep.asp' class="style3" style='color: #FFFFFF; text-decoration: none;'>&nbsp;<b>REPORTS</b></a></td>
	</tr>
	<% If UCase(Session("lngType")) = "1" or  UCase(Session("lngType")) = "2" Then %>
		<tr>
			<td height='30px' align='left' ><a href='Process.asp' class="style3" style='color: #FFFFFF; text-decoration: none;'>&nbsp;<b>PROCESS ITEMS</b></a></td>
		</tr>
		<tr>
			<td height='30px' align='left' ><a href='http://webapp1.lssnorth.org/feedback/' class="style3" style='color: #FFFFFF; text-decoration: none;' target="_BLANK">&nbsp;<b>FEEDBACK</b></a></td>
		</tr>
	<% Else %>
		<tr>
			<td ><span class="style6"></span></td>
		</tr>
	<% End If%>
	<% If UCase(Session("lngType")) = "2" or UCase(Session("lngType")) = "1" Then %>
		<tr>
			<td height='30px' align='left' ><a href="import.asp" class="style3" style='color: #FFFFFF; text-decoration: none;'>&nbsp;<b>IMPORT</b></a></td>
		</tr>
		<tr>
			<td height='30px' align='left' ><a href="adminchoice.asp" class="style3" style='color: #FFFFFF; text-decoration: none;'>&nbsp;<b>ADMIN TOOLS</b></a></td>
		</tr>
	<% Else %>
		<tr>
			<td ><span class="style6"></span></td>
		</tr>
	<% End If %>
	<tr>
			<td align="left" ><span class="style6"><a href="help.asp" class="style3" style='color: #FFFFFF; text-decoration: none;'></a></span></td>
  </tr>
</table>

