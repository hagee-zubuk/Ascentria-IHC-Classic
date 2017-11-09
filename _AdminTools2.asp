<%If UCase(Session("lngType")) <> 0 AND (Paychk2 = "checked" OR Medchk2 = "checked") Then%>
	<input type='button' value='Untag >>' onclick='JavaScript: Untag2();'>
<%Else%>
	&nbsp;		
<%End If%>