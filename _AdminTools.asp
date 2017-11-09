<%If UCase(Session("lngType")) <> 0  AND (Paychk = "checked" OR Medchk = "checked") Then%>
	<input type='button' value='Untag >>' onclick='JavaScript: Untag();'>
<%Else%>
	&nbsp;		
<%End If%>