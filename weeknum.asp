<%@Language=VBScript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<%
		seluri = Request("uri")
		selopt = Request("opt")
		seltype = Request("seltype")
		PDate = Request("tmpdate")
		difwk = DateDiff("ww", wk1, PDate)
		myDATE = PDate
    If difwk >= 0 Then
        wknum = difwk + 1
        If Z_IsOdd2(wknum) Then
            If WeekdayName(Weekday(myDATE), True) = "Sun" Then
                sunDATE = myDATE
                If selopt = 0 Then
                	satDATE = DateAdd("d", 13, sunDATE)
            		ElseIf selopt = 1 Then
            			satDATE = DateAdd("d", 27, sunDATE)
            		ElseIf selopt = 2 Then 	
            			satDATE = DateAdd("d", 41, sunDATE)
            		End If
            Else
                difDATE = DatePart("w", myDATE)
                sunDATE = DateAdd("d", -CInt(difDATE - 1), myDATE)
                If selopt = 0 Then
                	satDATE = DateAdd("d", 13, sunDATE)
            		ElseIf selopt = 1 Then
            			satDATE = DateAdd("d", 27, sunDATE)
            		ElseIf selopt = 2 Then 	
            			satDATE = DateAdd("d", 41, sunDATE)
            		End If
            End If
        Else
            If WeekdayName(Weekday(myDATE), True) = "Sun" Then
                sunDATE = DateAdd("d", -7, myDATE)
               If selopt = 0 Then
                	satDATE = DateAdd("d", 13, sunDATE)
            		ElseIf selopt = 1 Then
            			satDATE = DateAdd("d", 27, sunDATE)
            		ElseIf selopt = 2 Then 	
            			satDATE = DateAdd("d", 41, sunDATE)
            		End If
            Else
                difDATE = DatePart("w", myDATE)
                sunDATE = DateAdd("d", -CInt(difDATE + 6), myDATE)
                If selopt = 0 Then
                	satDATE = DateAdd("d", 13, sunDATE)
            		ElseIf selopt = 1 Then
            			satDATE = DateAdd("d", 27, sunDATE)
            		ElseIf selopt = 2 Then 	
            			satDATE = DateAdd("d", 41, sunDATE)
            		End If
            End If
        End If
    Else
        wknum = difwk
        If Not Z_IsOdd2(wknum) Then
            If WeekdayName(Weekday(myDATE), True) = "Sun" Then
                sunDATE = myDATE
                If selopt = 0 Then
                	satDATE = DateAdd("d", 13, sunDATE)
            		ElseIf selopt = 1 Then
            			satDATE = DateAdd("d", 27, sunDATE)
            		ElseIf selopt = 2 Then 	
            			satDATE = DateAdd("d", 41, sunDATE)
            		End If
            Else
                difDATE = DatePart("w", myDATE)
                sunDATE = DateAdd("d", -CInt(difDATE - 1), myDATE)
                If selopt = 0 Then
                	satDATE = DateAdd("d", 13, sunDATE)
            		ElseIf selopt = 1 Then
            			satDATE = DateAdd("d", 27, sunDATE)
            		ElseIf selopt = 2 Then 	
            			satDATE = DateAdd("d", 41, sunDATE)
            		End If
            End If
        Else
            If WeekdayName(Weekday(myDATE), True) = "Sun" Then
                sunDATE = DateAdd("d", -7, myDATE)
               If selopt = 0 Then
                	satDATE = DateAdd("d", 13, sunDATE)
            		ElseIf selopt = 1 Then
            			satDATE = DateAdd("d", 27, sunDATE)
            		ElseIf selopt = 2 Then 	
            			satDATE = DateAdd("d", 41, sunDATE)
            		End If
            Else
                difDATE = DatePart("w", myDATE)
                sunDATE = DateAdd("d", -CInt(difDATE + 6), myDATE)
               If selopt = 0 Then
                	satDATE = DateAdd("d", 13, sunDATE)
            		ElseIf selopt = 1 Then
            			satDATE = DateAdd("d", 27, sunDATE)
            		ElseIf selopt = 2 Then 	
            			satDATE = DateAdd("d", 41, sunDATE)
            		End If
            End If
        End If
    End If
		'response.write "SUN: " & sunDATE & "<br>SAT: " & satDATE & "<br>week#: " & tmpDate 
		if Request("selRep") = 69 Then satDate = DateAdd("d", CDate(sunDate), 6)
		Response.Redirect "specrep.asp?chkr=1&seluri=" & seluri & "&selopt=" & selopt & "&seltype=" & seltype & "&sunDATE=" & sunDATE & " &satDATE=" & satDATE & " &selRep=" & Request("selRep") & "&selcon=" & Request("selcon")
		
	
%>