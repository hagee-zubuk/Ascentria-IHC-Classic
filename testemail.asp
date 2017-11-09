<%Language=VBScript%>
<%
Set mlMail = CreateObject("CDO.Message")
			mlMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
			mlMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "localhost"
			mlMail.Configuration.Fields.Update
			mlMail.To = "patrick@zubuk.com"
			mlMail.From = "info@smart-care.org"
			mlMail.Subject = "test email"
			strBody = "<font size='2' face='trebuchet MS'>test email</font>"
			mlMail.HTMLBody = "<html><body>" & vbCrLf & strBody & vbCrLf & "</body></html>"
			mlMail.Send
			'response.write strBody & "<br><br><Br>"
			Set mlMail = Nothing
%>