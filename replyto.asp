<% 
Response.Buffer = False 
Dim strBody 
Dim objCDOMail 
Dim strMyEmailAddress 
Dim strCCEmailAddress 
Dim strBCCEmailAddress 
Dim strReturnEmailAddress 


name=Request.Form("name")
email=Request.Form("email")
msg=Request.Form("msg")
subject=Request.Form("subject")
emailfrom=Request.Form("emailfrom")

	strMyEmailAddress = Request.Form("email") 
	strReturnEmailAddress = Request.Form("emailfrom") 


strBody = strBody & " " &  msg  & " "

Set objCDOMail = Server.CreateObject("CDONTS.NewMail") 
objCDOMail.From = emailfrom 
objCDOMail.To = strMyEmailAddress
objCDOMail.Cc = strCCEmailAddress 
objCDOMail.Bcc = strBCCEmailAddress 
objCDOMail.Subject = subject 
objCDOMail.BodyFormat = 0 
objCDOMail.MailFormat = 0 
objCDOMail.Body = strBody 
objCDOMail.Importance = 2 
objCDOMail.Send 
Set objCDOMail = Nothing 
%>