<% 
Response.Buffer = False 
Dim strBody 
Dim objCDOMail 
Dim strMyEmailAddress 
Dim strCCEmailAddress 
Dim strBCCEmailAddress 
Dim strReturnEmailAddress 


subject=Request.Form("subject")
name=Request.Form("name")
email=Request.Form("email")
company=Request.Form("company")
address=Request.Form("address")
city=Request.Form("city")
state=Request.Form("state")
zip=Request.Form("zip")
country=Request.Form("country")
phone=Request.Form("phone")
interested=Request.Form("interested")
aboutus=Request.Form("aboutus")
msg=Request.Form("msg")

	strMyEmailAddress = "bharat.kr@gmail.com" 
	strReturnEmailAddress = Request.Form("email") 


strBody = "<table border=1 cellpadding=1 cellspacing=1 bordercolor=#000000 width=700><tr><td colspan=4><h4>E-Mail from " & name  & "</h4></td></tr>" 
strBody = strBody & "<tr><td colspan=3><table width=250><tr><td><br><b><h5>NAME:</h5></b></td></tr></table></td><td> " & name  & "</td></tr>"
strBody = strBody & "<tr><td colspan=3><table width=250><tr><td><br><b><h5>E-MAIL:</h5></b></td></tr></table></td><td> " &  email  & "</td></tr>"
strBody = strBody & "<tr><td colspan=3><table width=250><tr><td><br><b><h5>COMPANY:</h5></b></td></tr></table></td><td> " &  company  & "</td></tr>"
strBody = strBody & "<tr><td colspan=3><table width=250><tr><td><br><b><h5>ADDRESS:</h5></b></td></tr></table></td><td> " &  address  & "</td></tr>"
strBody = strBody & "<tr><td colspan=3><table width=250><tr><td><br><b><h5>CITY:</h5></b></td></tr></table></td><td> " &  city  & "</td></tr>"
strBody = strBody & "<tr><td colspan=3><table width=250><tr><td><br><b><h5>STATE:</h5></b></td></tr></table></td><td> " &  state  & "</td></tr>"
strBody = strBody & "<tr><td colspan=3><table width=250><tr><td><br><b><h5>ZIP:</h5></b></td></tr></table></td><td> " &  zip  & "</td></tr>"
strBody = strBody & "<tr><td colspan=3><table width=250><tr><td><br><b><h5>COUNTRY:</h5></b></td></tr></table></td><td> " &  country  & "</td></tr>"
strBody = strBody & "<tr><td colspan=3><table width=250><tr><td><br><b><h5>PHONE NO.:</h5></b></td></tr></table></td><td> " &  phone  & "</td></tr>"
strBody = strBody & "<tr><td colspan=3><table width=250><tr><td><br><b><h5>PERSON IS INTERESTED IN:</h5></b></td></tr></table></td><td> " &  interested  & "</td></tr>"
strBody = strBody & "<tr><td colspan=3><table width=250><tr><td><br><b><h5>PERSON HEARD ABOUT US FROM:</h5></b></td></tr></table></td><td> " &  aboutus  & "</td></tr>"
strBody = strBody & "<tr><td colspan=3><table width=250><tr><td><br><b><h5>MESSAGE:</h5></b></td></tr></table></td><td> " &  msg  & "</td></tr></table>"

Set objCDOMail = Server.CreateObject("CDONTS.NewMail") 
objCDOMail.From = email 
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