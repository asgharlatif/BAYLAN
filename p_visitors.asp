<!--#INCLUDE FILE="dbconnect.asp"-->
<html>
<head>
<title>Vistors Mail Send</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body bgcolor="#FFFFFF" text="#000000">
<%
email=Request.Cookies("EmailT")
Response.Cookies("EmailT") = ""
NameT=Request.Cookies("NameT")
Response.Cookies("NameT") = ""
subt=Request.Cookies("subt")
Response.Cookies("subt")=""
 Const adOpenForwardOnly = 0
 Const adLockReadOnly = 1
 Const adCmdTableDirect = &H0200
 Const adUseClient = 3
 Set Conn = Server.CreateObject("ADODB.Connection")
 conn.Open con
'on error resume next
if trim(email) <> "" then
set RsF=server.createobject("adodb.recordset")
RsF.CursorLocation = adUseClient
RsF.Open  "select * from visitors where email='" & trim(email) & "'",conn,adOpenForwardOnly,adLockReadOnly,adcmdtext
set RsF.ActiveConnection=nothing
	if rsf.EOF <> true  then
 if  (subt)="sc" then

if (rsf("myaction"))="nc" then
Conn.Execute("update visitors set myaction='sc' where email='" & trim(email) & "'")
Response.Write "<p align='center'> <font face='verdana,arial' color='blue'>Subscribe Complete</font></p>"
	else
Response.Write "<p align='center'> <font face='verdana,arial' color='red'>Already Subscribe Found.</font></p>"
	end if
	end if
  if  (subt)="nc" then
	 if rsf("myaction")="sc" then
Conn.Execute("update visitors set myaction='nc' where email='" & trim(email) & "'")
Response.Write "<p align='center'> <font face='verdana,arial' color='blue'>Unsubscribe Complete</font></p>"
else
Response.Write "<p align='center'> <font face='verdana,arial' color='red'>Already Unsubscribe Found.</font></p>"
end if
	end if
	end if	
	if rsf.EOF <> false  then	
	   		if subt="sc" then		
		SQLS="insert into visitors "
		SQLS=SQLS + "(name,visitdate,myaction,email) "
		SQLS=SQLS + "values ('" & NameT & "', "
		SQLS=SQLS + "'"  & date &  " ', "
		SQLS=SQLS + "'sc', "
		SQLS=SQLS + " '"  & (email) &  "') "
		Conn.Execute(sqls)

' mb = "<IMG SRC=""http://www.bbl.com.pk/images/logobbl3.gif""<BR><br><BR><BR>" & chr(13) &chr(10)
mb=mb & "<font face=Verdana size=2> "
mb=mb & "Dear " & Namet 
mb=mb & ",<br><br>We offer our sincere thanks to you for visiting our website <strong><font face=Verdana size=2 color=#FF0033><A HREF=""http://www.bbl.com.pk"">bbl.com.pk</A></font></strong> and enrolling / subsribing as a visitor."
mb=mb & "<BR><br>BBL is a computer-based enterprise with resolute commitment to provide best solution to Computer and Network Users. Our dedicated efforts have enabled us to achieve sole distributor status in Pakistan for BAYLAN and CELL POWER products." 
mb=mb & "<br><br>We also welcome you most cordially if you desire to subscribe as a Dealer / Reseller / Member for our products sale / use , which all carry BAYLAN (Networking, Communication & Structured Cabling Products) and CELL POWER (Intelligent UPS+AVR+LSP+CIP). We also keep stock of various computer related products."
mb=mb & "<br><br>We hope that you will be able to obtain all necessary information and technical details from our website for your personal and business purposes. We will keep you updated on all products information and specifications."
mb=mb & "<br><br>We will welcome all useful suggestions from you if you require any additional information and technical assistance."
mb=mb & "<br><br>We hope you will keep on visiting our website over the coming period for which we offer our sincere thanks."
mb=mb & "<br><br>Best wishes."
mb=mb & "</font>"


'Dim Mail, Body
'Set Mail = Server.CreateObject("Persits.MailSender")
Dim objMessage, Body
Set objMessage  = Server.CreateObject("CDO.Message")
set rsEmail= server.CreateObject("ADODB.Recordset")
Setxt="select * from extras" 
rsemail.CursorLocation = adUseClient
rsemail.Open Setxt,conn,adOpenForwardOnly,adLockReadOnly,adcmdtext
set rsemail.ActiveConnection = nothing
'set rsEmail= server.CreateObject("ADODB.Recordset")
'Setxt="select * from extras"
'rsemail.CursorLocation = adUseClient
'rsemail.Open Setxt,conn,adOpenForwardOnly,adLockReadOnly,adcmdtext
'set rsemail.ActiveConnection = nothing 
objMessage.From = "BBL.COM.PK (Subscribe)" & rsemail("subscribe")  'From 
objMessage.To = email 'email
objMessage.Subject = "Congratulations " & Namet
objMessage.HTMLBody = "<HTML><BODY>" & mb & "</BODY></HTML>"
objMessage.Send
'mail.Host =rsemail("ip")
'mail.From =rsemail("subscribe")  'From 

'Mail.FromName = "BBL.COM.PK (Subscribe)"
'Mail.AddAddress email 'email
'Mail.Subject = "Congratulations " & Namet
'Mail.IsHTML = True
'Mail.Body = "<HTML><BODY>" & mb & "</BODY></HTML>"
'Mail.Send
If Err <> 0 Then 
    Response.Write "An error occurred: " & Err.Description 
End If

'Set Mail = nothing
Set objMessage = nothing
Response.Write "<p align='center'> <font face='verdana,arial' color='blue'>Congratulations ! You have sign up successfully.</font></p>"
Response.Write "<p align='center'> <font face='verdana,arial' color='blue'><b> A mail has been sent to you as well.</b></font></p>"
			end if
	        if subt="nc" then
Response.Write "<p align='center'> <font face='verdana,arial' color='red'>Error! No subscription found</font></p>"
	        end if
	end if
end if
%>

</body>
</html>
