<%@ Language=VBScript %>
<%


Dim ObjSendMail
Set ObjSendMail = CreateObject("CDO.Message") 
     
'This section provides the configuration information for the remote SMTP server.
     
ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 'Send the message using the network (SMTP over the network).
ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") ="smtp.gmail.com"
ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 465 '25 for remote or local server 
ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = 1 'Use SSL for the connection (True or False)
ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
     
' If your server requires outgoing authentication uncomment the lines bleow and use a valid email address and password.

ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1 'basic (clear-text) authentication
ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusername") ="baylanpakistan@gmail.com"
ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendpassword") ="Baylan123!"  '"Smt123!"
     
ObjSendMail.Configuration.Fields.Update
     
'End remote SMTP server configuration section==
ObjSendMail.From = "baylanpakistan@gmail.com" 
ObjSendMail.To = "asgharlatif@hotmail.com"
ObjSendMail.Subject = "Download driver from this URL"

     
' we are sending a text email.. simply switch the comments around to send an html email instead
'ObjSendMail.HTMLBody = "this is the body"
ObjSendMail.TextBody = "Driver URL"
     
ObjSendMail.Send

    
     
Set ObjSendMail = Nothing 

%>