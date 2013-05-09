<%@ Language=VBScript %>
<link href="acss/newstyle.css" rel="stylesheet" type="text/css" />
<div class="Contact_UsPage">
    
<!--#INCLUDE FILE="header.asp"-->

<%
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
set conn = server.CreateObject ("ADODB.Connection")

 if Request.Cookies("isDealer") <> "" then
 xmsg="You are Logged in as " & Request.Cookies("typename") & " : [" & Request.Cookies("dealer_id") & "]"
 end if
%>
<!--#INCLUDE FILE="dbconnect.asp"-->

    <div class="HeroUnitContainer">
        <h3>
            Contact Us</h3>
            <!--#INCLUDE FILE="contactdetail.htm"-->
    </div>
 

  
 <!--#INCLUDE FILE="footer.asp"-->

  </div>
