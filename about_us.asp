<%@ Language=VBScript %>
<link href="acss/newstyle.css" rel="stylesheet" type="text/css" />
<div class="about_us_page" >

<!--#INCLUDE FILE="header.asp"-->
<%
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
set conn = server.CreateObject ("ADODB.Connection")
set rsOffer = server.CreateObject("ADODB.Recordset")
set rscat = server.CreateObject("ADODB.Recordset")
set rssubcat = server.CreateObject("ADODB.Recordset")
set rsdumm = server.CreateObject("ADODB.Recordset")
set rsdumm2 = server.CreateObject("ADODB.Recordset")
 if Request.Cookies("isDealer") <> "" then
 xmsg="You are Logged in as " & Request.Cookies("typename") & " : [" & Request.Cookies("dealer_id") & "]"
 end if
%>
<!--#INCLUDE FILE="dbconnect.asp"-->
<html><head><title>Baylan Technology Inc. Introduction</title> 
 
 <div style="text-align:center" >

<img src="catalog/bl-about-us.jpg">
   
     </div>

 <!--#INCLUDE FILE="footer.asp"-->

  </div>