<%@ Language=VBScript %>
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
<html><head><title>Baylan Technology Inc. About us</title> 
 <table width="100%" border="0">
  <tr>
  <td><div align="center"> 
  

</div>
 </td></tr></table>
  <tr></tr>

            <img src="http://www.bbl.com.pk/images/aboutus.jpg">
      <br>
      <table width="100%" border="0" cellspacing="0" cellpadding="0" height="1" background="aimages/dot_grey1.gif">
    <tr> 
    <td height="4"></td>
    </tr>
    </table>
    <table border="0" cellspacing="0" cellpadding="0" width="100%">
    <hr color=#000000 size=1 width="650" align="center" noShade>
    <tr> 
    <td width="100%" height="21" id=msviRegionGradient2 
    style="FILTER: progid:DXImageTransform.Microsoft.Gradient(startColorStr='#FFFFFF', endColorStr='#6487DB', gradientType='1')">	<p align="center" class="font">Baylan Technology Inc. Copyright &copy; 2010</p></td>
	</td>
    </tr>
    <tr> 
     <td width="100%" height="21" id=msviRegionGradient2 
    style="FILTER: progid:DXImageTransform.Microsoft.Gradient(startColorStr='#FFFFFF', endColorStr='#6487DB', gradientType='1')"> 
	    <div align="center"><b> </b><a href="sitemap.asp" class="text0"><u>Site Map 
        </u></a><b> &nbsp;&nbsp;</b>|<b>&nbsp; &nbsp;</b> <a href="our_services.asp" class="text0"><u>Our Services 
        </u></a><b> &nbsp;&nbsp;</b>|<b>&nbsp; &nbsp;</b><a href="private_policy.asp" class="text0"><u>Privacy 
        Policy</u></a><b> &nbsp;&nbsp;</b>|<b>&nbsp;&nbsp; </b><a href="feedback.asp" class="text0"><u>Feed 
        Back</u></a><b> &nbsp;&nbsp;</b>|<b>&nbsp;&nbsp; </b><a href="terms.asp" class="text0"><u>Terms 
        And Conditions</u></a></div></td>
    </tr>
  </table>

