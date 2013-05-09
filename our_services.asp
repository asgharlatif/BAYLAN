<%@ Language=VBScript %>
<link href="acss/newstyle.css" rel="stylesheet" type="text/css" />
<div class="OurServicesPage">

<!--#INCLUDE FILE="header.asp"-->
<br>
<%
 if Request.Cookies("isDealer") <> "" then
 xmsg="You are Logged in as " & Request.Cookies("typename") & " : [" & Request.Cookies("dealer_id") & "]"
 end if
%>

<html><head><title>Baylan Technology Inc. OurServices Page</title> 
 <table width="100%" border="0">
  <tr>
  <td><div align="center"> 
  <%  
   Dim objAd
   'Set objAd = Server.CreateObject("MSWC.AdRotator") 'Create the component
   'Response.Write  objAd.GetAdvertisement ("/banners/adrot.txt") 
   'Set objAd = Nothing 'Destroy the object
   %>
</div>
 </td></tr></table>
  <tr></tr>
  <table width="100%"  border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
  <div align="center"> 
    <table width="100%" border="0" cellspacing="0" cellpadding="0" height="1" background="aimages/dot_grey1.gif">
    <tr> 
   
    </tr>
    </table>
    

     <!--#INCLUDE FILE="footer.asp"-->

  </div>