<%@ Language=VBScript %>
<link href="acss/newstyle.css" rel="stylesheet" type="text/css" />
<div class="SiteMapPage" >

<!--#INCLUDE FILE="header.asp"-->
<br>
<%
 if Request.Cookies("isDealer") <> "" then
 xmsg="You are Logged in as " & Request.Cookies("typename") & " : [" & Request.Cookies("dealer_id") & "]"
 end if
 %>
<html><head><title>Baylan Technology Inc. SiteMap</title> 

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

<br>
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="38">
  <tr>
    <td height="18"> 
      <div align="left"><font size="6" face="Verdana, Arial, Helvetica, sans-serif" color="#666666"><b>SITE 
        MAP</b></font></div>
    </td>
  </tr>
  <tr>
    <td height="9">
      <div align="center"><img src="aimages/NOIMAGE1.gif" width="600" height="350" border="0"></div>
  
    </td>
  </tr>
</table>
<br>
<br>
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="1" background="aimages/dot_grey1.gif">
    <tr> 
   
    </tr>
</table>

 <!--#INCLUDE FILE="footer.asp"-->

</div>


</body>
</html>