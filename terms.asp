<%@ Language=VBScript %>



<link href="acss/newstyle.css" rel="stylesheet" type="text/css" />
<div class="TermPage">

<!--#INCLUDE FILE="header.asp"-->
<br>
 <% if Request.Cookies("isDealer") <> "" then
 xmsg="You are Logged in as " & Request.Cookies("typename") & " : [" & Request.Cookies("dealer_id") & "]"
 end if %>
 <html><head><title>Baylan Technology Inc. Term &amp; Conditions Page</title> 

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
      <table cellspacing=0 cellpadding=10 width=687 border=0 align="center">
        <tbody> 
        <tr> 
          <td valign="top"> 
          <p align="justify"> <br>
            </p></td>
        </tr>
        <tr> 
          <td valign="top">&nbsp; </td>
        </tr>
        </tbody> 
      </table>
    </td>
  </tr>
  <tr> 
    <td colspan="2" height="26"> 
      <div align="center">      </div>
    </td>
  </tr>
</table>
    
    <!--#INCLUDE FILE="footer.asp"-->
  </div>