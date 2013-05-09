<%@ Language=VBScript %>

<link href="acss/newstyle.css" rel="stylesheet" type="text/css" />
<div class="Private_Policy">


<!--#INCLUDE FILE="header.asp"-->
<br>
<%
 if Request.Cookies("isDealer") <> "" then
 xmsg="You are Logged in as " & Request.Cookies("typename") & " : [" & Request.Cookies("dealer_id") & "]"
 end if
%>

<html><head><title>BBL Computrade Privacy Policy Page</title> 
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
	  <br>
      
      <p align="justify">&nbsp;       <font face=ARIAL size=2> </font></td> 
      </table>
	  <table width="50%" border="0" cellspacing="0" cellpadding="0" height="1" background="aimages/dot_grey1.gif" align="center">
        <tr> 
     
        </tr>
      </table>
      <font face=ARIAL size=2> </font>      <font face=ARIAL size=2></font>
<td colspan="3" height="1981"> 
      <div align="center"> 
        
      </div>
      
</td>
  </tr>
</table>
    
  <div align="center"> 
    <table width="100%" border="0" cellspacing="0" cellpadding="0" height="1" background="aimages/dot_grey1.gif">
    <tr> 
   
    </tr>
    </table>
   
   <!--#INCLUDE FILE="footer.asp"-->


  
</div>