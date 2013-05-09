<%@ Language=VBScript %>
<%
Response.Buffer = True
 Response.ExpiresAbsolute = Now() - 1
 Response.Expires = 0
 Response.CacheControl = "no-cache"
	session.Timeout = 2
Response.Cookies("isMember").expires=#december, 31 1999 00:00:00#

Response.Cookies("dcategory")=""
Response.Cookies("EmailT") = ""
Response.Cookies("MemName")=""
%>
<p align="center">
<img src="admin/images/spacer.gif" width="600" height="1">
<br>
<br>
</p>
<%
session.Abandon 
%>
<p align="center"><b><font face="Verdana" size="5" color="#336699">
You Have Been Logout successfuly!</font></b></p>
<p align="center">
<br>
 <img src="admin/images/spacer.gif" width="600" height="1"> 
</p>
<div align="center"><font size="1" face="Verdana">Go back to <a href="computer.asp">Home Page</a></font></div>
</p>
<table width="94%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="50%">&nbsp;
     </td>
    <td width="7%"><font color="#CC0000"><b><font size="1" face="Verdana"></font></b></font></td>
    <td width="43%">&nbsp;
    </td>
  </tr>
</table>
<div align="center">
<br>

<img src="../aimages/spacer.gif" width="600" height="1"><br>
</div>

