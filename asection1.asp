<%
set conn = server.CreateObject ("ADODB.Connection")
set rsMopt = server.CreateObject("ADODB.Recordset")
set rsMcat = server.CreateObject("ADODB.Recordset")
set rsScat = server.CreateObject("ADODB.Recordset")
set rsVprice = server.CreateObject("ADODB.Recordset")
%>
<!--#INCLUDE FILE="dbconnect.asp"-->
<%
conn.Open con 

Stxt="SELECT   * FROM  Amainopt WHERE     msection = 'sectionA'"
rsMopt.CursorLocation = adUseClient
rsMopt.Open Stxt,conn,adOpenForwardOnly,adLockReadOnly,adcmdtext
mSec=rsMopt("catcode")
set rsMopt.ActiveConnection = nothing
%>
<html>
<head>
<title>Section Page 1</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">

<link rel="stylesheet" href="acss/standard.css" type="text/css">
</head>

<body bgcolor="#FFFFFF" text="#000000">
<table style="FONT-SIZE: 8pt; FONT-FAMILY: Verdana" 
                  cellspacing=1 width="28%" border=0 cellpadding="2">
  <tbody> 
  <%

SQLtxt="select * from ASubcat where catcode='"&mSec&"' " 
rsScat.CursorLocation = adUseClient
rsScat.Open SQLtxt,conn,adOpenForwardOnly,adLockReadOnly,adcmdtext
rsScat.ActiveConnection = nothing
bb=rsScat("catcode")
%>
  <tr bgcolor="#006699" align="center"> 
    <td colspan=2><strong><font style="FONT-SIZE: 8pt" face="Verdana" color="#ffffff" size="2"> 
      <img src="aimages/spacer.gif" width="180" height="1"><br>
      <%=rsscat("category")%> </font></strong></td>
  </tr>
  <tr bgcolor="#FF9900"> 
    <td width="66%" bgcolor="#0099CC" align="center">Item Name</td>
    <td width="40%" bgcolor="#0099CC" align="center">Price</td>
<%

sqltxt2="select *,[dbo].EditPrice (retail * " & price &" ) as retail,[dbo].EditPrice (pga * " & price &" ) as pga from AvPrice  where showmp= 'Yes' and category='"& rsscat("category")& "'" 
set rsvprice=conn.Execute(sqltxt2)

if rsvprice.EOF=false then
a1=rsvprice("itemcode")
end if

for i=1 to 5
                if rsvprice.EOF=false then
if Request.Cookies("isDealer") <> "" AND Request.Cookies("TypeName") ="Reseller"  OR Request.Cookies("TypeName") ="Cable Net" then 
%>
  <tr> 
    <td width="66%" class="text4"><a href="apage2.asp?Mdata=<%=rsvprice("catcode")%>&head1=<%=rsvprice("groupdesc")%>&head2=<%=rsvprice("category")%>&mcode=<%=rsvprice("CatId")%>"><%=rsVprice("item")%></a></td>
    <td width="34%" align="right" class="text4" ><%=formatnumber(trim(rsVprice("pga")) ,2)%></td>
  </tr>
<%
else if Request.Cookies("isDealer") <> "" then
%>
  <tr> 
    <td width="66%" class="text4" ><a href="apage2.asp?Mdata=<%=rsvprice("catcode")%>&head1=<%=rsvprice("groupdesc")%>&head2=<%=rsvprice("category")%>&mcode=<%=rsvprice("CatId")%>"><%=rsVprice("item")%></a></td>
    <td width="34%" align="right"class="text4" ><%=formatnumber(trim(rsVprice("retail")) ,2)%></td>
  </tr>
<%  

end if
end if 
xcode=rsvprice("itemcode")
rsVprice.MoveNext
                end if      
next  %>
</table>
</body>
</html>
