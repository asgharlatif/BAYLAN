<%@ Language=VBScript %>
<%
 Response.Buffer = True
 Response.ExpiresAbsolute = Now() - 1
 Response.Expires = 0
 Response.CacheControl = "no-cache"

if Request.Cookies("isLoggedinAs")= "" and  Request.Cookies("isWebmaster")= "" then
Response.Cookies("isloggedinas").expires=#december, 31 1999 00:00:00#
Response.Cookies("isWebmaster").expires=#december, 31 1999 00:00:00#
%>
<link rel="stylesheet" href="admin/arshilogin.css" type="text/css">


<p align="center"> <img src="aimages/spacer.gif" width="450" height="1"> <br>
<br>
</p>
<p align="center"><b><font face="Verdana" size="5" color="#336699">
Page Contents have Expired!</font></b></p>
<p align="center">
<br>
  <img src="aimages/spacer.gif" width="350" height="1"> </p>
<table width="93%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="53%"> 
      <p align="right"><font size="1" face="Verdana">&nbsp;Would you like to </font>
    </td>
    <td width="2%"><font color="#CC0000"><b><font size="1" face="Verdana"></font></b></font></td>
    <td width="45%"><font color="#CC0000"><b><font size="1" face="Verdana"><a href="admin/Default.htm">Login</a></font></b></font></td>
  </tr>
</table>
<div align="center">
<br>
  <img src="aimages/spacer.gif" width="350" height="1"><br>
</div>
<%else

Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3


mitemno=request("itemno")
set conn = server.CreateObject ("ADODB.Connection")
set rsData = server.CreateObject("ADODB.Recordset")
set rsDriver = server.CreateObject("ADODB.Recordset")
set rsVQty = server.CreateObject("ADODB.Recordset")
%>
<!--#INCLUDE FILE="dbconnect.asp"-->
<%
conn.Open con 

SQLtxt3="select * from avAllData where productid='"&mItemno&"'" 
	rsData.CursorLocation = adUseClient
	rsData.Open SQLtxt3,conn,adOpenForwardOnly,adLockReadOnly,adcmdtext
	set rsData.ActiveConnection = nothing
pid=rsdata("productid")
SQLtxt4="select name,driverURL from adrivers where productID='"&pid&"'" 
	rsDriver.CursorLocation = adUseClient
	rsDriver.Open SQLtxt4,conn,adOpenForwardOnly,adLockReadOnly,adcmdtext
    set rsDriver.ActiveConnection = nothing
	
Stxt="select * from avpqtyUnit where productid='"&pid&"'" 
	rsVQty.CursorLocation = adUseClient
	rsVQty.Open Stxt,conn,adOpenForwardOnly,adLockReadOnly,adcmdtext
	set rsVQty.ActiveConnection = nothing

%>	
<html>
<head>
<title>Preview Products</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="acss/standard.css">
</head>

<body bgcolor="#FFFFFF">
<table border=0 width=86%>
<tr>
<td class=text width=86%>Detail Preview:</td>
<td class=text width=14%>Short Preview</td>
</tr></table>
  <table width="86%" border="1" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="58%"> 
        <table width="440" border="1" cellspacing="0" cellpadding="0" height="385" bordercolorlight="#FFD700" bordercolordark="#FFD700" style="border-collapse: collapse" bordercolor="#111111">
          <tr valign="top"> 
            
          <td colspan="3" height="35" width="400" style="border-left-style: solid; border-left-width: 1; border-right-style: solid; border-right-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: none; border-bottom-width: medium"><span class="text2">Prod 
            ID :</span><span class="text"> <%=rsdata("itemcode")%> <span class="text2">Brand:</span>&nbsp;<%=rsdata("brand")%><span class="text2">&nbsp;Warranty 
            : </span>&nbsp;<%=rsdata("Warranty")%><br>
            <span class="text-heading"><%=rsdata("Description")%> &nbsp;</span> 
          </td>
          </tr>
          <tr> 
            
          <td align="center" width="133" rowspan="2" style="border-left-style: solid; border-left-width: 1; border-right-style: none; border-right-width: medium; border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium"> 
            <div align="center"><br>
              <%if (rsdata("picturelarge"))  <> "" then%>
              <img src="<%=rsdata("picturelarge")%>"> 
              <%else%>
              <img src="aimages/NOIMAGE1.GIF" border=0> 
              <%end if%>
              </div>
            
              <%if (rsdata("datasheet")) <> "" then%>
              <a class="text" href = "<%=rsData("datasheet")%>">Datasheet</a>
              <%end if%>
             
			  <%if  not rsdriver.EOF then%>             
              <a class="text" href = "<%=rsdriver("driverurl")%>">Download Driver</a> 
              <%end if%>
            <td align="left" width="100" rowspan="2" valign="top" style="border-style: none; border-width: medium"><span class="text-heading"><img src="<%=rsdata("image")%>" border=0></span></td>
            
          <td valign="bottom" width="150" height="185" style="border-left-style: none; border-left-width: medium; border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium"> 
            <br>
              <span class="text2">Price :</span> <span class="text3">Rs.<%=formatnumber(rsdata("retail"),2)%></span><br>
              <br>
              <span class="text2">Qty : </span> 
              <input type="text" name="mqty" size="6" maxlength="6" value=1 class="input1">
              <br>
              <br>
            </td>
          </tr>
          <tr> 
            <td width="160" style="border-left-style: none; border-left-width: medium; border-right-style: solid; border-right-width: 1; border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium"> 
              <input id=submit1 name="addincart" type=submit value="Add to Cart" 
              style="BACKGROUND-COLOR: gold; BORDER-BOTTOM-COLOR: gold; BORDER-LEFT-COLOR: gold; BORDER-RIGHT-COLOR: gold; BORDER-TOP-COLOR: gold; HEIGHT: 22px; WIDTH: 72px">
              <input id=checkout name="checkout" type=submit value="Checkout" 
              style="BACKGROUND-COLOR: gold; BORDER-BOTTOM-COLOR: gold; BORDER-LEFT-COLOR: gold; BORDER-RIGHT-COLOR: gold; BORDER-TOP-COLOR: gold; HEIGHT: 22px; WIDTH: 72px">&nbsp;
            </td>
          </tr>
          
            <%
ldesc=rsdata("longdescription")
Function GetHTMLFile(strFileName)
Dim fso, file, FName
FName = Server.MapPath(ldesc)
Set fso = Server.CreateObject ("Scripting.FileSystemObject")
Set file=fso.OpenTextFile(FName,1)
GetHTMLFile = Trim(file.readAll())
End Function%>
			<tr valign="top"> 
			<td colspan="3" width="400" style="border-bottom-style: solid; border-bottom-width: 1">
			<span class="text2">Description:</span><br>
             <%if ldesc <> "" and  ldesc <> "apimages/" then%>
                 <span class="text4"><%Response.Write GetHTMLFile(request("ldesc"))%></span></td> 
				     <%end if%>
			            </tr>
        </table>
</td>
      
    <td width="42%" class="text3" align="center"> <%=rsData("item")%><br>
        
      <table border="0" cellspacing="0" cellpadding="0" height="222" width="133">
        <tr> 
                <td height="2" colspan="2"><img src="aimages/spacer.gif" width="117" height="1"></td>
              </tr>
              <tr> 
                <%if not isnull(rsdata("image"))  then%>
                
          <td height="2" align="right" colspan="2"> <img src=<%=rsData("image")%>></td>
                <%else%>
                <td height="2" align="right" colspan="2">&nbsp;</td>
                <%end if%>
              </tr>
              <tr> 
                
          <td height="2" colspan="2" class="text3" align="center">&nbsp;</td>
              </tr>
			  <tr> 
                <td height="2" colspan="2" class="text3" align="center"><%=rsData("itemcode")%></td>
              </tr>
          <%if  (rsdata("picturethumb"))  <> "" then%>
          <tr align="center"> 
            <td height="28" colspan="2"> <a href="#"> <img src=<%=rsData("picturethumb")%>  border=0> 
              </a> </td>
          </tr>
          <%else%> 
          <tr align="center"> 
            <td height="28" colspan="2"> <a href="#"> <img src="aimages/noimage.gif" border=0></a></td>
          </tr>
          <%end if%> 
          <tr> 
            <td height="18" colspan="2" class=desc> <%=left(rsData("description"),100)%> 
            </td>
          </tr>
          <tr> 
            <td height="29"> <a href="#"> <img src="aimages/buynow.gif" border=0></a></td>
                        <td height="29" class="text">Price<br>
            Rs.<%=formatnumber(rsdata("retail"),2)%></td>
            </tr>
            </table>
      </td>
    </tr>
  </table>
  
<br>
<table width="86%" border="1" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="20%"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#FF0000">Thumbnail</font></td>
    <td width="30%"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#0000FF"><%=rsdata("picturethumb")%></font></td>
    <td width="20%"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#FF0000"> 
      Image</font></td>
    <td width="30%"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#0000FF"><%=rsdata("picturelarge")%></font></td>
  </tr>
  <tr> 
    <td width="203"><font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#FF0000">Feature 
      Pic Name</font></td>
    <td width="159"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#0000FF"><%=rsdata("picturefeature")%></font></td>
    <td width="106"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#FF0000">Long Description</font></td>
    <td width="168"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#0000FF"><%=rsdata("longdescription")%></font></td>
  </tr>
  <tr> 
    <td width="203"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#FF0000">DataSheet 
      Name</font></td>
    <td width="159"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#0000FF"><%=rsdata("datasheet")%></font></td>
     <td width="106"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#FF0000">Driver 
      Name</font></td>     
     <%if not rsdriver.EOF  then%>
       <td width="168"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#0000FF"><%=rsdriver("driverurl")%></font></td>
                <%else%>
                <td width="168"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#0000FF"></font></td>
                <%end if%>
    
    
    
    

  </tr>
  <tr> 
    <td width="203"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#FF0000">Taxable</font></td>
    <td width="159"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#0000FF"><%=rsdata("taxable")%></font></td>
    <td width="106"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#FF0000">Shipping 
      Charge</font></td>
    <td width="168"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#0000FF"><%=rsdata("shippingc")%></font></td>
  </tr>
  <tr> 
    <td width="203"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#FF0000">Weight</font></td>
    <td width="159"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#0000FF"><%=rsdata("Weight")%></font></td>
    <td width="106"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#FF0000">Status</font></td>
    <td width="168"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#0000FF"><%=rsdata("status")%></font></td>
  </tr>
  <tr> 
    <td width="203"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#FF0000">Main 
      Page Show Price</font></td>
    <td width="159"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#0000FF"><%=rsdata("showmp")%></font></td>
    <td width="106"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#FF0000">Sort ID#</font></td>
    <td width="168"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#0000FF"><%=rsdata("sortby")%></font></td>
  </tr>
</table>
<table width="86%" border="1" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="20%"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#FF0000">Group</font></td>
    <td width="80%"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#0000FF"><%=rsdata("groupdesc")%></font></td>
  </tr>
  <tr> 
    <td width="20%"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#FF0000">Category</font></td>
    <td width="80%"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#0000FF"><%=rsdata("category")%></font></td>
  </tr>
</table>
<table width="86%" border="1" cellspacing="0" cellpadding="0">
  <tr> 
      <%
  '  Response.Write stxt
  '  "--------" & price__ & "<br>"
    %>
    <td width="22%" CLASS="TEXT"><font color="#FF0000">Group Price</font></td>
    <td width="13%" align="center" class="text" ><%=rsVqty("qtyA")%>&nbsp;</td>
    <td width="13%" align="center" class="text" Color=Red><%=rsVqty("qtyB")%>&nbsp;</td>
    <td width="13%" align="center" class="text" Color=Red><%=rsVqty("qtyC")%>&nbsp;</td>
    <td width="13%" align="center" class="text" Color=Red><%=rsVqty("qtyD")%>&nbsp;</td>
    <td width="13%" align="center" class="text" Color=Red><%=rsVqty("qtyE")%></td>
    <td width="13%" align="center" class="text" Color=Red><font color="#FF0000">Retail</font>&nbsp;</td>
  </tr>
  <tr> 
    <td width="22%" class="text"><font color="#FF0000">Price</font></td>
  
    <td width="13%" class="text"><%=trim(rsdata("pgA"))%>&nbsp;</td>
    <td width="13%" class="text"><%=trim(rsdata("pgB"))%>&nbsp;</td>
    <td width="13%" class="text"><%=trim(rsdata("pgc"))%>&nbsp;</td>
    <td width="13%" class="text"><%=trim(rsdata("pgd"))%>&nbsp;</td>
    <td width="13%" class="text"><%=trim(rsdata("pge"))%></td>
    <td width="13%" class="text">&nbsp;<%=trim(rsdata("retail"))%></td>
  </tr>
  <tr> 
    <td width="22%" class="text"><font color="#FF0000">you save/unit</font></td>
    <td width="13%" class="text">-Nil-</td>
    <td width="13%" class="text"><%=cstr(trim(rsdata("pgA")))-cstr(trim(rsdata("pgb")))%>&nbsp;</td>
    <td width="13%" class="text"><%=cstr(trim(rsdata("pgA")))-cstr(trim(rsdata("pgc")))%>&nbsp;</td>
    <td width="13%" class="text"><%=cstr(trim(rsdata("pgA")))-cstr(trim(rsdata("pgd")))%>&nbsp;</td>
    <td width="13%" class="text"><%=cstr(trim(rsdata("pgA")))-cstr(trim(rsdata("pge")))%></td>
    <td width="13%" class="text">&nbsp;</td>
  </tr>
</table>
</body>
</html>
<%end if%>