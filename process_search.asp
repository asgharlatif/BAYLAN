<%@ Language=VBScript %>
<!--#INCLUDE FILE="header.asp"-->
<%

'Response.Write "<br>price ratio = " & session("price__") & "<br>"

p_usetext=request("p_usetext")
p_text=cstr(request("p_text"))
p_text=ucase(p_text)
p_usemcat=request("p_usemcat")
p_mcat=request("p_mcat")
p_usecat=request("p_usecat")
p_cat=request("p_cat")
p_usebrand=request("p_usebrand")
p_brand=request("p_brand")

p_textWhere = ""
p_mcatWhere = ""
p_catWhere = ""
p_brandWhere= ""
if p_usetext= "yes" then
	p_textwhere = " (Upper(item) like '%"&p_text&"%' "
	p_textwhere=p_textwhere & " or upper(description) like '%"&p_text&"%'" & "or upper(groupdesc) like '%"&p_text&"%'" &  "or upper(Category) like '%"&p_text&"%')"
end if


if p_usemcat="yes" then
	if p_usetext= "" then
    p_mcatwhere= " catID=  " &p_mcat
	else
    p_mcatwhere= " and catID= " &p_mcat
	end if  
end if

if p_useCat = "yes" then
	if p_usetext= "" and p_usemcat=""  then
	p_catwhere=" catcode= " &p_cat
	else
	p_catwhere=" and catcode= " &p_cat
	end if   
end if

if p_usebrand = "yes" then
	if p_usetext= "" and p_usemcat="" and p_usecat="" then
	p_brandwhere=" brandid= " &p_brand
	else
	p_brandwhere=" and brandid= " &p_brand
	end if   
end if

Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3

set conn = server.CreateObject ("ADODB.Connection")
set rsMydata = server.CreateObject("ADODB.Recordset")


set rsDriver = server.CreateObject("ADODB.Recordset")
if Request.Cookies("isDealer") <> "" then
 xmsg="You are Logged in as " & Request.Cookies("typename") & " : [" & Request.Cookies("dealer_id") & "]"
 end if%>


<!--#INCLUDE FILE="dbconnect.asp"-->
<%
conn.Open con 
%>
<html><head><title>Baylan Technology Inc. Product Search Page</title> 

  <tr> 
  <td colspan="2"  valign="top" width="100%" ><tr><td width="25%" valign="bottom" align="center">&nbsp;</td>
  <img src="aimages/spacer.gif" width="1" height="1"></tr>
 <table width="100%" border="0">
  <tr>
  <td><div align="center"> 
  <%  
   Dim objAd
   Set objAd = Server.CreateObject("MSWC.AdRotator") 'Create the component
   Response.Write  objAd.GetAdvertisement ("/banners/adrot.txt") 
   Set objAd = Nothing 'Destroy the object
   %>
</div>
 </td></tr></table>
  <tr></tr>
</table> 



<%if  p_usetext ="" and p_usemcat = "" and p_usecat="" and p_usebrand="" then%>
<head>
	<link rel="stylesheet" href="admin/arshilogin.css">

	<h3 align="center"><font face="Verdana, Arial" color="#FF0033"><br>
	<img src="admin/images/spacer.gif" width="400" height="1"><br>
	</font></h3>
	<h3 align="center"><font face="Verdana, Arial" color="#FF0033">
	You did not selected any option!</font> 
	</h3>
	<div align="center"><font face="Verdana, Arial" color="#FF0033"><img src="admin/images/spacer.gif" width="400" height="1"></font><br>
	<br>
	<a href="javascript:history.back()"><font size="1" face="Verdana, Arial">TRY 
	AGAIN</font></a> </div>
<%else

if Request.Cookies("isDealer") <> "" or Request.Cookies("isMember") <> "" then
sqltext=" select *,[dbo].EditPrice (retail * " & session("price__") &" ) as retail,[dbo].EditPrice (pga * " & session("price__") &" ) as pga  from  valldata where statuscode != '08' and" & p_textwhere & p_mcatwhere & p_catwhere & p_brandwhere
'Response.Write sqltext
else

sqltext=" select * from  valldata where statuscode != '08' and " & p_textwhere & p_mcatwhere & p_catwhere & p_brandwhere
end if
rsMydata.CursorLocation = adUseClient
rsMydata.Open SQLtext,conn,adOpenForwardOnly,adLockReadOnly,adcmdtext
set rsMydata.ActiveConnection = nothing
'set rsMydata=conn.Execute(sqltext)
'Response.Write sqltext
if not rsmydata.EOF then
sno=1
%><div align=left>
<h1><font face="Arial, Verdana" size="+7" color="#F4F4F4">Searh Results:</font></h1>
</div>
<div align=right>
<a href="javascript:history.back()"><font size="1" face="Verdana, Arial">Back</font></a>
</div>

<table width="100%" cellpadding="0" cellspacing="0" border="0" >
<tr bgcolor="#CCCCFF" class="normtext">
<td><div align="center">S#</div></td>
<td><div align="left">Product Id</div></td>
<td><div align="left">Category</div></td>
<td><div align="left">Brand</div></td>
<td><div align="left">Product Name</div></td>
<% if request("isdealer") <> "" or request("ismember") <> "" then %>
<td><div align="center">Price (Rs.)</div></td>
    <%else%>
<td>&nbsp;</td>
    <%end if%>
<td><div align="center">DataSheet</div></td>
<td><div align="center">Drivers</div></td>
<td>Size</td>
<% if request("isdealer") <> "" or request("ismember") <> "" then %>
<td>&nbsp;</td>
</td>
 <%else%>
 <%end if%>
 <%
while not rsMydata.EOF
aa=sno/2
ars=Instr(aa,".")

 a=rsMydata("datasheet")
    if not isnull(a)  then
    tmpLng1 = instr(a,"/")
    dsheet=mid(a,tmpLng1+1,25)
    end if
    if not isnull(b)  then
    tmpLng1 = instr(b,"/")
    drvr=mid(b,tmpLng1+1,25)
    end if    
if ars=0 then%>

  <tr bgcolor="#E8F9FF"> 
    <%else%>
  <tr> 
    <%end if%>

<td align="right" class="text" height="20"><%=sno%>&nbsp;</td>
<td class="text" height="20"><a class="text" href="product.asp?pid=<%=rsmydata("productid")%>"><%=rsmydata("itemcode")%>&nbsp;<a/></td>
<td class="text" height="20"><%=rsmydata("category")%>&nbsp;</td>
<td class="text" height="20"><%=rsmydata("brand")%>&nbsp;</td>
<% if Request.Cookies("isDealer") <> "" AND Request.Cookies("TypeName") ="Reseller"  OR Request.Cookies("TypeName") ="Cable Net" then%>
	<td class="text" height="20"><a class="text" href="apage2.asp?a=<%="ars"%>&Mcode=<%=rsmydata("catID")%>&head1=<%=rsmydata("GroupDesc")%>&head2=<%=rsmydata("Category")%>&pID=<%=rsMydata("productID")%>&price=<%=rsMydata("pga")%>&pic=<%=rsMydata("picturelarge")%>&stimage=<%=rsMydata("image")%>&mbrand=<%=rsMydata("brand")%>&mwar=<%=rsMydata("warranty")%>&mdsheet=<%=rsMydata("datasheet")%>"> 
      <%=rsmydata("Description")%> </a></td>
<%else if Request.Cookies("isDealer") <> "" then%>
<td class="text" height="20"><a class="text" href="apage2.asp?a=<%="ars"%>&Mcode=<%=rsmydata("catID")%>&head1=<%=rsmydata("GroupDesc")%>&head2=<%=rsmydata("Category")%>&pID=<%=rsMydata("productID")%>&price=<%=rsMydata("retail")%>&pic=<%=rsMydata("picturelarge")%>&stimage=<%=rsMydata("image")%>&mbrand=<%=rsMydata("brand")%>&mwar=<%=rsMydata("warranty")%>&mdsheet=<%=rsMydata("datasheet")%>"> 
      <%=rsmydata("Description")%> </a></td>
	   <%else%>
<td class="text" height="20"> <a class="text" href="apage2.asp?a=<%="ars"%>&Mcode=<%=rsmydata("catID")%>&head1=<%=rsmydata("GroupDesc")%>&head2=<%=rsmydata("Category")%>&pID=<%=rsMydata("productID")%>&pic=<%=rsMydata("picturelarge")%>&stimage=<%=rsMydata("image")%>&mbrand=<%=rsMydata("brand")%>&mwar=<%=rsMydata("warranty")%>&mdsheet=<%=rsMydata("datasheet")%>"> 
      <%=rsmydata("Description")%> </a></td>
    <%end if%>
	<%end if%>
 <%if Request.Cookies("isDealer") <> "" AND Request.Cookies("TypeName") ="Reseller"  OR Request.Cookies("TypeName") ="Cable Net" then%>
    <% if   rsmydata("statuscode") = "08" then%>
	            <td>&nbsp;</td>
			     <%else%>
	<td  align="right" class="text" height="20"><%=formatnumber(trim(rsmydata("pga")),2) %></td>
     <%end if%>
	<% else if Request.Cookies("isDealer") <> "" then %>
     <% if   rsmydata("statuscode") = "08" then%>
	            <td>&nbsp;</td>
			     <%else%>
	<td  align="right" class="text" height="20"><%=formatnumber(rsmydata("retail"),2)%></td>
    <%end if%>
	<%else%>
    <td>&nbsp;</td>
    <%end if%>
	<%end if%>
<td class="text" height="20"><div align="center">
 <%if rsmydata("datasheet") <> "" and  trim(rsmydata("datasheet")) <> "PIMAGES\" then%>
	  <A class="textl" href= "<%=(rsmydata("datasheet"))%>" target="_blank">Download</a> 
      <%end if%>
</div></td>
<td class="text" height="20">
    <div align="center">
       <%if rsmydata("driver") <> "" and  trim(rsmydata("driver")) <> "PIMAGES\" then%>
	  <A class="textl" href= "<%=(rsmydata("driver"))%>">Download&nbsp;&nbsp;</a> 
    </div></td>
	 <td align="right" height="10" align=left class="text">
	  <% dim fso, file, fname
	  fname = Server.Mappath(rsmydata("driver"))
	  set fso = server.CreateObject("Scripting.FileSystemObject")  
      Set file = fso.Getfile(fname)
	  Response.Write (formatnumber(file.Size/1024,0) & " KB <BR>") 
  	  Set file = Nothing  
      Set fso = Nothing  
	  else%>
    <td>&nbsp;</td>
	     	  <%end if%>
	     </td>
<% if request("isdealer") <> "" or request("ismember") <> "" then %>
<% if   rsmydata("statuscode") = "08" then%>
	        			     <%else%>
 <td class="text" align="center" height="20"><a class="text" href="acart2.asp?xpid=<%=rsmydata("productID")%>&Mcode=<%=rsmydata("catID")%>&mScat=<%=rsmydata("catCode")%>&bid=<%=cartida%>&mqty=1">Add to Cart</a></td>
 <%end if%>
 <%else%>


    <%end if%>
<%sno=sno+1
rsMydata.MoveNext
wend
else
Response.Write "<br><br><hr><br>"
Response.Write "<div align=center><font color=red size=2><b>No Record Found</b></font></div>"
Response.Write "<br><br><hr><br>"
end if
end if %>
<tr bgcolor="#CCCCFF"> 
  </tr>
</table>
	
<br>
 <div align="center"></div>
  <table border="0" cellspacing="0" cellpadding="0" width="100%">
    <hr color=#000000 size=1 width="650" align="center" noShade>
    <tr> 
    <td width="100%" height="21" id=msviRegionGradient2 
    style="FILTER: progid:DXImageTransform.Microsoft.Gradient(startColorStr='#FFFFFF', endColorStr='#6487DB', gradientType='1')">	<p align="center" class="font">Baylan Technology Inc. Copyright &copy; 2010</p></td>
	
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
</html>