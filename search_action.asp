<%@ Language=VBScript %>
<%
p_usetext=request("p_usetext")
p_text=cstr(request("p_text"))
p_text=ucase(p_text)
p_useprice=request("p_useprice")
p_price=request("p_price")
p_usecat=request("p_usecat")
p_cat=request("p_cat")

p_textWhere = ""
p_priceWhere = ""
p_catWhere = ""

if p_usetext= "yes" then
	p_textwhere = " (Upper(item) like '%"& p_text &"%' "
	p_textwhere=p_textwhere & " or upper(longdescription) like '%"& p_text &"%')"
end if


if p_useprice="yes" and p_price <> "all" then
	if p_usetext= "" then
    p_pricewhere= " retail < " &p_price
	else
    p_pricewhere= " and retail < " &p_price
	end if  
end if

if p_useCat = "yes" then
	if p_usetext= "" and p_useprice=""  then
	p_catwhere=" catcode= " &p_cat
	else
	p_catwhere=" and catcode= " &p_cat
	end if   
end if
      	
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
set conn = server.CreateObject ("ADODB.Connection")
set rsMydata = server.CreateObject("ADODB.Recordset")


%>
<!--#INCLUDE FILE="dbconnect.asp"-->
<%
conn.Open con 

if p_usetext ="" and p_useprice = "" and p_usecat="" then%>
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
sqltext=" select * from  valldata where " & p_textwhere & p_pricewhere & p_catwhere 
'Response.Write sqltext
rsMydata.CursorLocation = adUseClient
rsMydata.Open SQLtext,conn,adOpenForwardOnly,adLockReadOnly,adcmdtext
set rsMydata.ActiveConnection = nothing
'set rsMydata=conn.Execute(sqltext)
'Response.Write sqltext
%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="admin/arshilogin.css">
</head>
<h1><font face="Arial, Verdana" size="+7" color="#F4F4F4">Searh Results:</font></h1>

<table width="100%" cellpadding="0" cellspacing="0" border="0" >
  <tr bgcolor="#CCCCFF"> 
    <td width="1%">&nbsp;</td>
    <td width="3%" class="text">S#</td>
    <td width="25%" class="text">Product Name</td>
    <td width="14%" class="text" align="right">Price (Rs.)</td>
    <td width="18%" class="text" align="center">Data Sheet</td>
    <td width="23%" class="text" align="center">Driver</td>
    <td width="15%" class="text"><span class="text"></span></td>
  </tr>
<%
if not rsmydata.EOF then
sno=1
while not rsMydata.EOF
aa=sno/2
zaf=Instr(aa,".")
'Response.Write aa&"-"&zaf
 a=rsMydata("datasheet")
    if not isnull(a)  then
    tmpLng1 = instr(a,"/")
    dsheet=mid(a,tmpLng1+1,25)
    end if
    
 b=rsMydata("driverURl")
    if not isnull(b)  then
    tmpLng1 = instr(b,"/")
    drvr=mid(b,tmpLng1+1,25)
    end if    

if zaf=0 then%>
  <tr bgcolor="#E8F9FF"> 
<%else%>  
<tr>
<%end if%>
    <td width="1%" height="10">&nbsp;</td>
    <td width="3%" class="text" height="10"><%=sno%></td>
    <td width="25%" class="text" height="10">
 <a class="links1" href="page2.asp?a=<%="zaf"%>&Mcode=<%=rsmydata("catID")%>&pID=<%=rsMydata("productID")%>&name=<%=rsMydata("Item")%>&price=<%=rsMydata("retail")%>&pic=<%=rsMydata("picturelarge")%>&lDesc=<%=rsMydata("longdescription")%>&stimage=<%=rsMydata("image")%>&mbrand=<%=rsMydata("brand")%>&mwar=<%=rsMydata("warranty")%>&mdsheet=<%=rsMydata("datasheet")%>"> 
<%=rsmydata("item")%>    
    </a></td>
    <td width="14%" align="right" class="text" height="10"><%=formatnumber(rsmydata("retail"),2)%></td>
    <td width="18%" class="text" height="10" align=right>
    <%if dsheet <> "" then%>
    <A class="links1" href="ftp://arshi/<%=rsmydata("datasheet")%>"><%=dsheet%></a>
    <%end if%>
    </td>
    <td width="23%" height="10" align=right>
    <%if drvr <> "" then%>
    <A class="links1" href="ftp://arshi/<%=rsMydata("driverurl")%>"><%=drvr%></a>
    <span class=text>Size : <%=rsmydata("filesize")%></span>
    <%end if%>
    </td>
    <td width="15%" class="text" height="10" align=center> 
      <a class="links1" href="cart2.asp?xpid=<%=rsmydata("productID")%>&Mcode=<%=rsmydata("catID")%>&mScat=<%=rsmydata("catCode")%>&bid=<%=cartID%>&mqty=1">Add to Cart</a>
    </td>
  </tr>
<%sno=sno+1
rsMydata.MoveNext
wend
else
Response.Write "No Record Found"
end if
end if %>
 <tr  bgcolor="#CCCCFF"> 
    <td colspan="7" height="2"></td>
  </tr>
</table>
</html>