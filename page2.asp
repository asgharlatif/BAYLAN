<%@ Language=VBScript %>
<%
'on error resume next
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
'Response.Write "=____________=" &  trim(Request.Cookies("isdealer"))

mcd=request("mcd") ' data from cart2

if mcd <> "" then
myrange=split(mcd,"-")

for i =0 to ubound(myrange)
	p1=myrange(0)  'mcode to view same selected catgory as before o ca
	
	P2=myrange(1)  'subcategory id

	P3=myrange(2)  'total
	P4=myrange(3)  'no of items
	P5=myrange(4)  'Basket ID

next
session("scode")=p1
scode=p1   'main catid
xdata=p2   'sbcat id
session("tot")=p3
session("nitems")=p4
session("cartID")=p5
else
scode=request("mcode")
xData=request("mData")

end if


if request("mData") <> ""  then
mData=request("mData")
session("mData_")=request("mData")
xData=mData
end if

if request("mData") = ""  and Request.QueryString("pgt") = "" then 
mData=session("mData_")
xData=mData
end if


a=request("a")

'------------Parameters for order procedure-----------
pid=request("pid")				'product id
price=request ("price")		'price

'----------------------------------------------------

set conn = server.CreateObject ("ADODB.Connection")
set rsMain = server.CreateObject("ADODB.Recordset")
set rsSub = server.CreateObject("ADODB.Recordset")
set rsData = server.CreateObject("ADODB.Recordset")
set rsDriver = server.CreateObject("ADODB.Recordset")


set RScartid = server.CreateObject("ADODB.Recordset")
set RScartitems = server.CreateObject("ADODB.Recordset")
set rsVQty = server.CreateObject("ADODB.Recordset")

%>
<!--#INCLUDE FILE="dbconnect.asp"-->
<%
conn.Open con 

if trim(Request.Cookies("isDealer")) <> "" or trim(Request.Cookies("isMember")) <> "" then
SQLtxt_="select * from DCategory where category='"& Request.Cookies("dcategory") &"'" 
set RsPricing=server.CreateObject("adodb.recordset")
RsPricing.CursorLocation = adUseClient
RsPricing.Open SQLtxt_,conn,adOpenForwardOnly,adLockReadOnly,adcmdtext
set RsPricing.ActiveConnection = nothing
price=RsPricing("mration")
price__=RsPricing("mration")
session("price__")=RsPricing("mration")
end if


SQLtxt="select * from vmainCAT order by sort_it" 
rsMain.CursorLocation = adUseClient
rsMain.Open SQLtxt,conn,adOpenForwardOnly,adLockReadOnly,adcmdtext
set rsMain.ActiveConnection = nothing

SQLtxt2="select * from vSUBCat where CatID='"&Scode&"' " 
rsSub.CursorLocation = adUseClient
rsSub.Open SQLtxt2,conn,adOpenForwardOnly,adLockReadOnly,adcmdtext
set rsSub.ActiveConnection = nothing
	if not rssub.eof then
		msid2=rssub("catcode")
	end if



if a = "" then

if xData <> "" then
if Request.Cookies("isDealer") <> "" or Request.Cookies("isMember") <> "" then
	SQLtxt3="select productid,image,itemcode,picturethumb,description,[dbo].Editprice(retail * "& price &") as retail,[dbo].Editprice(pga * "& price &") as pga,statuscode from vAllData where Catcode='"&xData&"' and statuscode != '08' " 
else
SQLtxt3="select productid,image,itemcode,picturethumb,description,retail,pga,statuscode from vAllData where Catcode='"&xData&"' and statuscode != '08' " 	
end if
	rsData.CursorLocation = adUseClient
	rsData.Open SQLtxt3,conn,adOpenForwardOnly,adLockReadOnly,adcmdtext
	set rsData.ActiveConnection = nothing
    
               mloop1=rsdata.RecordCount/3
               mloop=int(rsdata.RecordCount/3)
               if mloop1 > mloop then
				mloop3=mloop+1
               end if


				if rsdata.EOF=false then
				if Request.Cookies("isDealer") <> "" then
					mprice=rsdata("pga")
				End if
				if Request.Cookies("isMember") <> "" then
					mprice=rsdata("retail")
				End if
                end if  
else
'	SQLtxt3="select * from vAllData where Catid='"&scode&"'" 
if Request.Cookies("isDealer") <> "" or Request.Cookies("isMember") <> "" then
	SQLtxt3="select productid,image,itemcode,picturethumb,description,[dbo].Editprice(retail * "& price &") as retail,[dbo].Editprice(pga * "& price &") as pga,statuscode from vAllData where Catid='"&scode&"' and statuscode != '08'" 
else
SQLtxt3="select productid,image,itemcode,picturethumb,description,retail,pga,statuscode from vAllData where Catid='"&scode&"' and statuscode != '08'" 
end if
	rsData.CursorLocation = adUseClient
	rsData.Open SQLtxt3,conn,adOpenForwardOnly,adLockReadOnly,adcmdtext
	set rsData.ActiveConnection = nothing
               mloop1=rsdata.RecordCount/3
               mloop=int(rsdata.RecordCount/3)
               if mloop1 > mloop then
				mloop3=mloop+1
              end if
end if

else

if Request.Cookies("isDealer") <> "" or Request.Cookies("isMember") <> "" then
	SQLtxt3="select *  ,[dbo].Editprice(pga * "& price &") as pga  from vAllData where productid='"&pid&"'" 
else
SQLtxt3="select  *  from vAllData where productid='"&pid&"'" 
end if
	'SQLtxt3="select productid,image,itemcode,picturethumb,description,retail,pga,statuscode from vAllData where Catid='"&scode&"' and statuscode != 'S008'" 
	rsData.CursorLocation = adUseClient
	rsData.Open SQLtxt3,conn,adOpenForwardOnly,adLockReadOnly,adcmdtext

pid=rsdata("productid")				'product id
name=rsdata("item")
price=request("price")		'price
pic=rsdata("picturelarge")				'large image
ldesc=rsdata("longdescription")		'long description
stimage=rsdata("image")	'status image like NEW
'mScat=rsdata("mscat")	'subcategory id for qty unit and items in subcategory
brand=rsdata("brand")	
warranty=rsdata("warranty")	
driver=rsdata("driver")
'mdriver=request("mdrvr")	
mdsheet=rsdata("datasheet")	

	'SQLtxt4="select name,driverURL from drivers where productID='"&pid&"'" 
	'rsDriver.CursorLocation = adUseClient
	'rsDriver.Open SQLtxt4,conn,adOpenForwardOnly,adLockReadOnly,adcmdtext
end if


%>
<html> <head> 
<meta http-equiv="Content-Language" content="en-us">
<title>Products</title> 
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<link rel="stylesheet" href="css/standard.css" type="text/css">

<script language=JavaScript>
var imgurl="images/";
</script>
<SCRIPT language=JavaScript src="scripts/image_init.js" type=text/javascript></SCRIPT>

<script language="JavaScript">
function popLogin() 
{
var coki = document.cookie;
var oklogin=0;

if (coki.indexOf("isDealer=") != -1 ){
var start = coki.indexOf("isDealer=");
alert ('Already Login as Dealer')
oklogin=1
}

if (coki.indexOf("isMember=") != -1 ){
var start = coki.indexOf("isDealer=");
alert ('Already Login as Member')
oklogin=1
}

if (oklogin==0){

if (start >= 0)
{
var start = coki.indexOf("=",start) + 1;
var end = coki.indexOf(";",start);

var value = unescape(coki.substring(start,end));
alert("You are Already Logged-in as DEALER");
}
else
{

if (navigator.appName == "Netscape")
{
if (navigator.appVersion.charAt(0) < 4) 
 {
        window.open('dLogin.htm','propop','screenX='+ x1 + ',screenY=' + y1 + ',width=400,height=240,resizable=no,toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no');
} 
else
{
var x1 = this.screenX + (this.outerWidth  - 385)/2;
var y1 = this.screenY + (this.outerHeight - 260)/2;
window.open('dLogin.htm','propop','screenX='+ x1 + ',screenY=' + y1 + ',width=400,height=260,resizable=no,toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no');
}
}
if (navigator.appName == "Microsoft Internet Explorer")					
{
var x1 =  (screen.width  - 385)/2;
var y1 =  (screen.height - 260)/2;

window.open('dLogin.htm','propop','screenX='+ x1 + ',screenY=' + y1 + ',left='+ x1 + ',top='+ y1 + ',width=450,height=250,resizable=no,toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no');
}
}
}
}


function popLogin_() 
{
var coki = document.cookie;
var oklogin=0;

if (coki.indexOf("isDealer=") != -1 ){
var start = coki.indexOf("isDealer=");
alert ('Already Login as Dealer')
oklogin=1
}

if (coki.indexOf("isMember=") != -1 ){
var start = coki.indexOf("isDealer=");
alert ('Already Login as Member')
oklogin=1
}

if (oklogin==0){

if (start >= 0)
{
var start = coki.indexOf("=",start) + 1;
var end = coki.indexOf(";",start);

var value = unescape(coki.substring(start,end));
alert("You are Already Logged-in as DEALER");
}
else
{

if (navigator.appName == "Netscape")
{
if (navigator.appVersion.charAt(0) < 4) 
 {
 window.open('mLogin.htm','propop','screenX='+ x1 + ',screenY=' + y1 + ',width=400,height=240,resizable=no,toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no');
} 
else
{
var x1 = this.screenX + (this.outerWidth  - 385)/2;
var y1 = this.screenY + (this.outerHeight - 260)/2;
window.open('dLogin.htm','propop','screenX='+ x1 + ',screenY=' + y1 + ',width=400,height=260,resizable=no,toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no');
}
}
if (navigator.appName == "Microsoft Internet Explorer")					
{
var x1 =  (screen.width  - 385)/2;
var y1 =  (screen.height - 260)/2;

window.open('mLogin.htm','propop','screenX='+ x1 + ',screenY=' + y1 + ',left='+ x1 + ',top='+ y1 + ',width=450,height=250,resizable=no,toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no');
}
}
}
}



//-->

</script>

<script language="JavaScript">
<!--
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);
// -->

function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
//-->
</script>
</head>

<body bgcolor="#ffffff" vlink="#000000" onLoad="MM_preloadImages('images/contact_event.gif')" topmargin="8">
<table border="0" cellspacing="0" cellpadding="0" width="100%">
  <tr> 
    <td valign="top" colspan="2" width="752"> 
      <table border="0" cellpadding="0" cellspacing="0" width="100%">
        <tr> 
          <td height="54" valign="top"><img name="arshlogo" src="images/logobbl3.gif" width="150" height="82" border="0"></td>
          <td valign="top" height="54"> 
            <table border="0" cellpadding="0" cellspacing="0" width="602">
              <tr> 
                <td> 
                  <table border="0" cellpadding="0" cellspacing="0" width="595">
                    <tr> 
                      <td><img name="blank1" src="images/blank1.gif" width="370" height="34" border="0"></td>
                      <td><img name="topbar_r1_c3" src="images/topbar_r1_c3.gif" width="8" height="34" border="0"></td>
                      <td> 
                        <table border="0" cellpadding="0" cellspacing="0" width="34">
                          <tr> 
                            <td><img name="images/topbar_r1_c4.gif" width="34" height="2" border="0" src="images/topbar_r1_c4.gif"></td>
                          </tr>
                          <tr> 
                            <td><a href="default.asp" onMouseOut="MM_swapImgRestore()"  onMouseOver="MM_swapImage('home','','images/home_f2.gif',1)" ><img name="home" src="images/home.gif" width="34" height="29" border="0"></a></td>
                          </tr>
                          <tr> 
                            <td><img name="topbar_r3_c4" src="images/topbar_r3_c4.gif" width="34" height="3" border="0"></td>
                          </tr>
                        </table>
                      </td>
                      <td><img name="topbar_r1_c5" src="images/topbar_r1_c5.gif" width="3" height="34" border="0"></td>
                      <td> 
                        <table border="0" cellpadding="0" cellspacing="0" width="34">
                          <tr> 
                            <td><img name="topbar_r1_c6" src="images/topbar_r1_c6.gif" width="34" height="2" border="0"></td>
                          </tr>
                          <tr> 
                            <td><a href="international.asp" onMouseOut="MM_swapImgRestore()"  onMouseOver="MM_swapImage('intern','','images/intern_f2.gif',1)" ><img name="intern" src="images/intern.gif" width="34" height="29" border="0"></a></td>
                          </tr>
                          <tr> 
                            <td><img name="topbar_r3_c6" src="images/topbar_r3_c6.gif" width="34" height="3" border="0"></td>
                          </tr>
                        </table>
                      </td>
                      <td><img name="topbar_r1_c7" src="images/topbar_r1_c7.gif" width="4" height="34" border="0"></td>
                      <td> 
                        <table border="0" cellpadding="0" cellspacing="0" width="34">
                          <tr> 
                            <td><img name="topbar_r1_c8" src="images/topbar_r1_c8.gif" width="34" height="2" border="0"></td>
                          </tr>
                          <tr> 
                            <td><a href="search.asp" onMouseOut="MM_swapImgRestore()"  onMouseOver="MM_swapImage('search','','images/search_f2.gif',1)" ><img name="search" src="images/search.gif" width="34" height="29" border="0"></a></td>
                          </tr>
                          <tr> 
                            <td><img name="topbar_r3_c8" src="images/topbar_r3_c8.gif" width="34" height="3" border="0"></td>
                          </tr>
                        </table>
                      </td>
                      <td><img name="topbar_r1_c9" src="images/topbar_r1_c9.gif" width="4" height="34" border="0"></td>
                      <td> 
                        <table border="0" cellpadding="0" cellspacing="0" width="34">
                          <tr> 
                            <td><img name="topbar_r1_c10" src="images/topbar_r1_c10.gif" width="34" height="2" border="0"></td>
                          </tr>
                          <tr> 
                            <td><a href="sitemap.asp" onMouseOut="MM_swapImgRestore()"  onMouseOver="MM_swapImage('sitemap','','images/sitemap_f2.gif',1)" ><img name="sitemap" src="images/sitemap.gif" width="34" height="29" border="0"></a></td>
                          </tr>
                          <tr> 
                            <td><img name="topbar_r3_c10" src="images/topbar_r3_c10.gif" width="34" height="3" border="0"></td>
                          </tr>
                        </table>
                      </td>
                      <td><img name="topbar_r1_c11" src="images/topbar_r1_c11.gif" width="4" height="34" border="0"></td>
                      <td> 
                        <table border="0" cellpadding="0" cellspacing="0" width="34">
                          <tr> 
                            <td><img name="topbar_r1_c12" src="images/topbar_r1_c12.gif" width="34" height="2" border="0"></td>
                          </tr>
                          <tr> 
                            <td><a href="faqpage.asp" onMouseOut="MM_swapImgRestore()"  onMouseOver="MM_swapImage('help','','images/help_f2.gif',1)" ><img name="help" src="images/help.gif" width="34" height="29" border="0"></a></td>
                          </tr>
                          <tr> 
                            <td><img name="topbar_r3_c12" src="images/topbar_r3_c12.gif" width="34" height="3" border="0"></td>
                          </tr>
                        </table>
                      </td>
                      <td><img name="topbar_r1_c13" src="images/topbar_r1_c13.gif" width="39" height="34" border="0"></td>
                    </tr>
                  </table>
                </td>
              </tr>
              <tr> 
                <td align="right" height="27" valign="top"> 
                  <table border="0" cellspacing="0" cellpadding="0" width="85%" align="right">
                    <tr> 
                      <td align="right" valign="top" height="9"><img src="images/curve.gif" width="64" height="25"></td>
                      <td align="center" bgcolor="#FFCA95" valign="top" height="9"><a href="about.asp" class="textl">About</a>&nbsp;&nbsp;<span class="text3">|</span> 
                        <a href="services.asp" class="textl">Our Services</a> 
                        &nbsp;<span class="text3">|</span>&nbsp; <a href="pricelist.asp" class="textl">Price 
                        List </a>&nbsp;<span class="text3">|</span>&nbsp; <a href="driverpage.asp" class="textl">Downloads</a>&nbsp;<span class="text3">|</span>&nbsp; 
                        <a onClick = "popLogin()" href="#" class="textl"><strong> 
                        Login</strong></a></td>
                    </tr>
                  </table>
                </td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
    <td valign="top" align="right" colspan="2" class="text" width="752"> 
      <%if Request.Cookies("isDealer") <> "" then%>
      You are Logged in as DEALER Don't forget to <a class="link1" href ="dealerlogout.asp">LOGOUT</a> 
      <% elseif Request.Cookies("isMember") <> "" then%>
      You are Logged in as Member Don't forget to <a class="link1" href ="memberlogout.asp">LOGOUT</a> 
      <%end if%>
      <br>
  </tr>
  <tr> 
    <td valign="top" align="center" colspan="2" width="752"> <br>
      <%  
   Dim objAd
   Set objAd = Server.CreateObject("MSWC.AdRotator") 'Create the component
   '''Response.Write objAd.GetAdvertisement("adrot.txt") 
'   Response.Write objAd.GetAdvertisement("/banners/adrot.txt") 
   Set objAd = Nothing 'Destroy the object
   %>
    </td>
  </tr>
  <tr> 
    <td valign="top" align="center" width="75%">&nbsp; </td>
<%'15-MAR    <td valign="top" align="center" width="25%"><%if Request.Cookies("isDealer") <> "" then%> 
<%'15-MAR      <a href="admin/ORDtracking.asp?dealID=<%=trim(Request.Cookies("dealerid"))%><%'"  class="textl">[Track your Orders]</a> %>
<%'15-MAR      <%end if%>
<%'15-MAR          <%if Request.Cookies("isMember") <> "" then%>
<%'15-MAR          <a href="admin/ORDtracking.asp?memID=<%=trim(Request.Cookies("actid"))%><%'" class="textl"><font color="#0000FF">[Track your Orders]</font></a> %>
 <%'15-MAR      <%end if%>
</td>
  </tr>
		   <table border="0" cellspacing="0" cellpadding="0"  width="100%">
    <tr align="left"> 
      <td valign="top" height="27" align="left" width="300" style="border-bottom-style: solid; border-bottom-width: 1" > 
        <span class=p2li1><a href="default.asp" class=p2li1>Home :</a> <br>
        <font color=blue><%=request("head1")%></font> <font color=red><%=request("head2")%></font></span> 
      </td>
      <%
       if request("isdealer") <> "" or request("ismember") <> "" then
       %>
      <td valign="top" height="27" align="right" width="400" style="border-bottom-style: solid; border-bottom-width: 1" > 
<%'15-MAR        <b><font size="1">Items in CART:<font color="#FF0000"> %>
        <%
       if session("tot") <> "" then
         'response.write "<br>i am here" & session("tot") & "--<br>"
       %>
        <%=session("nitems")%></font>, Value :<font color="#FF0000"><%=trim(session("tot"))%></font> 
        : <a href="cart2.asp?bid=<%=session("cartid")%>&mcode=<%=scode%>">VIEW 
        CART</a> <a href="checkout.asp?bid=<%=session("cartid")%>"> Checkout</a> 
        <%else%>
<%'15-MAR        None</font></b> %>
        <%end if%>
      </td>
      <%
       end if
       %>
    </tr>
    <tr> 
      <td valign=top colspan="2"></td>
		  <table width="100%" border="0" cellspacing="0" cellpadding="0" class="b2b">
        <tr> 
          <td valign="top" nowrap width="20%"><img src="images/spacer.gif" width="145" height="1"></td>
          <td valign="top" nowrap width="1"><img src="images/spacer.gif" width="1" height="1" ></td>
          <td valign="top" nowrap width="20%"><img src="images/spacer.gif" width="145" height="1" ></td>
          <td valign="top" nowrap width="1"><img src="images/spacer.gif" width="1" height="1" ></td>
          <td valign="top" nowrap width="60%"> </td>
        </tr>
        <tr> 
          <td valign="top" nowrap width="20%" height="78">

<table width="100%" border="0" cellspacing="0" cellpadding="0" height="20">
              <%if rsmain.RecordCount >0 then
                    do while (not rsMain.EOF )%>
<tr> 
<td width="100%" height="15"><a href="page2.asp?pgt=tt&Mcode=<%=rsMain("CatID")%>&head1=<%=rsMain("groupdesc")%>&mData="" """"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""" class="p2li2">
<%=rsMain("groupdesc")%></a></td>
</tr>
              <%
          if rsMain.EOF=false then
          rsMain.movenext
		  end if
		  loop
		end if%>
</table>
          </td>
          <td valign="top" nowrap width="1" height="78"  bgcolor="silver" ></td>
          <td valign="top" nowrap width="20%" height="78"> 
            <table border="0" cellspacing="0" cellpadding="0" width="100%" align="left">
              <%if rssub.RecordCount > 0 then
                  do while (not rsSub.EOF )%>
              <tr> 
                <td width="100%" height="1"><a href="page2.asp?mData=<%=rsSub("CatCode")%>&Mcode=<%=rsSub("CatID")%>&head1=<%=request("head1")%>&head2=<%=rsSub("category")%>" class="p2li2"><%=rsSub("category")%></a></td>
				</tr>
              <%
          if rssub.Eof = false then
          rsSub.movenext
          end if
					loop%>
              <%else%>
 <tr> 
                <td width="152" height="15"></td>
              </tr>
              <%end if%>
            </table>
          </td>
          <td valign="top" nowrap width="1" height="78" bgcolor="silver" class="textl" ></td>
          <td valign="top" nowrap width="100%" height="78">
          <% if a="" then  %>
          <%if rsdata.RecordCount >0 then
            if xdata <> "" then
               mloop1=rsdata.RecordCount/3
               mloop=int(rsdata.RecordCount/3)
                if mloop1 > mloop then
				mloop3=mloop+1
				else
				mloop3=mloop1
               end if
           else
           mloop1=rsdata.RecordCount/3
           if mloop1 < 1 then
				mloop3=1
			else	           
             mloop3=3
            end if 
           end if  
      
      for i=1 to mloop3
      
      if rsdata.EOF=false then
      %>       
		 <table width="100%" border="0" cellspacing="0" cellpadding="0" >
            <tr> 
              <td width="33%"><img src="images/spacer.gif" width="123" height="1" ></td>
              <td width="1"><img src="images/spacer.gif" width="1" height="1" ></td>
              <td width="33%"><img src="images/spacer.gif" width="123" height="1" ></td>
              <td width="1"><img src="images/spacer.gif" width="1" height="1" ></td>
            </tr><td width="33%" height="1">
            <table border="0" cellspacing="0" cellpadding="0" height="222" width="100%">
              <tr> 
              <td height="2" colspan="2" width="100%"><img src="images/spacer.gif" width="117" height="1"></td>
              </tr>
              <tr> 
                <%if not isnull(rsdata("image")) then%>
                <td height="2" align="center" colspan="2"> <img src=<%=rsData("image")%>></td>
                <%else%>
                <td height="2" align="center" colspan="2">&nbsp;</td>
                <%end if%>
              </tr>
              <tr> 
                <td height="2" colspan="2" class="text3" align="center"><%=rsData("itemcode")%></td>
              </tr>
              <%
				if Request.Cookies("isDealer") <> "" then
				mprice_= trim(rsdata("pga")) 
				elseif Request.Cookies("isMember") <> "" then
				mprice_= trim(rsdata("retail")) 
				else
				mprice_= ""
				end if
				%>

		      <%if  (rsdata("picturethumb"))  <> "" then%>
              <tr align="center"> 
                
			    <td height="28" colspan="2"> <a href="page2.asp?a=<%="zaf"%>&Mcode=<%=scode%>&pID=<%=rsdata("productID")%>&price=<%=(mprice_)%>"> 
                  <img src=<%=rsData("picturethumb")%>  border=0> </a> </td>
              </tr>
          
			  <%else%>
              <tr align="center"> 
                <td height="28" colspan="2"> <a href="page2.asp?a=<%="zaf"%>&Mcode=<%=scode%>&pID=<%=rsdata("productID")%>&price=<%=mprice_%>"> 
                  <img src="images/noimage.gif" border=0></a></td>
              </tr>
              <%end if%>
              <tr> 
                <td height="38" colspan="2" class=desc align="center"> <%=left(rsData("description"),100)%> 
                </td>
              </tr>
              <tr> 
                <td height="29"> 
                  <%if Request.Cookies("isDealer") <> "" then%>
                  <a href="page2.asp?a=<%="zaf"%>&Mcode=<%=scode%>&pID=<%=rsdata("productID")%>&price=<%=formatnumber(trim(rsdata("pga")),2) %>"> 
                  <img src="images/buynow.gif" border=0></a></td>
                <%
                if rsdata("pga") = "" or trim(price) = 0 then
                %>
                <td height="29" class="text">Call<br>
                  for Price</td>
                <%else%>
                <td height="29" class="text">Price<br>
                  Rs.<%=formatnumber(trim(rsdata("pga")),2)%> </td>
                <%end if
                end if

                if Request.Cookies("isMember") <> "" then 
                %>
                <a href="page2.asp?a=<%="zaf"%>&Mcode=<%=scode%>&pID=<%=rsdata("productID")%>&price=<%=formatnumber(trim(rsdata("retail")),2)%>"> 
                <img src="images/buynow.gif" border=0></a></td>
              <%if rsdata("retail") = "" or trim(rsdata("retail"))=0 or trim(price) = 0 then%>
              <td height="29" class="text">Call<br>
                for Price</td>
              <%else%>
              <td height="29" class="text">Price<br>
                Rs.<%=formatnumber(trim(rsdata("retail")),2)%></td>
              <%end if
                
                
                end if%>
              </tr>
            </table>
            <%
            if rsdata.EOF=false then
            rsdata.MoveNext
            end if
                %></td>
            <td width="1" background="images/dot_grey.gif" height="29">&nbsp;</td><td width="33%" height="29">
            <%if  not rsdata.EOF then %>
            <table border="0" cellspacing="0" cellpadding="0" height="222" width="100%">
              <tr> 
                <td height="2" colspan="2" width="100%" ><img src="images/spacer.gif" width="117"  height="1"></td>
              </tr>
<tr>              
 <%
                 if not isnull(rsdata("image")) then%>
                <td height="2" align="center" colspan="2"><img src=<%=rsData("image")%>></td>
                <%else%>
                <td height="2" align="center" colspan="2">&nbsp;</td>
                <%end if%>
              </tr>
              <tr> 
                <td height="2" colspan="2" class="text3" align="center"><%=rsData("itemcode")%></td>
              </tr>
              <%
				if Request.Cookies("isDealer") <> "" then
				mprice_= trim(rsdata("pga"))
				elseif Request.Cookies("isMember") <> "" then
				mprice_= trim(rsdata("retail"))
				else
				mprice_= ""
				end if
				%>
             <%if  (rsdata("picturethumb"))  <> "" then%>
              <tr align="center"> 
                <td height="28" colspan="2"> <a href="page2.asp?a=<%="zaf"%>&Mcode=<%=scode%>&pID=<%=rsdata("productID")%>&price=<%=mprice_%>"> 
                  <img src=<%=rsData("picturethumb")%> border=0></a> </td>
              </tr>
              <%else%>
              <tr align="center"> 
                <td height="28" colspan="2"> <a href="page2.asp?a=<%="zaf"%>&Mcode=<%=scode%>&pID=<%=rsdata("productID")%>&price=<%=mprice_%>"> 
                  <img src="images/noimage.gif" border=0 width="75" height="86"></a></td>
              </tr>
              <%end if%>
              <tr> 
                <td height="38" colspan="2" class=desc align="center"><%=left(rsData("description"),100)%></td>
              </tr>
              <tr> 
                <td height="29"> 
                  <%
                if Request.Cookies("isDealer") <> "" then
                %>
                  <a href="page2.asp?a=<%="zaf"%>&Mcode=<%=scode%>&pID=<%=rsdata("productID")%>&price=<%=formatnumber(trim(rsdata("pga")),2)%>"> 
                  <img src="images/buynow.gif" border=0></a></td>
                <%if rsdata("pga") = "" or trim(price)=0 then%>
                <td height="29" class="text">Call<br>
                  for Price</td>
                <%else%>
                <td height="29" class="text">Price<br>
                  Rs.<%=formatnumber(trim(rsdata("pga")) ,2)%></td>
                <%end if
                end if
                
                
                if Request.Cookies("isMember") <> "" then
       
                %>
                <a href="page2.asp?a=<%="zaf"%>&Mcode=<%=scode%>&pID=<%=rsdata("productID")%>&price=<%=formatnumber(trim(rsdata("retail")),2)%>"> 
                <img src="images/buynow.gif" border=0></a></td>
              <%if rsdata("retail") = "" or trim(rsdata("retail"))=0 or trim(price) = 0 then%>
              <td height="29" class="text">Call<br>
                for Price</td>
              <%else%>
              <td height="29" class="text">Price<br>
                Rs.<%=formatnumber(trim(rsdata("retail")) ,2)%></td>
              <%end if
                end if
                %>
              </tr>
            </table>
            <%else%>
           
            <%end if%>
            <%
          if rsdata.EOF=false then
          rsdata.MoveNext
          end if
          %></td>
            <td width="1" background="images/dot_grey.gif" height="29">&nbsp;</td><td width="33%" height="29">
            <%if  not rsdata.EOF then %>
            <table border="0" cellspacing="0" cellpadding="0" height="222" width="100%">
              <tr> 
                <td height="2" colspan="2" width="100%"><img src="images/spacer.gif" width="117" height="1"></td>
              </tr>
              <tr> 
                <%if not isnull(rsdata("image"))then %>
                <td height="2" align="center" colspan="2" > <img src=<%=rsData("image")%>> 
                </td>
                <%else%>
                <td height="2" align="center" colspan="2" width="25">&nbsp;</td>
                <%end if%>
              </tr>
              <tr> 
                <td height="2" colspan="2" class="text3" align="center" ><%=rsData("itemcode")%></td>
              </tr>
              <%
				if Request.Cookies("isDealer") <> "" then
				mprice_= trim(rsdata("pga"))
				elseif Request.Cookies("isMember") <> "" then
				mprice_= trim(rsdata("retail"))
				else
				mprice_= ""
				end if
			%>
              <%if  (rsdata("picturethumb"))  <> "" then%>
              <tr align="center"> 
                <td height="28" colspan="2" > <a href="page2.asp?a=<%="zaf"%>&Mcode=<%=scode%>&pID=<%=rsdata("productID")%>&price=<%=mprice_%>"> 
                  <img src=<%=rsData("picturethumb")%>  border=0></a> </td>
              </tr>
              <%else%>
              <tr align="center"> 
                <td height="28" colspan="2" > <a href="page2.asp?a=<%="zaf"%>&Mcode=<%=scode%>&pID=<%=rsdata("productID")%>&price=<%=mprice_%>"> 
                  <img src="images/noimage.gif" border=0></a></td>
              </tr>
              <%end if%>
              <tr> 
               <td height="38" colspan="2" class=desc align="center"> <%=left(rsData("description"),100)%> </td>
              </tr>
              <tr> 
                <td height="29" width="76"> 
                  <%if Request.Cookies("isDealer") <> "" then %>
                  <a href="page2.asp?a=<%="zaf"%>&Mcode=<%=scode%>&pID=<%=rsdata("productID")%>&price=<%=formatnumber(trim(rsdata("pga")),2)%>"> 
                  <img src="images/buynow.gif" border=0></a></td>
                <%if rsdata("pga") = "" or trim(price)=0 then%>
                <td height="29" class="text" width="41">Call<br>
                  for Price</td>
                <%else%>
                <td height="29" class="text" width="25">Price<br>
                  Rs.<%=formatnumber(trim(rsdata("pga")),2)%></td>
                <%end if
              
              end if
              
              if Request.Cookies("isMember") <> "" then
              %>
                <a href="page2.asp?a=<%="zaf"%>&Mcode=<%=scode%>&pID=<%=rsdata("productID")%>&price=<%=formatnumber(trim(rsdata("retail")),2) %>"> 
                <img src="images/buynow.gif" border=0></a></td>
              <%if rsdata("retail") = "" or trim(rsdata("retail"))=0 or trim(price) = 0 then%>
              <td height="29" class="text" width="41">Call<br>
                for Price</td>
              <%else%>
              <td height="29" class="text" >Price<br>
                Rs.<%=formatnumber(trim(rsdata("retail")) ,2)%></td>
              <%end if
              end if
              %>
              </tr>
            </table>
            <%else%>
           
            <%end if%>
            <%
        if rsdata.EOF=false then
        rsdata.MoveNext
        end if
        %></td>
            <tr align="center"> 
              <td colspan="5" width="100%"> 
                <table width="80%" border="0" cellspacing="0" cellpadding="0" height="1">
                  <tr> 
                    <td bgcolor="#6666CC"></td>
                  </tr>
                </table>
              </td>
            </tr>
          </table>
          <%   
 
 end if
 next
 end if
else%>
          <form action="cart2.asp?xpid=<%=pid%>&Mcode=<%=scode%>&mScat=<%=mScat%>&bid=<%=cartID%>" method="post">
            <table width="420" border="1" cellspacing="0" cellpadding="0" height="385" bordercolorlight="#FFD700" bordercolordark="#FFD700" style="border-collapse: collapse" bordercolor="#111111">
              <tr valign="top"> 
                <td colspan="3" height="35" width="400" style="border-left-style: solid; border-left-width: 1; border-right-style: solid; border-right-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: none; border-bottom-width: medium"><span class="text2">Prod ID 
                  :</span><span class="text5">&nbsp;<%=rsdata("itemcode")%></span> <span class="text2">Brand 
                  :</span><span class="text5">&nbsp;<%=brand%></span><span class="text2">&nbsp;Warranty : 
                  </span><span class="text5"><%=Warranty%></span><br>
                  <span class="text-heading"><%=rsdata("description")%>

                  &nbsp;</span> 
                  <%if Request.Cookies("isDealer") <> "" then              
	Stxt="select * ,[dbo].Editprice(pga * "& price__ &") as pga,[dbo].Editprice(pgb * "& price__ &") as pgb,[dbo].Editprice(pgc * "& price__ &") as pgc,[dbo].Editprice(pgd * "& price__ &") as pgd,[dbo].Editprice(pge * "& price__ &") as pge from vPqtyUnit where productid="&pid&"" 
	rsVQty.CursorLocation = adUseClient
	rsVQty.Open Stxt,conn,adOpenForwardOnly,adLockReadOnly,adcmdtext
	set rsVQty.ActiveConnection = nothing
%>
                  <table width="90%" border="1" cellspacing="0" cellpadding="0">
                    <tr> 
                      <td width="20%" CLASS="TEXT">if you buy</td>
                      <td width="16%" align="center" class="text" ><%=rsVqty("qtyA")%>&nbsp;</td>
                      <td width="16%" align="center" class="text" Color=Red><%=rsVqty("qtyB")%>&nbsp;</td>
                      <td width="16%" align="center" class="text" Color=Red><%=rsVqty("qtyC")%>&nbsp;</td>
                      <td width="16%" align="center" class="text" Color=Red><%=rsVqty("qtyD")%>&nbsp;</td>
                      <td width="16%" align="center" class="text" Color=Red><%=rsVqty("qtyE")%>&nbsp;</td>
                    </tr>
                    <tr> 
                      <td width="20%" class="text">you pay/unit</td>
                      <%
      %>
                      <td width="16%" align="right" class="text"><%=formatnumber(trim(rsVqty("pgA")),2)%></td>
                      <td width="16%" align="right" class="text"><%=formatnumber(trim(rsVqty("pgB")),2)%></td>
                      <td width="16%" align="right" class="text"><%=formatnumber(trim(rsVqty("pgc")),2)%></td>
                      <td width="16%" align="right" class="text"><%=formatnumber(trim(rsVqty("pgd")),2)%></td>
                      <td width="16%" align="right" class="text"><%=formatnumber(trim(rsVqty("pge")),2)%></td>
                    </tr>
                    <tr> 
                      <td width="20%" class="text">you save/unit</td>
                      <td width="16%" align="center" class="text">-Nil-</td>
                      <td width="16%" align="right" class="text"><%=formatnumber(cstr(trim(rsVqty("pgA")))-cstr(trim(rsVqty("pgb"))),2)%></td>
                      <td width="16%" align="right" class="text"><%=formatnumber(cstr(trim(rsVqty("pgA")))-cstr(trim(rsVqty("pgc"))),2)%></td>
                      <td width="16%" align="right" class="text"><%=formatnumber(cstr(trim(rsVqty("pgA")))-cstr(trim(rsVqty("pgd"))),2)%></td>
                      <td width="16%" align="right" class="text"><%=formatnumber(cstr(trim(rsVqty("pgA")))-cstr(trim(rsVqty("pge"))),2)%></td>
                    </tr>
                  </table>
                  <br>
                  <%end if%>
                </td>
              </tr>
              <tr> 
                <td align="center" width="133" rowspan="2" style="border-left-style: solid; border-left-width: 1; border-right-style: none; border-right-width: medium; border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium"> 
                  <br>
                    <%if pic <> "" and trim(pic) <> "PIMAGES\" then%>
                  <img src="<%=pic%>"> 
				  <%else%>
				  <img src="images/NOIMAGE1.GIF" border=0>
                  <%end if%>
                 
		
				 
				   <br>
                  <%if mdsheet <> "" and  trim(mdsheet) <> "PIMAGES\" then%>
				  
				  <a class="text" href = "<%=mdsheet%>"><img src="images/PDF.GIF" alt="Data Sheet" border="0"></a><%end if%>&nbsp;&nbsp;&nbsp;
                  <%'if not rsdriver.EOF then%>
                  
				  
                <td align="left" width="100" rowspan="2" valign="top" style="border-style: none; border-width: medium"><span class="text-heading"> 
                  <%if stimag <> "" then%>
                  <img src="<%=stimage%>" border=0> 
                  <%end if%>
                  </span></td>
                <%
              if Request.Cookies("ismember") <> "" then
              %>
                <td valign="bottom" width="150" height="185" style="border-left-style: none; border-left-width: medium; border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium"> 
                  <br>
                  <span class="text2">Price :</span> <span class="text3">Rs.<%=formatnumber(price,2)%></span><br>
                  <br>
<%'15-MAR           <span class="text2">Qty : </span> %>
<%'15-MAR                  <input type="text" name="mqty" size="8" maxlength="8"  value=1  class="input1"  > %>
                  <br>
                  <br>
                </td>
                <%
              end if
              %>
                <%
              if Request.Cookies("isdealer") <> "" then
              %>
                <td valign="bottom" width="150" height="185" style="border-left-style: none; border-left-width: medium; border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium"> 
                  <br>
                  <span class="text2">Price :</span> <span class="text3">Rs:<%=formatnumber(price,2)%></span><br>
                  <br>
<%'15-MAR                  <span class="text2">Qty : </span>  %>
<%'15-MAR                  <input type="text" name="mqty" size="8" maxlength="8" value=1 class="input1"> %>
                  <br>
                  <br>
                </td>
                <%
              end if
              %> 
              </tr>
              <tr> 
                <%
            if request("isdealer") <> "" or request("ismember") <> "" then
            %>
                <td width="160" style="border-left-style: none; border-left-width: medium; border-right-style: solid; border-right-width: 1; border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium"> 
<%' 15-mar  <input id=submit1 name="addincart" type=image  src="../images/addcart.gif"  value="Add to Cart" > %></td>
                <%end if%>
              </tr>
               <%
Function GetHTMLFile(strFileName)
Dim fso, file, FName
FName = Server.MapPath(ldesc)
Set fso = Server.CreateObject ("Scripting.FileSystemObject")
'if Not(fso Is Nothing) Then
'If fso.FileExists(FName) Then
Set file=fso.OpenTextFile(FName,1)
GetHTMLFile = Trim(file.readAll())
'response.write "sdfsdkj"
'Else
'GetHTMLFile =   "dsds" & FName  & vbCrLf 
'End If
'Set fso = Nothing
'Else
'GetHTMLFile= "bnfjfj" & err.description 
'End If
End Function%>
			  
			  <tr valign="top">  
                <td colspan="3"  style="border-bottom-style: solid; border-bottom-width: 1">
               <%if ldesc <> "" and  trim(ldesc) <> "PIMAGES\" then%>
                 <span class="text4"><%Response.Write GetHTMLFile(request("ldesc"))%></span></td> 
               <%end if%>
			  </tr>
            </table>
          </form>
          <% end if %></td>
        </tr>
      </table>
    </tr>
  </table>
</table>

      <table border="0" cellspacing="0" cellpadding="0" width="100%">
	    <tr> 
	     <td  align="center" height="23" width="100%"><br>
            <p align="center" class="font"><b>Baylan Technology Inc. Copyright &copy; 
              2007</b></p>
          </td>
        </tr>
        <tr> 
          <td class="font" align="center" height="20" width="100%"> <b> </b> <a href="faqpage.asp" class="text">Customer 
            Support</a><b> &nbsp;&nbsp;</b>|<b>&nbsp; &nbsp;</b><a href="privacy.asp" class="text">Privacy 
            Policy</a><b> &nbsp;&nbsp;</b>|<b>&nbsp;&nbsp; </b><a href="feedback.asp" class="text">Feed 
            Back</a><b> &nbsp;&nbsp;</b>|<b>&nbsp;&nbsp; </b><a href="terms.asp" class="text">Terms 
            And Conditions</a></td>
        </tr>
        <tr> 
          
    <td class="font" align="center" width="100%">&nbsp;</td>
        </tr>
        <tr> 
          
    <td class="font" align="center" width="100%">&nbsp;</td>
        </tr>
      </table>
<br>
<div id="contact" style="position:absolute; width:34; height:34; z-index:2; left: 722px; top: 11px"><a href="contactus.asp"><img src="images/contact_base.gif" width="34" height="29" border="0" name="Image1" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image1','','images/contact_event.gif',1)"></a></div>
<%
'response.Write(Request.Cookies("isdealer")	)
%>
</table></tr></a></td>
 </body>
</html>