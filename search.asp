<%@ Language=VBScript %>
<%
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
set conn = server.CreateObject ("ADODB.Connection")
set rsSubCat = server.CreateObject("ADODB.Recordset")
set rsMCat = server.CreateObject("ADODB.Recordset")
%>
<!--#INCLUDE FILE="dbconnect.asp"-->
<%
conn.Open con 
SQLtxt="select * from Vsubcat " 
rsSubCat.CursorLocation = adUseClient
rsSubCat.Open SQLtxt,conn,adOpenForwardOnly,adLockReadOnly,adcmdtext
set rsSubCat.ActiveConnection = nothing
SQLtxt1="select * from VMaincat " 
rsMCat.CursorLocation = adUseClient
rsMCat.Open SQLtxt1,conn,adOpenForwardOnly,adLockReadOnly,adcmdtext
set rsMCat.ActiveConnection = nothing




%>
<head>
<title>Search Results:</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">

<script language=JavaScript>
window.name = "arshimain";
var imgurl="images/";
</script>
<SCRIPT language=JavaScript src="scripts/image_init.js" type=text/javascript></SCRIPT>

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

</script>
<link rel="stylesheet" href="css/standard.css" type="text/css">
</head>

<body bgcolor="#ffffff" vlink="#000000" 
onLoad="MM_preloadImages('images/contact_event.gif')" topmargin="8">
<table border="0" cellspacing="0" cellpadding="0" width="515">
  <tr> 
    <td valign="top" colspan="2" width="515"> 
      <table border="0" cellpadding="0" cellspacing="0" width="748">
        <tr> 
          
          <td height="54" valign="top" width="150"><img name="arshlogo" src="images/logobbl3.gif" width="150" height="82" border="0"></td>
          <td valign="top" height="54" width="598"> 
            <table border="0" cellpadding="0" cellspacing="0" width="577">
              <tr> 
                <td> 
                  <table border="0" cellpadding="0" cellspacing="0" width="600">
                    <tr> 
                      <td width="384"><img name="blank1" src="images/blank1.gif" width="375" height="34" border="0"></td>
                      <td width="10"><img name="topbar_r1_c3" src="images/topbar_r1_c3.gif" width="1" height="34" border="0"></td>
                      <td width="34"> 
                        <table border="0" cellpadding="0" cellspacing="0" width="34">
                          <tr> 
                            <td><img name="images/topbar_r1_c4.gif" width="34" height="2" border="0" src="images/topbar_r1_c4.gif"></td>
                          </tr>
                          <tr> 
                            
                          <td><a href="default.asp" onMouseOut="MM_swapImgRestore()"  onMouseOver="MM_swapImage('home','','images/home_f2.gif',1)" ><img name="home" src="images/home.gif" width="34" height="29" border="0" alt="About"></a></td>
                          </tr>
                          <tr> 
                            <td><img name="topbar_r3_c4" src="images/topbar_r3_c4.gif" width="34" height="3" border="0"></td>
                          </tr>
                        </table>
                      </td>
                      <td width="3"><img name="topbar_r1_c5" src="images/topbar_r1_c5.gif" width="3" height="34" border="0"></td>
                      <td width="34"> 
                        <table border="0" cellpadding="0" cellspacing="0" width="34">
                          <tr> 
                            <td><img name="topbar_r1_c6" src="images/topbar_r1_c6.gif" width="34" height="2" border="0"></td>
                          </tr>
                          <tr> 
                            
                          <td><a href="international.asp" onMouseOut="MM_swapImgRestore()"  onMouseOver="MM_swapImage('intern','','images/intern_f2.gif',1)" ><img name="intern" src="images/intern.gif" width="34" height="29" border="0" alt="International"></a></td>
                          </tr>
                          <tr> 
                            <td><img name="topbar_r3_c6" src="images/topbar_r3_c6.gif" width="34" height="3" border="0"></td>
                          </tr>
                        </table>
                      </td>
                      <td width="4"><img name="topbar_r1_c7" src="images/topbar_r1_c7.gif" width="4" height="34" border="0"></td>
                      <td width="34"> 
                        <table border="0" cellpadding="0" cellspacing="0" width="34">
                          <tr> 
                            <td><img name="topbar_r1_c8" src="images/topbar_r1_c8.gif" width="34" height="2" border="0"></td>
                          </tr>
                          <tr> 
                            
                            <td><a href="search.asp" onMouseOut="MM_swapImgRestore()"  onMouseOver="MM_swapImage('search','','images/search_f2.gif',1)" ><img name="search" src="images/search.gif" width="34" height="29" border="0" alt="Search"></a></td>
                          </tr>
                          <tr> 
                            <td><img name="topbar_r3_c8" src="images/topbar_r3_c8.gif" width="34" height="3" border="0"></td>
                          </tr>
                        </table>
                      </td>
                      <td width="4"><img name="topbar_r1_c9" src="images/topbar_r1_c9.gif" width="4" height="34" border="0"></td>
                      <td width="34"> 
                        <table border="0" cellpadding="0" cellspacing="0" width="34">
                          <tr> 
                            <td><img name="topbar_r1_c10" src="images/topbar_r1_c10.gif" width="34" height="2" border="0"></td>
                          </tr>
                          <tr> 
                            
                          <td><a href="sitemap.asp" onMouseOut="MM_swapImgRestore()"  onMouseOver="MM_swapImage('sitemap','','images/sitemap_f2.gif',1)" ><img name="sitemap" src="images/sitemap.gif" width="34" height="29" border="0" alt="Site Map"></a></td>
                          </tr>
                          <tr> 
                            <td><img name="topbar_r3_c10" src="images/topbar_r3_c10.gif" width="34" height="3" border="0"></td>
                          </tr>
                        </table>
                      </td>
                      <td width="4"><img name="topbar_r1_c11" src="images/topbar_r1_c11.gif" width="4" height="34" border="0"></td>
                      <td width="34"> 
                        <table border="0" cellpadding="0" cellspacing="0" width="34">
                          <tr> 
                            <td><img name="topbar_r1_c12" src="images/topbar_r1_c12.gif" width="34" height="2" border="0"></td>
                          </tr>
                          <tr> 
                            
                            <td><a href="faqpage.asp" onMouseOut="MM_swapImgRestore()"  onMouseOver="MM_swapImage('help','','images/help_f2.gif',1)" ><img name="help" src="images/help.gif" width="34" height="29" border="0" alt="Help"></a></td>
                          </tr>
                          <tr> 
                            <td><img name="topbar_r3_c12" src="images/topbar_r3_c12.gif" width="34" height="3" border="0"></td>
                          </tr>
                        </table>
                      </td>
                      <td width="36"><img name="topbar_r1_c13" src="images/topbar_r1_c13.gif" width="45" height="34" border="0"></td>
                    </tr>
                  </table>
                </td>
              </tr>
              <tr> 
                <td align="right" height="27" valign="top"> 
                  <table border="0" cellspacing="0" cellpadding="0" width="85%">
                    <tr> 
                      <td align="right" valign="top" height="24"><img src="images/curve.gif" width="64" height="24"></td>
                       <td align="center" bgcolor="#FFCA95" valign="top" height="24"><a href="about.asp" class="textl">About</a>&nbsp;&nbsp;<span class="text3">|</span> 
                      <a href="services.asp" class="textl">Our Services</a> &nbsp;<span class="text3">|</span>&nbsp; 
                      <a href="pricelist.asp" class="textl">Price List </a>&nbsp;<span class="text3">|</span>&nbsp; 
                      <a href="driverpage.asp" class="textl">Downloads</a>&nbsp;<span class="text3">|</span>&nbsp; 
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
      </table>
  
<table border="0" cellspacing="0" cellpadding="0" width="750">
    <td valign="top" align="right" colspan="2" class="text" height="25" width="750"> 
      <%if Request.Cookies("isDealer") <> "" then%>
      You are Logged in as DEALER Don't forget to<a class="link1" href ="dealerlogout.asp">&nbsp;<font color="#FF0000">LOGOUT</font></a> 
      <%end if%>
      <% if Request.Cookies("isMember") <> "" then%>
      You are Logged in as Member Don't forget to <a class="link1" href ="memberlogout.asp"><font color="#FF0000">LOGOUT</font></a> 
      <%end if%>
    </td>
  <tr> 
    <td align=center width="750"> 
      <%  
   Dim objAd
   Set objAd = Server.CreateObject("MSWC.AdRotator") 'Create the component
  '' Response.Write objAd.GetAdvertisement("adrot.txt") 
    Response.Write objAd.GetAdvertisement("/banners/adrot.txt") 
   Set objAd = Nothing 'Destroy the object
   %>
    </td>
  </tr>
 
    <td colspan="2"  valign="top" > 
      <table border="0" cellspacing="0" cellpadding="0" width="750">
        <tr> 
         
          <td valign="bottom" align="right"> 
            <%if Request.Cookies("isDealer") <> "" then%>
      <a href="admin/ORDtracking.asp?dealID=<%=trim(Request.Cookies("dealerid"))%>" class="textl"><font color="#0000FF">[Track 
      your Orders]</font></a> 
      <%end if%>
      <%if Request.Cookies("isMember") <> "" then
          %>
      <a href="admin/ORDtracking.asp?memID=<%=trim(Request.Cookies("actid"))%>" class="textl"><font color="#0000FF">[Track 
      your Orders]</font></a> 
      <%end if%></td>
        </tr>
      </table>
</table>
<head>
<title>Search:</title>
<link rel="stylesheet" href="admin/arshilogin.css">
</head>

<body bgcolor="#FFFFFF">
<br>
<form action="process_search.asp" method=post id=form1 name=form1>
  <table class=TableBorder cellspacing=2 cellpadding=4 width=685 border=0>
    <tbody> 
    <tr class=TableHeading> 
      <td class=TableHeading valign=top colspan=3>Search Page </td>
    </tr>
    <tr class=TableText> 
      <td class=TableText align=right width="184"> <span class="text"> Look for</span> 
        <input type="checkbox" name="p_usetext" value="yes">
      </td>
      <td class=TableText width="254"> 
        <input size=25 name=p_text>
      </td>
      <td class=text align=center rowspan="4"><font color="#0000FF"><b>Select 
        Your desired options to Search Products.</b></font><br>
        <br>
        <br>
        &quot;Don't forget to Select <br>
        at least one Option&quot;</td>
    
    <tr class=TableText> 
      <td class=TableText align=right width="184"> <span class="text">in Group 
        Category</span> 
        <input type="checkbox" name="p_usemcat" value="yes">
        <span class="text"> </span></td>
      <td class=TableText width="254"> 
        <select name="p_MCat">
          <%Do While Not rsMCat.EOF%> 
          <option value="<%=rsMCat("catid")%>"> <%=rsMCat("GroupDesc")%> 
       </option>
          <%rsMCat.MoveNext 
            Loop%> 
        </select>
      </td>
    <tr class=TableText> 
      <td class=text align=right width="184"> in Category 
        <input type="checkbox" name="p_usecat" value="yes">
      </td>
      <td class=TableText width="254"> 
        <select name="p_Cat">
          <%Do While Not rsSubCat.EOF%> 
          <option value="<%=rsSubCat("CatCode")%>"> <%=rsSubCat("Category")%> 
          </option>
          <%rsSubCat.MoveNext 
            Loop%> 
        </select>
      </td>
    <tr class=TableText> 
      <td class=TableText align=left width="184">&nbsp; </td>
      <td class=TableText width="254">&nbsp; </td>
    <tr class=TableText> 
      <td class=TableText align=right colspan=3>
<INPUT id=IMAGE 
            onclick=submit(); type=image 
            alt="Click to Continue" src="images/search2.gif" value=Login 
            border=0 name=IMAGE>
      </td>
    </tr>
    </tbody> 
  </table>
  <br>
</form>
	
<br>
  <hr color=#000000 size=1 width="760" noShade align="left">
  <div align="center"></div>
  <tr valign="middle"> 
    <table border="0" cellspacing="0" cellpadding="0" width="760">
      <tr> 
        
      <td  align="center" height="14">&nbsp; </td>
      </tr>
      <tr> 
        <td  align="center" height="23"> 
          
        <p align="center" class="font"><b>Baylan Technology Inc. Copyright &copy; 2007</b></p>
        </td>
      </tr>
      <tr> 
        <td class="font" align="center" height="20"> <b> </b> <a href="faqpage.asp" class="text">Customer 
          Support</a><b> &nbsp;&nbsp;</b>|<b>&nbsp; &nbsp;</b><a href="privacy.asp" class="text">Privacy 
          Policy</a><b> &nbsp;&nbsp;</b>|<b>&nbsp;&nbsp; </b><a href="feedback.asp" class="text">Feed 
          Back</a><b> &nbsp;&nbsp;</b>|<b>&nbsp;&nbsp; </b><a href="terms.asp" class="text">Terms 
          And Conditions</a></td>
      </tr>
      <tr> 
        
      <td class="font" align="center">&nbsp;</td>
      </tr>
      <tr> 
        
      <td class="font" align="center">&nbsp;</td>
      </tr>
    </table>
  </tr>
  <img src="images/spacer.gif" width="1" height="1"><br>
  
<div id="contact" style="position:absolute; width:34; height:34; z-index:2; left: 725px; top: 10px"><a href="contactus.asp"><img src="images/contact_base.gif" width="34" height="29" border="0" name="Image1" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image1','','images/contact_event.gif',1)" alt="Contact Us"></a></div>
</body>
</html>