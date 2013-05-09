<%@ Language=VBScript %>
<%
 if Request.Cookies("isDealer") <> "" then
 xmsg="You are Loggedin as Dealer :"
 end if

 if Request.Cookies("isMember") <> "" then
 xmsg="You are Logged in as : " & Request.Cookies("MemName")
 end if
%>
<html> <head> <title>Baylan Technology Inc.</title> 
<meta http-equiv="Content-Type" content="text/html;">
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



<link rel="stylesheet" href="/css/standard.css" type="text/css">
<style type="text/css">
.style2 {
	text-align: left;
}
</style>
</head>

<body bgcolor="#ffffff" vlink="#000000" 
onLoad="MM_preloadImages('images/contact_event.gif','images/home_f2.gif','images/intern_f2.gif','images/search_f2.gif','images/sitemap_f2.gif','images/help_f2.gif')" topmargin="8">
<table border="0" cellspacing="0" cellpadding="0" width="758">
  <tr> 
    <td valign="top" colspan="2" width="758"> 
      <table border="0" cellpadding="0" cellspacing="0" width="699">
        <tr> 
          <td height="54" valign="top" width="150"><img name="arshlogo" src="images/logobbl3.gif" width="150" height="82" border="0"></td>
          <td valign="top" height="54" width="637"> 
            <table border="0" cellpadding="0" cellspacing="0" width="607">
              <tr> 
                <td> 
                  <table border="0" cellpadding="0" cellspacing="0" width="594">
                    <tr> 
                      <td width="392"> 
                        <div align="left"><img name="blank1" src="images/blank1.gif" width="384" height="34" border="0"></div>
                      </td>
                      <td width="1"><img name="topbar_r1_c3" src="images/topbar_r1_c3.gif" width="1" height="34" border="0"></td>
                      <td width="34"> 
                        <table border="0" cellpadding="0" cellspacing="0" width="34">
                          <tr> 
                            <td><img name="images/topbar_r1_c4.gif" width="34" height="2" border="0" src="images/topbar_r1_c4.gif"></td>
                          </tr>
                          <tr> 
                            <td><a href="default.asp" onMouseOut="MM_swapImgRestore()"  onMouseOver="MM_swapImage('home','','images/home_f2.gif',1)" ><img name="home" src="images/home.gif" width="34" height="29" border="0" alt="Home"></a></td>
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
                      <td width="55"><img name="topbar_r1_c13" src="images/topbar_r1_c13.gif" width="38" height="34" border="0"></td>
                    </tr>
                  </table>
                </td>
              </tr>
              <tr> 
                <td align="right" height="27" valign="top"> 
                  <table border="0" cellspacing="0" cellpadding="0" width="86%">
                    <tr> 
                      <td align="right" valign="top" height="24" width="12%"><img src="images/curve.gif" width="64" height="24"></td>
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
  <tr> 
    <td valign="top" align="right" colspan="2" class="text" height="25" width="758"> 
      <%if Request.Cookies("isDealer") <> "" then%>
      You are Logged in as DEALER Don't forget to<font color="#FF0000"> <a class="link1" href ="dealerlogout.asp">LOGOUT</a></font> 
      <% elseif Request.Cookies("isMember") <> "" then%>
      You are Logged in as Member Don't forget to <a class="link1" href ="memberlogout.asp"><font color="#FF0000">LOGOUT</font></a> 
      <%end if%>
    </td>
  </tr>
  <tr> 
    <td valign="top" align="center" width="758"> 
      <div align="center">
	  <%  
   Dim objAd
   Set objAd = Server.CreateObject("MSWC.AdRotator") 'Create the component
  '' ''Response.Write objAd.GetAdvertisement("adrot.txt") 
   Response.Write objAd.GetAdvertisement("/banners/adrot.txt") 
   Set objAd = Nothing 'Destroy the object
   %>
    </div></td>
    
  </tr>
  <tr> 
    <td colspan="2"  valign="top" align="left" height="554" width="758" > 
      <table border="0" cellspacing="0" cellpadding="0" width="755">
        <tr> 
         
          <td width="25%" valign="bottom" align="right"> 
            <%if Request.Cookies("isDealer") <> "" then%>
            <a href="admin/ORDtracking.asp?dealID=<%=trim(Request.Cookies("dealerid"))%>" class="text">[Track 
            your Orders]</a> 
            <%end if%>
            <%if Request.Cookies("isMember") <> "" then
          %>
            <a href="admin/ORDtracking.asp?memID=<%=trim(Request.Cookies("actid"))%>" class="text">[Track 
            your Orders]</a> 
            <%end if%>
          </td>
        </tr>
      </table>

<p class="style2"><strong>BBL</strong> is a&nbsp; computer- based enterprise&nbsp; with resolute 
commitment&nbsp; to provide best&nbsp; solution to computer&nbsp; &amp;&nbsp; 
Network&nbsp; Users. Our dedicated efforts have enabled us to achieve OEM 
Products, Sole Distributor &amp; Authorized Reseller status in Pakistan various 
brands.</p>
<p class="style2"><strong>BBL</strong>&nbsp; Offers&nbsp; customers&nbsp; a&nbsp; broad&nbsp; range of products such as 
Wireless&nbsp; /&nbsp; Wan&nbsp; /&nbsp; Lan Products; Lan Cards&nbsp; /&nbsp; Adapters, PCMCIA, SOHO&nbsp; /&nbsp; Smart&nbsp; / 
Intelligent&nbsp; /&nbsp; Web Management&nbsp; /&nbsp; SWITCHs &amp; HUBs,&nbsp; Module,&nbsp; ISDN&nbsp; /&nbsp;&nbsp; xDSL&nbsp; / &nbsp; Broadband&nbsp; Router,&nbsp; Bridge,&nbsp; Remote&nbsp; Node&nbsp; /&nbsp; Internet 
Sharing Server, Parallel&nbsp; /&nbsp; USB Print Servers, Media Converters, &amp;&nbsp; Single&nbsp; /&nbsp; Multi 
Mode&nbsp; Fiber-Optic&nbsp; Devices,&nbsp; Changers&nbsp; &amp;&nbsp; Accessories;&nbsp; 
Twisted&nbsp; Pairs / Fiber-Optic&nbsp; Cables,&nbsp; Pigtail&nbsp; /&nbsp; Patch&nbsp; /&nbsp; Drop Cords,&nbsp; Face Plates 
Outlets,&nbsp;&nbsp; Patch Panels&nbsp; /&nbsp; Accessories,&nbsp; Connectors,&nbsp; Plugs,&nbsp; 
Testers &amp; Tools; Tele-communication Accessories; Communication 
Hub / Switches / Server/Racks /Cabinets, Cable Management / Accessories;&nbsp; 
Fiber-Optic&nbsp; Patch Panels&nbsp;/&nbsp; LIU&nbsp; /&nbsp; ODF / Splice 
Closure / Pigtail / Drop Cords, Couplers, Connectors, Changers; Multi PC 
Controller / KVM Switches, KVM Drawers, Monitor Splitters, Extenders, Cables, Y 
Cables &amp; Changers; Biometric Finger&nbsp; print&nbsp; /&nbsp; Barcode&nbsp; /&nbsp; Magnetic&nbsp; /&nbsp; RFID&nbsp; Time&nbsp; Management&nbsp; /&nbsp; Access&nbsp; Control&nbsp; / Attendance,&nbsp; Guard Scanners&nbsp; &amp;&nbsp; Data Collection Devices&nbsp; &amp;&nbsp;&nbsp; Accessories;&nbsp; Point&nbsp;&nbsp; of&nbsp;&nbsp; Sales&nbsp; Devices,&nbsp;&nbsp; Scanners,&nbsp;&nbsp; Price&nbsp; Checker,&nbsp; Cash&nbsp; 
Drawers,&nbsp; Customer&nbsp;&nbsp; Display,&nbsp;&nbsp;&nbsp; Barcode Label&nbsp;&nbsp; /&nbsp;&nbsp; Receipt&nbsp;&nbsp; Printers,&nbsp;&nbsp; Programmable&nbsp;&nbsp; Keyboards&nbsp; &amp;&nbsp; Accessories;&nbsp;&nbsp;&nbsp; Intelligent&nbsp; Long&nbsp;&nbsp; Backup&nbsp; UPS + AVR ; Stabilizers;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; You name it, we have it !</p>
<p>Innovative products to support your customers and capture New opportunities. 
Investing in the Right infrastructure Today Will Pay Off in the Long Run.</p>
<p><strong>BBL </strong>provides superior technology, quality, reliability, power and profit. 
Continuously expanding &amp; modernizing the products range &amp; specifications We 
expect to shore our bright future with you company.</p>
		<p align="justify"><br>
      </p>
      <p align="justify"><br>
      </p></td>
  </tr>
  <tr> 
    <td colspan="3" height="26" width="758"> 
      <div align="justify"></div>
      <hr color=#000000 size=1 width="650" align="center" noShade>
    </td>
  </tr>
</table>
    
<tr valign="middle"></tr>
<div align="center"></div>
<tr valign="middle">
  <table border="0" cellspacing="0" cellpadding="0" width="755">
    <tr> 
      <td  align="center" height="23"> 
        <p align="center" class="font"><b>Baylan Technology Inc. Copyright &copy; 2013</b></p>
      </td>
    </tr>
    <tr> 
      <td class="font" align="center" height="20"> <b> </b> <a href="faqpage.asp" class="text">Customer 
        Support</a><b> &nbsp;&nbsp;</b>|<b>&nbsp; &nbsp;</b><a href="privacy.asp" class="text">Privacy 
        Policy</a><b> &nbsp;&nbsp;</b>|<b>&nbsp;&nbsp; </b><a href="feedback.asp" class="text">Feed 
        Back</a><b> &nbsp;&nbsp;</b>|<b>&nbsp;&nbsp; </b><a href="terms.asp" class="text">Terms 
        And Conditions</a></td>
    </tr>
  </table>
    
</tr>

<img src="images/spacer.gif" width="1" height="1"><br>
<div id="contact" style="position:absolute; width:34; height:34; z-index:2; left: 733px; top: 11px"><a href="contactus.asp"><img src="images/contact_base.gif" width="34" height="29" border="0" name="Image1" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image1','','images/contact_event.gif',1)" alt="Contact Us"></a></div>

</body>
</html>