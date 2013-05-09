<%@ Language=VBScript %>
<%
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
set conn = server.CreateObject ("ADODB.Connection")
set rsOffer = server.CreateObject("ADODB.Recordset")
set rscat = server.CreateObject("ADODB.Recordset")
set rssubcat = server.CreateObject("ADODB.Recordset")
set rsdumm = server.CreateObject("ADODB.Recordset")
set rsdumm2 = server.CreateObject("ADODB.Recordset")
''if request("rd") = "" then
''Response.Redirect("default.asp?rd=yes")
''end if
 if Request.Cookies("isDealer") <> "" then
 xmsg="You are Loggedin as Dealer :"
 end if

 if Request.Cookies("isMember") <> "" then
 xmsg="You are Logged in as : " & Request.Cookies("MemName")
 end if

%>
<!--#INCLUDE FILE="dbconnect.asp"-->
<%
conn.Open con 

%>
<html> <head> 
<meta http-equiv="Content-Language" content="en-us">
<title>Baylan Technology Inc.</title> 
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<script language=JavaScript>
window.name = "arshimain";
var imgurl="images/";
</script>
<SCRIPT language=JavaScript src="scripts/image_init.js" type=text/javascript></SCRIPT>
<SCRIPT language=JavaScript>
<!--
function bring_forth_date()
{
    months = new Array ("January", "February", "March", "April", "May",
                        "June", "July", "August", "September", "October",
			"November", "December");
    days = new Array ("Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday",
                        "Saturday");
    mydate = new Date();
    hour      = mydate.getHours();
    minute    = mydate.getMinutes();
    day_month = mydate.getDate();
    month     = mydate.getMonth();
    year      = mydate.getYear();    
    day       = mydate.getDay(); 
    if (minute < 10)  { minute = "0" + minute; }
    ampm = "am";
    if (hour == 0)  { hour  = 12; }
    if (hour > 12)  { hour -= 12;  ampm = "pm"; }
    if (year >= 90 && year < 200) {
        year += 1900;
    }
	date_str  = days [day] + ", ";
    date_str += months [month] + " ";
    date_str += day_month + ", ";
    date_str += year + ", ";
    date_str += hour + ":";
    date_str += minute + "";
    document.write(date_str);

}
//-->
</SCRIPT>
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
var start = coki.indexOf("isDealer=");
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
        window.open('dLogin.htm','propop','screenX='+ x1 + ',screenY=' + y1 + ',width=350,height=240,resizable=no,toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no');
} 
else
{
var x1 = this.screenX + (this.outerWidth  - 385)/2;
var y1 = this.screenY + (this.outerHeight - 260)/2;
						window.open('dLogin.htm','propop','screenX='+ x1 + ',screenY=' + y1 + ',width=385,height=260,resizable=no,toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no');
}
}
if (navigator.appName == "Microsoft Internet Explorer")					
{
var x1 =  (screen.width  - 385)/2;
var y1 =  (screen.height - 260)/2;
window.open('dLogin.htm','propop','screenX='+ x1 + ',screenY=' + y1 + ',left='+ x1 + ',top='+ y1 + ',width=400,height=250,resizable=no,toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no');
}
}
}
//-->
</script>
<script language="JavaScript">
function popLogin2() 
{
var coki = document.cookie;
var start = coki.indexOf("isDealer=");

if (start >= 0)
{
var start = coki.indexOf("=",start) + 1;
var end = coki.indexOf(";",start);

var value = unescape(coki.substring(start,end));
alert("You are Already Logged-in as Dealer");
}
else
{
if (navigator.appName == "Netscape")
{
if (navigator.appVersion.charAt(0) < 4) 
 {
        window.open('mLogin.htm','propop','screenX='+ x1 + ',screenY=' + y1 + ',width=350,height=240,resizable=no,toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no');
} 
else
{
var x1 = this.screenX + (this.outerWidth  - 385)/2;
var y1 = this.screenY + (this.outerHeight - 260)/2;
						window.open('dLogin.htm','propop','screenX='+ x1 + ',screenY=' + y1 + ',width=385,height=260,resizable=no,toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no');
}
}
if (navigator.appName == "Microsoft Internet Explorer")					
{
var x1 =  (screen.width  - 385)/2;
var y1 =  (screen.height - 260)/2;
window.open('mLogin.htm','propop','screenX='+ x1 + ',screenY=' + y1 + ',left='+ x1 + ',top='+ y1 + ',width=400,height=250,resizable=no,toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no');
}
}
}
//-->
function CheckBlan(){
var correctEmail=0;
var CNameOK = 1;
var emaillength = 0;
var x1 =  (screen.width-400)/2;
var y1 =  (screen.height-300 )/2;

if (CNameOK==1) {
emaillength=document.homeP.Email.value.length;
	if (emaillength==0)
	alert ('Please provide the email address for subscription.')
	document.homeP.Email.focus()

if (emaillength != 0 ){
    for ( var i=0; i<emaillength;i++){
    if (document.homeP.Email.value.charAt(i) == "@")
	correctEmail=1;
    }
       if (correctEmail==0)
	   alert("Incorrect Email Address.")
	   else
	   document.cookie = "emailT=" + escape(document.homeP.Email.value);
	   document.cookie = "NameT=" + escape(document.homeP.NameT.value);
	   document.cookie = "subt=" + escape(subt);
	  // alert (subt)
	   childWin= open('p_visitors.asp' ,'childWin','screenX='+ x1 + ',screenY=' + y1 + ',left='+ x1 + ',top='+ y1 + ',width=420,height=100,resizable=yes,toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no');  
	   //alert (document.homeP.emailstatus.value )
	   //document.homeP.submit()
	   }
}
}                
var subt = 'sc'
function checkboxo(pickit){
subt=pickit.value
//alert (subt)
}


function userCat()
{
var x1 =  (screen.width-400)/2;
var y1 =  (screen.height-300 )/2;
childWin= open('verify.asp' ,'childWin','screenX='+ x1 + ',screenY=' + y1 + ',left='+ x1 + ',top='+ y1 + ',width=420,height=100,resizable=no,toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no');
}


</script>




<link rel="stylesheet" href="css/standard.css" type="text/css">
</head>

<body bgcolor="#ffffff" vlink="#000000" 
onLoad="MM_preloadImages('images/contact_event.gif','/admin/images/topbar_r2_c7_f2.gif','images/topbar_r2_c11_f2.gif','images/topbar_r2_c13_f2.gif','images/topbar_r2_c17_f2.gif','images/topbar_r2_c19_f2.gif','images/topbar_r4_c2_f2.gif','images/topbar_r4_c4_f2.gif','admin/images/topbar_r4_c6_f2.gif','images/topbar_r4_c9_f2.gif','images/topbar_r4_c15_f2.gif','images/topbar_r4_c21_f2.gif','images/home_f2.gif','images/intern_f2.gif','images/search_f2.gif','images/sitemap_f2.gif','images/help_f2.gif')" topmargin="8"><form method="post" action="p_visitors.asp" name="homeP"> 
<table border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td valign="top" colspan="2"> 
      <table border="0" cellpadding="0" cellspacing="0" width="699">
        <tr> 
          <td height="54" valign="top" width="150"><img name="arshlogo" src="images/logobbl3.gif" width="150" height="82" border="0"></td>
          <td valign="top" height="54" width="635"> 
            <table border="0" cellpadding="0" cellspacing="0" width="603">
              <tr> 
                <td> 
                  <table border="0" cellpadding="0" cellspacing="0" width="561">
                    <tr> 
                      <td width="404"><img name="blank1" src="images/blank1.gif" width="388" height="34" border="0"></td>
                      <td width="8"><img name="topbar_r1_c3" src="images/topbar_r1_c3.gif" width="1" height="34" border="0"></td>
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
                      <td width="8"><img name="topbar_r1_c5" src="images/topbar_r1_c5.gif" width="3" height="34" border="0"></td>
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
                      <td width="8"><img name="topbar_r1_c7" src="images/topbar_r1_c7.gif" width="4" height="34" border="0"></td>
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
                      <td width="8"><img name="topbar_r1_c9" src="images/topbar_r1_c9.gif" width="4" height="34" border="0"></td>
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
                      <td width="8"><img name="topbar_r1_c11" src="images/topbar_r1_c11.gif" width="4" height="34" border="0"></td>
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
                      <td width="35"><img name="topbar_r1_c13" src="images/topbar_r1_c13.gif" width="36" height="34" border="0"></td>
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
    <td valign="top" align="right" colspan="2" class="normtext" height="25"> 
      <%if Request.Cookies("isDealer") <> "" then%>
      <font color=blue> <%=xmsg%> <b> <font color="#FF0000">Don't forget to</font></b> 
      <a class="textl" href ="dealerlogout.asp">| LOGOUT |</a></font> 
      <%end if%>
      <%if Request.Cookies("isMember") <> "" then%>
      <font color=blue> <%=xmsg%> <b><font color="#FF0000">Don't forget to</font></b> 
      <a class="textl" href ="memberlogout.asp">| LOGOUT |</a></font> 
      <%end if%>
      <script language=JavaScript>
<!--  
bring_forth_date ();
//-->
      </script>
    </td>
  </tr>
  <tr> 
    <td valign="top" align="center"> 
      <%  
   Dim objAd
   Set objAd = Server.CreateObject("MSWC.AdRotator") 'Create the component
   ''Response.Write objAd.GetAdvertisement("adrot.txt") 
   Response.Write objAd.GetAdvertisement("/banners/adrot.txt") 
   Set objAd = Nothing 'Destroy the object
   %>
    </td>
    <td valign="top" align="center"> </td>
  </tr>
  <tr> 
    <td colspan="2"  valign="top" > 
      <table border="0" cellspacing="0" cellpadding="0" width="100%">
        <tr> 
          <td colspan="2" align="center" >&nbsp;&nbsp;</td>
          <td valign="bottom" align="center"> 
            <%if Request.Cookies("isDealer") <> "" then%>
            <a href="admin/ORDtracking.asp?dealID=<%=session("dealerid")%>" class="textl">[Track 
            your Orders]</a> 
            <%end if%>
            <%if Request.Cookies("isMember") <> "" then%>
            <a href="admin/ORDtracking.asp?memID=<%=session("memberid")%>" class="textl"><font color="#0000FF">[Track 
            your Orders]</font></a> 
            <%end if%>
          </td>
        </tr>
		</table>
		<table cellspacing=0 cellpadding=0 width=687 border=0 align="center">
        <tr> 
          <td valign="bottom" align="center">
<p align="left"> 
            <table width="687" border="0" cellspacing="0" cellpadding="1" align="left">
              <tr>
                <td width="129" style="height: 20px"><font face="Arial, Helvetica, sans-serif" size="2"><b>Company 
                  Outline: </b></font></td>
                <td width="554" style="height: 20px"></td>
              </tr>
              <tr> 
                <td width="129"><font face="Arial, Helvetica, sans-serif" size="2">Address</font></td>
                <td width="554"><font face="Arial, Helvetica, sans-serif" size="2">708-Japan 
                  Plaza, M.A.Jinnah Road, Karachi-74400 (PAKISTAN)</font></td>
              </tr>
              <tr> 
                <td width="129"><font face="Arial, Helvetica, sans-serif" size="2">Phone</font></td>
                <td width="554"><font face="Arial, Helvetica, sans-serif" size="2">+92-21-32764881 
                  / 32764900&nbsp; / 32720969</font></td>
              </tr>
              <tr> 
                <td width="129" style="height: 18px"><font face="Arial, Helvetica, sans-serif" size="2">Fax</font></td>
                <td width="554" style="height: 18px"><font face="Arial, Helvetica, sans-serif" size="2">+92-21-32737819</font></td>
              </tr>
              <tr> 
                <td width="129"><font face="Arial, Helvetica, sans-serif" size="2">Email</font></td>
                <td width="554"><font face="Arial, Helvetica, sans-serif" size="2"><a href="mailto:info@bbl.com.pk">info@bbl.com.pk</a></font></td>
              </tr>
              <tr> 
                <td width="129"><font face="Arial, Helvetica, sans-serif" size="2">Sales</font></td>
                <td width="554"><font face="Arial, Helvetica, sans-serif" size="2"><a href="mailto:hiladha@bbl.com.pk">hiladha@bbl.com.pk</a></font></td>
              </tr>
              <tr> 
                <td width="129"><font face="Arial, Helvetica, sans-serif" size="2">Web</font></td>
                <td width="554"><font face="Arial, Helvetica, sans-serif" size="2"><a href="http://www.bbl.com.pk">www.bbl.com.pk</a></font></td>
              </tr>
            </table>
            <p align="left"><font face="Arial, Helvetica, sans-serif" size="2"><b><br>
              </b></font>
          </td>
        </tr>
      </table>


  </tr>
  <tr> 
    <td colspan="3" height="26"> 
      <div align="center"> 
        <table width="90%" border="0" cellspacing="0" cellpadding="0" height="1" background="images/dot_grey1.gif">
          <tr> 
            <td height="4"></td>
          </tr>
        </table>
      </div>
      <hr color=#000000 size=1 width="650" align="center" noShade>
    </td>
  </tr>
</table>
    
<tr valign="middle"></tr>
<div align="center"></div>
<tr valign="middle">
  <table border="0" cellspacing="0" cellpadding="0" width="93%">
    <tr> 
      <td  align="center" height="23"> 
        <p align="center" class="font"><b>Baylan Technology Inc. Copyright © 2009</b></p>
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
<div id="contact" style="position:absolute; width:34; height:34; z-index:2; left: 736px; top: 11px"><a href="contactus.asp"><img src="images/contact_base.gif" width="34" height="29" border="0" name="Image1" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image1','','images/contact_event.gif',1)" alt="Contact Us"></a></div>
<SCRIPT language=JavaScript src="scripts/buzz.js" type=text/javascript></SCRIPT>
<SCRIPT language=JavaScript src="scripts/dynlayer.js" type=text/javascript></SCRIPT>
<SCRIPT language=JavaScript src="scripts/buzz_scroll.js" type=text/javascript></SCRIPT>
<SCRIPT language=JavaScript type=text/javascript>
</form>
</body>
</html>

