<%@ Language=VBScript %>
<%
on error resume next
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
set conn = server.CreateObject ("ADODB.Connection")
set rsFaq = server.CreateObject("ADODB.Recordset")
set rsAns= server.CreateObject("ADODB.Recordset")
if Request.Cookies("isDealer") <> "" then
 xmsg="You are Logged in as " & Request.Cookies("typename") & " : [" & Request.Cookies("dealer_id") & "]"
 end if
%>
<!--#INCLUDE FILE="dbconnect.asp"-->
<%
conn.Open con 

SQLtxt="select * from faqtype" 
rsFaq.CursorLocation = adUseClient
rsFaq.Open SQLtxt,conn,adOpenForwardOnly,adLockReadOnly,adcmdtext
set rsFaq.ActiveConnection = nothing
ftype=rsfaq("FAqType")

%>
<html><head><title>Arshi Computer Network Products Faq/Help</title> 
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
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
 days = new Array ("Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday","Saturday");
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
var oklogin=0;
if (coki.indexOf("isDealer=") != -1 ){
var start = coki.indexOf("isDealer=");
alert ('Already Login ')
oklogin=1
}
if (coki.indexOf("isMember=") != -1 ){
var start = coki.indexOf("isDealer=");
alert ('Already Login ')
oklogin=1
}
if (oklogin==0){
if (start >= 0)
{
var start = coki.indexOf("=",start) + 1;
var end = coki.indexOf(";",start);
var value = unescape(coki.substring(start,end));
alert("You are Already Logged-in ");
}
else
{
if (navigator.appName == "Netscape")
{
if (navigator.appVersion.charAt(0) < 4) 
{
  window.open('Login.htm','propop','screenX='+ x1 + ',screenY=' + y1 + ',width=400,height=240,resizable=no,toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no');
} 
else
{
var x1 = this.screenX + (this.outerWidth  - 385)/2;
var y1 = this.screenY + (this.outerHeight - 260)/2;
window.open('Login.htm','propop','screenX='+ x1 + ',screenY=' + y1 + ',width=400,height=260,resizable=no,toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no');
}
}
if (navigator.appName == "Microsoft Internet Explorer")					
{
var x1 =  (screen.width  - 385)/2;
var y1 =  (screen.height - 260)/2;
window.open('Login.htm','propop','screenX='+ x1 + ',screenY=' + y1 + ',left='+ x1 + ',top='+ y1 + ',width=450,height=250,resizable=no,toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no');
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
alert("You are Already Logged-in ");
}
else
{
if (navigator.appName == "Netscape")
{
if (navigator.appVersion.charAt(0) < 4) 
{
 window.open('amLogin.htm','propop','screenX='+ x1 + ',screenY=' + y1 + ',width=400,height=240,resizable=no,toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no');
} 
else
{
var x1 = this.screenX + (this.outerWidth  - 385)/2;
var y1 = this.screenY + (this.outerHeight - 260)/2;
window.open('Login.htm','propop','screenX='+ x1 + ',screenY=' + y1 + ',width=400,height=260,resizable=no,toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no');
}
}
if (navigator.appName == "Microsoft Internet Explorer")					
{
var x1 =  (screen.width  - 385)/2;
var y1 =  (screen.height - 260)/2;
window.open('amLogin.htm','propop','screenX='+ x1 + ',screenY=' + y1 + ',left='+ x1 + ',top='+ y1 + ',width=450,height=250,resizable=no,toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no');
}
}
}
}
//-->


function subscribe() 
{
if (navigator.appName == "Netscape")
{
if (navigator.appVersion.charAt(0) < 4) 
{
 window.open('subscribe.asp','propop','screenX='+ x1 + ',screenY=' + y1 + ',width=400,height=240,resizable=no,toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no');
} 
else
{
var x1 = this.screenX + (this.outerWidth  - 385)/2;
var y1 = this.screenY + (this.outerHeight - 260)/2;
window.open('subscribe.asp','propop','screenX='+ x1 + ',screenY=' + y1 + ',width=400,height=260,resizable=no,toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no');
}
}
if (navigator.appName == "Microsoft Internet Explorer")					
{
var x1 =  (screen.width  - 385)/2;
var y1 =  (screen.height - 260)/2;
window.open('subscribe.asp','propop','screenX='+ x1 + ',screenY=' + y1 + ',left='+ x1 + ',top='+ y1 + ',width=450,height=250,resizable=no,toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no');
}
}


//-->

</script>
<link rel="stylesheet" href="acss/standard.css" type="text/css">
<SCRIPT language=JavaScript>
var doImage=doImage;var TType=TType;
function mhHover(tbl,idx,cls){var t,d;if(document.getElementById)t=document.getElementById(tbl);else t=document.all(tbl);if(t==null)return;if(t.getElementsByTagName)d=t.getElementsByTagName("TD");else d=t.all.tags("TD");if(d==null)return;if(d.length<=idx)return;d[idx].className=cls;}
function footerjs(doc){if(doImage==null){var tt=TType==null?"PV":TType;doc.write('<layer visibility="hide"><div style="display:none"></div></layer>');}}
</SCRIPT>
</head>
<BODY bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<table cellspacing=0 cellpadding=0 width="100%" border=0>
 <tbody>
 <tr>
 <td width="44%">
  <table height=22 cellspacing=0 cellpadding=0 width="100%" border=0>
    <tbody>
     <tr>
       <td id=arshiid bgcolor=#ffffff></td>
      <td id=msviRegionGradient1 
    style="FILTER: progid:DXImageTransform.Microsoft.Gradient(startColorStr='#FFFFFF', endColorStr='#AEC0EC', gradientType='1')" 
    width="50%"></td>
      <td id=msviRegionGradient2 
    style="FILTER: progid:DXImageTransform.Microsoft.Gradient(startColorStr='#AEC0EC', endColorStr='#6487DB', gradientType='1')" 
    width="50%"></td>
    </tr>
    </tbody>
    </table></td>
    <td height=22 colspan="6" align=left nowrap bgcolor=#6487db id=actoolbar dir=ltr>
      <table align=right cellspacing=0 cellpadding=0 border=0>
      <tbody>
      <tr>
      <td class=ac0 onMouseOver="mhHover('actoolbar', 0, 'ac1')"
	  onMouseOut="mhHover('actoolbar', 0, 'ac0')" nowrap><a href="default.asp">Home</a></td>
	  <td class=acsep>|</td>
	  <td class=ac1 onMouseOver="mhHover('actoolbar', 2, 'ac1')"
	  onMouseOut="mhHover('actoolbar', 2, 'ac1')" nowrap><a href="network.asp">Networking Products</a></td>
	  <td class=acsep>|</td>
	  <td class=ac0 onMouseOver="mhHover('actoolbar', 4, 'ac1')"
	  onMouseOut="mhHover('actoolbar', 4, 'ac0')" nowrap><a href="computer.asp">Computer Products </a></td>
      <td class=acsep>|</td>
      <td class=ac0 onMouseOver="mhHover('actoolbar', 6, 'ac1')" 
      onMouseOut="mhHover('actoolbar', 6, 'ac0')" nowrap><a href="searchpage.asp" >Search</a></td>
      <td class=acsep>|</td>
      <td class=ac0 onMouseOver="mhHover('actoolbar', 8, 'ac1')" 
       onMouseOut="mhHover('actoolbar', 8, 'ac0')" nowrap><a href="#" >Faq / Help</a></td>
      <td class=acsep>|</td>
      <td class=ac0 onMouseOver="mhHover('actoolbar', 10, 'ac1')" 
       onMouseOut="mhHover('actoolbar', 10, 'ac0')" nowrap><a href="contact.asp" >Contact</a></td>
     </tr>
   </tbody>
   </table></td>
    </tr>
    <tr valign=top>
    <td width="44%" rowspan="6">    <table height=42 cellspacing=0 cellpadding=0 width="100%" border=0>
      <tbody>
        <tr valign=top>
          <td bgcolor=#ffffff><a> <img title="" height=101 alt=Windows src="aimages/logo.gif" width=160 border=0></a></td>
          <td style="FILTER: progid:DXImageTransform.Microsoft.Gradient(startColorStr='#FFFFFF', endColorStr='#B6C5EE', gradientType='1')" 
   width="100%"></td>
        </tr>
      </tbody>
    </table></td>
   <td  height=22 colspan="6" align=right vAlign=top bgcolor=#b6c5ee>
   <table class=menucolornormal  cellspacing=0 cellpadding=0 border=0>
    <tbody>
    <tr>
    <td class=ab0 nowrap> <a href="aboutus.asp" class="style2"><SPAN class=menufontnormal
    onmouseover="this.className='menufonthover'" 
    onmouseout="this.className='menufontnormal'" style2>About Us</SPAN></a></td>
    <td class=absep>|</td>
	<td class=ab0 nowrap><a href="price_list.asp" class="style2"><SPAN class=menufontnormal 
    onmouseover="this.className='menufonthover'" 
    onmouseout="this.className='menufontnormal'">Price List</span></a></td>
    <td class=absep>|</td>
    <td class=ab0  nowrap><a href="download.asp" class="style2"><SPAN class=menufontnormal 
    onmouseover="this.className='menufonthover'" 
    onmouseout="this.className='menufontnormal'">Downloads</span></a></td>
    <td class=absep>|</td>
    <td class=ab0  nowrap><a onClick = "popLogin()" href="#" class="style2"><SPAN class=menufontnormal 
    onmouseover="this.className='menufonthover'" 
    onmouseout="this.className='menufontnormal'">Login</span></a></td>
    <td class=absep>|</td>
    <td class=ab0  nowrap><a href=javascript:window.open('adealers/signup.asp','arshimain'); class="style2"><SPAN class=menufontnormal 
    onmouseover="this.className='menufonthover'" 
    onmouseout="this.className='menufontnormal'">Signup</span></a></td>
	 	<td class=absep>|</td>
	<td class=ab0 nowrap><a onClick = "subscribe()" href="#" class="style2"><SPAN class=menufontnormal 
    onmouseover="this.className='menufonthover'" 
    onmouseout="this.className='menufontnormal'">Subscribe</span></a></td>
    </tr>
    </tbody>
    </table></td>
    </tr>
    <tr valign=top>
    <td height=11 colspan="6" align=right nowrap bgcolor=#b6c5ee class="style3" id=actoolbar dir=ltr>
     <script language=JavaScript>
<!--  
bring_forth_date ();
//-->
      </script>	  </td>
    </tr>
  <td height=22 colspan="6" align=right nowrap bgcolor=#b6c5ee >
      <table align=right cellspacing=0 cellpadding=0 border=0>
      <tbody>
      <tr Valign="TOP">
		<td  align="right" bgcolor=#b6c5ee  class=headbar><div align="right">
	   <%
       if request("isdealer") <> "" or request("ismember") <> "" then
       %>
       <b>Items in CART:<font color="#FF0000"> 
       <%
       if session("tot") <> "" then%>
       <%=session("nitems")%></font>, Value : <font color="#FF0000"><%=trim(session("tot"))%></font> 
       : </div></td>
		<td class=absep><div align="right">|</div></td> 
		<td  align="right" bgcolor=#b6c5ee  class=headbar>  <div align="right">
		<a href="acart2.asp?bid=<%=session("cartid")%>&mcode=<%=scode%>"class="headbar"><SPAN class=menufontnormal 
    onmouseover="this.className='menufonthover'" 
    onmouseout="this.className='menufontnormal'">VIEW CART</A></div></td>
		<td class=absep><div align="right">|</div></td> 
        <td  align="right" bgcolor=#b6c5ee  class=headbar>  <div align="right">
		<a href="acheckout.asp?bid=<%=session("cartid")%>"class="headbar"> <SPAN class=menufontnormal 
    onmouseover="this.className='menufonthover'" 
    onmouseout="this.className='menufontnormal'"> Checkout</a></DIV></td>
	
	<td> 
         <%else%>
          None</font> 
         <%end if%>        
        </div></td>
	 <%
       end if
       %>
</tr>
	</table>
	
	
	
	
	</tr>
     <tr valign="top">
     <td  height="11"  align="right" colspan="6" bgcolor=#b6c5ee class=headbar><div align="right">
	  <%if Request.Cookies("isDealer") <> "" then%>
	  <a href="admin/ORDtracking.asp?dealID=<%=trim(Request.Cookies("dealerid"))%>" class="headbar"><SPAN class=menufontnormal 
    onmouseover="this.className='menufonthover'"
	onmouseout="this.className='menufontnormal'">[Track your Orders]</a></font>
	  <%end if%>
      <%if Request.Cookies("isMember") <> "" then %>
      <a href="admin/ORDtracking.asp?memID=<%=trim(Request.Cookies("actid"))%>" class="headbar"><SPAN class=menufontnormal
	onmouseover="this.className='menufonthover'"
    onmouseout="this.className='menufontnormal'">[Track your Orders]</a></font>
      <%end if%></div></td></tr>
	<tr valign=top>
    <td height="11" colspan="6" align="right" bgcolor=#b6c5ee  class=headbar><div align="right">
          <%if Request.Cookies("isDealer") <> "" then%>
           <%=xmsg%> <b> <font color="#FF0000">Don't forget to</font></b> <a class="headbar" href ="userlogout.asp"><SPAN class=menufontnormal 
    onmouseover="this.className='menufonthover'" 
    onmouseout="this.className='menufontnormal'">| LOGOUT |</a></font>
          <%end if%>
          <%if Request.Cookies("isMember") <> "" then%>
          <%=xmsg%> <b><font color="#FF0000">Don't forget to</font></b> <a class="headbar" href ="amemberlogout.asp"><SPAN class=menufontnormal 
    onmouseover="this.className='menufonthover'" 
    onmouseout="this.className='menufontnormal'">| LOGOUT |</a></font>
          <%end if%>
        </div></td>
    </tr>
    
   
</table>
<td valign="top"  colspan="2"  class="normtext" height="25" width="100%"> 
<td valign="top" align="center" width="45"></td>
</tr>
  <tr> 
  <td colspan="2"  valign="top" width="100%" ><tr><td width="25%" valign="bottom" align="center">&nbsp;</td>
  <img src="aimages/spacer.gif" width="1" height="1"></tr>
 <table width="100%" border="0">
  <tr>
  <td><div align="center"> 
  <%  
   Dim objAd
   Set objAd = Server.CreateObject("MSWC.AdRotator") 'Create the component
   Response.Write  objAd.GetAdvertisement ("/abanners/adrot.txt") 
   Set objAd = Nothing 'Destroy the object
   %>
</div>
 </td></tr></table>
  <tr></tr>
<table cellspacing=0 cellpadding=1 width=100% border=0>
    <tbody> 
    <tr> 
      
    <td bgcolor=#b5c4f9 height=19>&nbsp;<a 
            name=top></a><span class="text-heading"><font color="#000000">Help/FAQ</font></span></td>
    </tr>
    <tr> 
      <td bgcolor=> 
            <table cellspacing=0 cellpadding=5 width="100%" bgcolor=#ffffff border=0>
              <tbody> 
              <tr> 
                <td class=largefont colspan="2"><span class="normtext">Below is 
                  a list of Frequently Asked Questions. If you do not find an 
                  answer to your question here, please contact</span> <a 
                  href="mailto:support@arshi.com.pk" class="textl">support@arshi.com.pk</a> 
                  <span class="normtext">Thank you.</span></td>
              </tr>
              <tr> 
                <td class=header align=middle colspan="2"> 
                  <table cellspacing=0 cellpadding=0 width="100%" border=0>
                    <tbody> 
                    <%do while not rsfaq.EOF%>
                    <tr valign=top> 
                      <%if rsfaq("faqtype")="" then%>
                      <td>&nbsp;</td>
                      <%else%>
                      <td class=largefont width="50%"> 
                        <li><a class="text-heading" href="#<%=rsfaq("faqtype")%>"><%=rsfaq("faqtype")%></a> 
                      </td>
                      <%end if
                  rsfaq.MoveNext
                  %>
                      <td width="10%">&nbsp;</td>
                      <%if rsfaq("faqtype")="" then%>
                      <td>&nbsp;</td>
                      <%else%>
                      <td width="50%"> 
                        <li><a class="text-heading" href="#<%=rsfaq("faqtype")%>"><%=rsfaq("faqtype")%></a> 
                      </td>
                      <%end if%>
                    </tr>
                    <%rsfaq.MoveNext
                
                loop%>
                    </tbody> 
                  </table>
                </td>
              </tr>
              <%rsfaq.MoveFirst
          sno=1
          do while not rsFaq.EOF%>
              <tr> 
                <td class=text bgcolor=#cccccc width="91%"> <font color="red"><a name="<%=rsfaq("faqtype")%>"><%=rsfaq("faqtype")%></a></font></td>
                
          <td  bgcolor=#cccccc width="9%" align="center"> <a class="text-heading" href="#top">Top</a></td>
              </tr>
              <%SQLtxt2="select * from faq where faqtype='"&ftype&"'" 
			set rsAns=conn.Execute(sqltxt2)
		
			if not rsAns.EOF then	
			do while not rsAns.EOF
          %>
              <tr> 
                <td class=text colspan="2"><font color="#3333FF"><b>
                <%=sno%>. 
                <%=rsAns("question")%>?</b></font></td>
              </tr>
              <tr>
                <td class=normtext colspan="2"><%=rsAns("answer")%></td>
              </tr>
              <%sno=sno+1
				rsAns.MoveNext
				loop

			else%>
              <tr> 
                
          <td class=text colspan="2"><b><font color="#000000">No Help Available</font></b></td>
              </tr>
              <%end if
			
		  
          rsfaq.MoveNext
          if rsfaq.EOF=false then		
          ftype=rsfaq("FAqType")
          end if
          
          sno=1
          loop
          %>
              <tr>
                <td class="text" colspan="2">&nbsp;</td>
              </tr>
              </tbody> 
            </table>
      </td>
    </tr>
    </tbody> 
</table>

	
<br>
  <table border="0" cellspacing="0" cellpadding="0" width="100%">
    <hr color=#000000 size=1 width="650" align="center" noShade>
    <tr> 
    <td width="100%" height="21" id=msviRegionGradient2 
    style="FILTER: progid:DXImageTransform.Microsoft.Gradient(startColorStr='#FFFFFF', endColorStr='#6487DB', gradientType='1')">	<p align="center" class="font">ARSHI COMPUTER Copyright &copy; 2006</p></td>
	
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