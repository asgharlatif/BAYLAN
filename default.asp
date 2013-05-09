<%@ Language=VBScript %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"><head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<meta name="google-site-verification" content="qHg6Snd0EEbPAr-4PCIwGR4MJNMi5j8RnAwum5sGkVI" />
<head>
<title>Baylan Home Page</title>
<link rel="stylesheet" href="acss/standard.css" type="text/css">
<link rel="stylesheet" type="text/css" href="acss/newstyle.css" />
    <link href="assets/css/bootstrap.css" rel="stylesheet" type="text/css" />
        <script src="JQuery/jquery-1.8.3.js" type="text/javascript"></script>
    <script src="JQuery/jquery-1.8.3.min.js" type="text/javascript"></script>

<script src="/scripts/PopBox.js" type="text/javascript"></script>
<script type="text/javascript">
    popBoxWaitImage.src = "/images/spinner40.gif";
    popBoxRevertImage = "/images/magminus.gif";
    popBoxPopImage = "/images/magplus.gif";
</script> 
	
<script language="JavaScript" type="text/javascript">



    function openMainPagePrintPopup(parameter1) {

       

        //if (windowObjectReference == null || windowObjectReference.closed) {
        ///windowObjectReference = window.open("printpicturecatalogue.asp?" + parameter1, "Product Print Page", "resizable=yes,scrollbars=yes,status=yes");
        window.open("printpicturecatalogue.asp?" + parameter1, "_blank", "resizable=yes,scrollbars=yes,status=yes")
        //}
//        else {
//            windowObjectReference.focus();
        //                };
       // alert(parameter1);
        
    }

    function onImgErrorSmall(source) {
        source.src = "PIMAGES/C/NIMAGE.GIF";
        // disable onerror to prevent endless loop
        source.onerror = "";
        return true;
    }

    function onImgErrorLarge(source) {
        source.src = "PIMAGES/C/NIMAGE.GIF";
        // disable onerror to prevent endless loop
        source.onerror = "";
        return true;
    }
</script>


</head>

<body>


 <!-- Begin Wrapper -->
   <div id="wrapper">
       <!-- Begin Header -->
       <div id="header_new">
           <!--#INCLUDE FILE="header.asp"-->
       </div>
       <!-- End Header -->
       <%
'ON ERROR RESUME NEXT
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
 if Request.Cookies("isDealer") <> "" then
 xmsg="You are Logged in as " & Request.Cookies("typename") & " : [" & Request.Cookies("dealer_id") & "]"
 end if
       %>
       <!--#INCLUDE FILE="dbconnect.asp"-->
       <%
conn.Open con 
SQLtxt_="select * from DCategory where category='"& Request.Cookies("dcategory") &"'" 
set RsPricing=server.CreateObject("adodb.recordset")
RsPricing.CursorLocation = adUseClient
RsPricing.Open SQLtxt_,conn,adOpenForwardOnly,adLockReadOnly,adcmdtext
set RsPricing.ActiveConnection = nothing

if RsPricing.EOF=false then
price=RsPricing("mration")
price__=RsPricing("mration")
session("price__")=RsPricing("mration")
'Response.Write "price1=" & price
end if

SQLcat="select * from vmaincat order by sort_it"
rsCat.CursorLocation = adUseClient
rsCat.Open SQLcat,conn,adOpenForwardOnly,adLockReadOnly,adcmdtext
set rsCat.ActiveConnection = nothing

SQLscat="select * from vsubCat order by category" 
rssubCat.CursorLocation = adUseClient
rssubCat.Open SQLscat,conn,adOpenForwardOnly,adLockReadOnly,adcmdtext
set rssubCat.ActiveConnection = nothing
rssubcat.MoveFirst
msID=rssubcat("catcode")
       %>
       <table width="978" border="0">
           <tr>
               <td width="972">
                   <div align="center">
                       <%  
   Dim objAd
   'Set objAd = Server.CreateObject("MSWC.AdRotator") 'Create the component
   'Response.Write  objAd.GetAdvertisement ("/abanners/adrot.txt") 
   Set objAd = Nothing 'Destroy the object
                       %>
                   </div>
               </td>
           </tr>
       </table>
       <%
set RSoffer = server.CreateObject("ADODB.Recordset")
sqltxt="select * from valldata where statuscode != '08' order by itemcode " 
RSoffer.CursorLocation = adUseClient
RSoffer.Open Sqltxt,conn,adOpenForwardOnly,adLockReadOnly,adcmdtext
set RSoffer.ActiveConnection = nothing

''here was the Quick Finder
       %>

     
        
       <div id="footer">
              <div id="leftcolumn">
               <%
set rsMain = server.CreateObject("ADODB.Recordset")
set rsSub = server.CreateObject("ADODB.Recordset")
 if Request.Cookies("isDealer") <> "" then
 xmsg="You are Logged in as " & Request.Cookies("typename") & " : [" & Request.Cookies("dealer_id") & "]"
 end if
               %>
               <%
SQLtxt="select * from VmainCAT where catid in (select catm  from item where brshds='BayLan' group by catm) order by sort_it" 
rsMain.CursorLocation = adUseClient
'rsMain.Open SQLtxt,conn,adOpenForwardOnly,adLockReadOnly,adcmdtext
'set rsMain.ActiveConnection = nothing
set rsmain=conn.Execute(SQLtxt)
'SQLtxt2="select * from VSUBCat where CatID='"&Scode&"' order by category" 
'rsSub.CursorLocation = adUseClient
'rsSub.Open SQLtxt2,conn,adOpenForwardOnly,adLockReadOnly,adcmdtext
'set rsSub.ActiveConnection = nothing
'	if not rssub.eof then
'		msid2=rssub("catcode")
'	end if
               %>
               <table width="100%" border="0" cellspacing="0" cellpadding="0" class="b2b">
                   <link rel="stylesheet" type="text/css" href="sdmenu/sdmenu.css" />
                   <script type="text/javascript" src="sdmenu/sdmenu.js">
                   </script>
                   <script type="text/javascript">
	// <![CDATA[
                       var myMenu;
                       window.onload = function () {
                           myMenu = new SDMenu("my_menu");
                           myMenu.oneSmOnly = true
                           myMenu.init();
                           myMenu.collapseAll();
                       };
	// ]]>
                   </script>
                   <tr>
                       <td valign="top" nowrap width="100%">
                           <%'if rsmain.RecordCount >0 then%>
                           <div style="float: left" id="my_menu" class="sdmenu">
                               <%  do while (not rsMain.EOF )%>
                               <div>
                                   <span >
                                       <%=rsmain("GroupDesc")%></span>
                                   
                                   <%SQLtxt2="select * from VSUBCat where catcode in (select cat1  from item where brshds='BayLan'  group by cat1) and CatID='"&rsmain("catid")&"' order by category"  
			set rssub=conn.Execute(SQLtxt2)%>
                                   <%do while not rsSub.EOF %>
                                   <a href="apage2.asp?mData=<%=rsSub("CatCode")%>&mcode=<%=rsSub("CatID")%>&head1=<%=rsmain("groupdesc")%>&head2=<%=rsSub("category")%>" />
                                   <%=rsSub("category")%></a>
                                   <%rsSub.movenext
        	loop%>
                               </div>
                               <%
 		 rsmain.MoveNext 
		 loop		 
                               %>
                           </div>
                       </td>
                   </tr>
               </table>
           </div>

           <div id="rightcolumn">
               <img alt="Baylan" src="images/mainimage.jpg" usemap="#planetmap" />

               <map name="planetmap" id="planetmap" >
               
               

                   <area shape="rect" coords="173,57,281,115"  href="#" onclick="openMainPagePrintPopup('catrange=006,019&textn=Point Of Sales Product and Ribbon')" alt="Point Of Sales Product and Ribbon" style="background-color:Red" />
                   <area shape="rect" coords="307,64,423,112"  href="#" onclick="openMainPagePrintPopup('catrange=025&textn=EAS Anti Shoplifting')"  alt="EAS Anti-Shoplifting.." style="background-color:Red" />
                   <area shape="rect" coords="445,49,552,125"  href="#" onclick="openMainPagePrintPopup('catrange=007&textn=Multi Sharing Devices/KVM')" alt="Multi Sharing Devices/KVM" style="background-color:Red" />
                   <area shape="rect" coords="508,174,622,240" href="#" onclick="openMainPagePrintPopup('catrange=002,001&textn=Networking Products and Accessories')" alt="Networking Products and Accessories" style="background-color:Red" />
                   <area shape="rect" coords="580,286,688,363" href="#" alt="POS Touch System & Terminal"   onclick="openMainPagePrintPopup('catrange=014&textn=POS Touch System & Terminal')" style="background-color:Red" />
                   <area shape="rect" coords="502,409,631,467" href="#" alt="Networking Tools & Testers"    onclick="openMainPagePrintPopup('catrange=020&textn=Networking Tools & Testers')" style="background-color:Red" />
                   <area shape="rect" coords="438,520,556,585" href="#" alt="Fiber Products & Accessories"  onclick="openMainPagePrintPopup('catrange=021&textn=Fiber Products & Accessories')" style="background-color:Red" />
                   <area shape="rect" coords="300,520,420,588" href="#" alt="Telecommunication Accessories" onclick="openMainPagePrintPopup('catrange=016&textn=Telecommunication Accessories')" style="background-color:Red" />
                   <area shape="rect" coords="170,521,280,589" href="#" alt="UPS & EVR"  onclick="openMainPagePrintPopup('catrange=005&textn=UPS & EVR')" style="background-color:Red" />
                   <area shape="rect" coords="99,406,222,465" href="#" onclick="openMainPagePrintPopup('catrange=018&textn=Cards / Tags / Disc')" alt="Cards / Tags / Disc" style="background-color:Red" />
                   <area shape="rect" coords="27,299,154,341"  href="#" onclick="openMainPagePrintPopup('catrange=004&textn=Time Attendance & Access Control')" alt="Time Attendance & Access Control" style="background-color:Red" />
                   <area shape="rect" coords="100,167,220,240" href="#" onclick="openMainPagePrintPopup('catrange=013&textn=FingerPrint Time Attendance & AC')" alt="FingerPrint Time Attendance & AC" style="background-color:Red" />
             </map>
         </div>

            <!--#INCLUDE FILE="footer.asp"-->
       </div>
       
		 
   </div>

   
   <!-- End Wrapper -->
    <!-- Begin Footer -->
          
  
   <script type="text/javascript">
var gaJsHost = (("https:" == document.location.protocol) ? "https://ssl." : "http://www.");
document.write(unescape("%3Cscript src='" + gaJsHost + "google-analytics.com/ga.js' type='text/javascript'%3E%3C/script%3E"));
</script>
<script type="text/javascript">
try {
var pageTracker = _gat._getTracker("UA-16201621-1");
pageTracker._trackPageview();
} catch(err) {}</script>


</body>
</html>
