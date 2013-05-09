<head>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language=JavaScript>
window.name = "arshimain";
var imgurl="images/";
</script>

<link rel="stylesheet" type="text/css" href="anylinkmenu.css" />
<link rel="stylesheet" type="text/css" href="assets/css/bootstrap.css" />


 <script>
     function geticode() {

         window.location = 'default.asp?' + document.getElementById('icode').value;
     }

     function getscode() {

         window.location = 'apage2.asp?' + document.getElementById('scode').value;
     }
     
     function getpcode() {

         window.location = 'apage2.asp?' + document.getElementById('pcode').value;
        
     }
     function getitcdcode() {

         window.location = 'apage2.asp?' + document.getElementById('itcd').value;

     }
     function getproductidcode() {

         window.location = 'apage2.asp?' + document.getElementById('itemcode').value;

     }
</script>
<script type="text/javascript" src="menucontents.js"></script>

<script type="text/javascript" src="anylinkmenu.js">



/***********************************************
* AnyLink JS Drop Down Menu v2.0- © Dynamic Drive DHTML code library (www.dynamicdrive.com)
* This notice MUST stay intact for legal use
* Visit Project Page at http://www.dynamicdrive.com/dynamicindex1/dropmenuindex.htm for full source code
***********************************************/
   

</script>

<script type="text/javascript">

//anylinkmenu.init("menu_anchors_class") //Pass in the CSS class of anchor links (that contain a sub menu)
anylinkmenu.init("menuanchorclass")

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
</script>
<script language="JavaScript">

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




<script language="JavaScript">

function userCat()
{
var x1 =  (screen.width-400)/2;
var y1 =  (screen.height-300 )/2;
childWin= open('verify.asp' ,'childWin','screenX='+ x1 + ',screenY=' + y1 + ',left='+ x1 + ',top='+ y1 + ',width=420,height=100,resizable=no,toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no');
}


function giveFocus()
{
//alert (ascii('a') )
//document.homeP.SubmitB.focus()
}




</script>

<script language="JavaScript">
var doImage=doImage;var TType=TType;
function mhHover(tbl,idx,cls){var t,d;if(document.getElementById)t=document.getElementById(tbl);else t=document.all(tbl);if(t==null)return;if(t.getElementsByTagName)d=t.getElementsByTagName("TD");else d=t.all.tags("TD");if(d==null)return;if(d.length<=idx)return;d[idx].className=cls;}
function footerjs(doc){if(doImage==null){var tt=TType==null?"PV":TType;doc.write('<layer visibility="hide"><div style="display:none"></div></layer>');}}
</SCRIPT>

<link rel="stylesheet" href="acss/standard.css" type="text/css">


    <style type="text/css">
        .style4
        {
            COLOR: white;
            FONT-FAMILY: Arial, Helvetica, sans-serif;
            FONT-SIZE: 11px;
            FONT-WEIGHT: bold;
            height: 11px;
        }
    </style>


<script language="javascript" type="text/javascript">

   


</script>



</head>


    

<div class="headerdiv">

<table cellspacing="0" cellpadding="0" width="964px" border="0" style="margin-left:8px">
    <tr>
        <td height="24" colspan="6" align="left" nowrap bgcolor="#6487db" id="actoolbar"
            dir="ltr">
            <table align="right" cellspacing="0" cellpadding="0" border="0">
                <tbody>
                    <tr>
                        <td class="ac0" onmouseover="mhHover('actoolbar', 0, 'ac1')" onmouseout="mhHover('actoolbar', 0, 'ac0')"
                            nowrap>
                            <a href="default.asp" class="btn btn-warning btn-primary">Home</a>
                        </td>
                        <td class="acsep">
                            |
                        </td>
                        <td class="ac0" onmouseover="mhHover('actoolbar', 2, 'ac1')" onmouseout="mhHover('actoolbar', 2, 'ac0')"
                            nowrap>
                            <a href="about_us.asp" class="btn btn-warning btn-primary"><span class="menufontnormal">
                                About Us</span></a>
                        </td>
                        <td class="acsep">
                            |
                        </td>
                        
                        <td class="ac0" onmouseover="mhHover('actoolbar', 6, 'ac1')" onmouseout="mhHover('actoolbar', 6, 'ac0')"
                            nowrap>
                            <a href="contact_us.asp" class="btn btn-warning btn-primary">Contact</a>
                        </td>
                        
                        
                    </tr>
                </tbody>
            </table>
        </td>
    </tr>
    <tr>
    <td class="logodiv">
    &nbsp;
    </td>
    </tr>
   <tr>
   <td>
   <div class="FindDiv">
   <!--#INCLUDE FILE="dbconnect.asp"-->
 
<%

set conn = server.CreateObject ("ADODB.Connection")

conn.Open con 
set rsMainF = server.CreateObject("ADODB.Recordset")


SQLtxt="select catid as Catm,GroupDesc as ElementDescription from VmainCAT where catid in (select catm  from item where brshds='BayLan' group by catm) order by sort_it" 

rsMainF.CursorLocation = adUseClient

''---------------- Category String and RecordSet Preparation
set rssubCatF = server.CreateObject("ADODB.Recordset")

if (request.QueryString("mcode") ="") then
SQLscat="select catcode as Cats,Category as ElementDescription,catid,[desc] from vsubCat order by category" 
else
SQLscat="select catcode as Cats,Category as ElementDescription,catid,[desc] from vsubCat where catid='"& request.QueryString("mcode") &"' order by category" 
end if

rssubCatF.CursorLocation = adUseClient
rssubCatF.Open SQLscat,conn,adOpenForwardOnly,adLockReadOnly,adcmdtext
set rssubCatF.ActiveConnection = nothing
rssubCatF.MoveFirst
''---------------- End Category String and RecordSet Preparation
''---------------- Category String and RecordSet Preparation
set rsitemF = server.CreateObject("ADODB.Recordset")
set rsitemFModel = server.CreateObject("ADODB.Recordset")
set rsitemByCode = server.CreateObject("ADODB.Recordset")

''mData=0139&mcode=013

baylenProduct="  productid in (select itcd from item where brshds='BayLan' group by itcd) "

if (request.QueryString("mcode") <>"") and (request.QueryString("mData") <>"") then
    SQLscat="SELECT  productID as   itcd,CatID as catm, CatCode AS cats, Description AS elementdescription,category,groupdesc,itemcode FROM          vAllData where "& baylenProduct &" and CatID='"& request.QueryString("mcode") &"' and CatCode='"& request.QueryString("mData") &"' and statuscode != '08'  order by Description"    
    SQLscatModel="SELECT  productID as   itcd,CatID as catm, CatCode AS cats, Description AS elementdescription,category,groupdesc,itemcode FROM          vAllData where "& baylenProduct &" and CatID='"& request.QueryString("mcode") &"' and CatCode='"& request.QueryString("mData") &"' and statuscode != '08'  order by itemcode"    
    SQLscatByCode="SELECT  productID as   itcd,CatID as catm, CatCode AS cats, Description AS elementdescription,category,groupdesc,itemcode FROM          vAllData where "& baylenProduct &" and CatID='"& request.QueryString("mcode") &"' and CatCode='"& request.QueryString("mData") &"' and statuscode != '08'  order by productID"    
elseif (request.QueryString("itcd") <>"")  then
    SQLscat="SELECT   productID as   itcd,CatID as catm, CatCode AS cats, Description AS elementdescription,category,groupdesc,itemcode FROM         vAllData where "& baylenProduct &" and productID='"& request.QueryString("itcd") &"' and statuscode != '08'  order by Description"    
    SQLscatModel="SELECT   productID as   itcd,CatID as catm, CatCode AS cats, Description AS elementdescription,category,groupdesc,itemcode FROM         vAllData where "& baylenProduct &" and productID='"& request.QueryString("itcd") &"' and statuscode != '08'  order by itemcode"  
    SQLscatByCode="SELECT   productID as   itcd,CatID as catm, CatCode AS cats, Description AS elementdescription,category,groupdesc,itemcode FROM         vAllData where "& baylenProduct &" and productID='"& request.QueryString("itcd") &"' and statuscode != '08'  order by productID"  
else
    SQLscat="SELECT   productID as   itcd,CatID as catm, CatCode AS cats, Description AS elementdescription,category,groupdesc,itemcode FROM         vAllData where  "& baylenProduct &" and statuscode != '08'  order by Description "
    SQLscatModel="SELECT   productID as   itcd,CatID as catm, CatCode AS cats, Description AS elementdescription,category,groupdesc,itemcode FROM         vAllData where  "& baylenProduct &" and statuscode != '08'  order by itemcode "
    SQLscatByCode="SELECT   productID as   itcd,CatID as catm, CatCode AS cats, Description AS elementdescription,category,groupdesc,itemcode FROM         vAllData where  "& baylenProduct &" and statuscode != '08'  order by productID "
end if

'response.Write(SQLscat)



rsitemF.CursorLocation = adUseClient
rsitemF.Open SQLscat,conn,adOpenForwardOnly,adLockReadOnly,adcmdtext
set rsitemF.ActiveConnection = nothing

rsitemFModel.CursorLocation = adUseClient
rsitemFModel.Open SQLscatModel,conn,adOpenForwardOnly,adLockReadOnly,adcmdtext
set rsitemFModel.ActiveConnection = nothing

rsitemByCode.CursorLocation = adUseClient
rsitemByCode.Open SQLscatByCode,conn,adOpenForwardOnly,adLockReadOnly,adcmdtext
set rsitemByCode.ActiveConnection = nothing



'rsitemF.MoveFirst
''---------------- End Category String and RecordSet Preparation




set rsMainF=conn.Execute(SQLtxt)
   %>
  
    <table class="FindElementTable">
    <tr>

    <td>
       
       &nbsp;<select  onChange = "geticode();" id="icode" name="icode" class="text4" class="ListWidth" >
            <font size="1" face="Verdana, Arial">
            <option value="0"  selected>[Select Main Category]</option>
            </font> 
        <%
        Do While Not rsMainF.EOF%>
            <font size="1" face="Verdana, Arial">
            <%
            if request.QueryString("mcode")=rsMainF("Catm") then
             %>
            <option value="mcode=<%=rsMainF("Catm")%>" selected ><%=rsMainF("ElementDescription")%></option>
            <%else %>
            <option value="mcode=<%=rsMainF("Catm")%>" ><%=rsMainF("ElementDescription")%></option>
            <%end if %>

            </font>
            <% rsMainF.MoveNext 
        Loop %>
        
        </select>

      

    </td>

    <td>
        &nbsp;<select onChange = "getscode();" id="scode" name="scode" class="text4" >
           <option value="0"  selected>[Select Sub Category]</option>
        <%
        Do While Not rssubCatF.EOF%>
             <%
            if request.QueryString("mData")=rssubCatF("Cats") then
             %>
           <option value="mData=<%=rssubCatF("Cats")%>&mcode=<%=rssubCatF("CatId")%>&head2=<%=rssubCatF("ElementDescription")%>&head1=<%=rssubCatF("desc")%>" selected><%=rssubCatF("ElementDescription")%></option>
           <%else %>
            <option value="mData=<%=rssubCatF("Cats")%>&mcode=<%=rssubCatF("CatId")%>&head2=<%=rssubCatF("ElementDescription")%>&head1=<%=rssubCatF("desc")%>"><%=rssubCatF("ElementDescription")%></option>
            <%end if %>

            <% rssubCatF.MoveNext 
        Loop %>
        </select>

    </td>
    <td id="internaltabel">
      <table>
      <tr>
      <td> <select onChange = "getpcode();"  id="pcode" name="pcode"  class="text4">
        
           <option value="0" "selected">[Select By Product Description]</option>
        <%
        Do While Not rsitemF.EOF%>
            <%
            if request.QueryString("itcd")=rsitemF("itcd") then
             %>
            <option value="itcd=<%=rsitemF("itcd")%>&mData=<%=rsitemF("cats")%>&mcode=<%=rsitemF("catm")%>&head1=<%=rsitemF("category")%>&head2=<%=rsitemF("groupdesc")%>" selected  ><%=rsitemF("ElementDescription")%></option>
            <%else %>
            <option value="itcd=<%=rsitemF("itcd")%>&mData=<%=rsitemF("cats")%>&mcode=<%=rsitemF("catm")%>&head1=<%=rsitemF("category")%>&head2=<%=rsitemF("groupdesc")%>"><%=rsitemF("ElementDescription")%></option>
            <%end if %>

            <% rsitemF.MoveNext 
        Loop %>
        </select></td>
        <td><select onChange = "getitcdcode();"  id="itcd" name="itcd"  class="text4">
        
           <option value="0" "selected">[Select By Code]</option>
            <%
        
        if rsitemByCode.recordcount>0 then
        rsitemByCode.movefirst
        end if
        
        Do While Not rsitemByCode.EOF%>
            <%
            if request.QueryString("itcd")=rsitemByCode("itcd") then
             %>
            <option value="itcd=<%=rsitemByCode("itcd")%>&mData=<%=rsitemByCode("cats")%>&mcode=<%=rsitemByCode("catm")%>&head1=<%=rsitemByCode("category")%>&head2=<%=rsitemByCode("groupdesc")%>" selected  ><%=rsitemByCode("itcd")%></option>
            <%else %>
            <option value="itcd=<%=rsitemByCode("itcd")%>&mData=<%=rsitemByCode("cats")%>&mcode=<%=rsitemByCode("catm")%>&head1=<%=rsitemByCode("category")%>&head2=<%=rsitemByCode("groupdesc")%>"><%=rsitemByCode("itcd")%></option>
            <%end if %>

            <% rsitemByCode.MoveNext 
        Loop %>
           </select>
           </td>
        <td>
        
        <select onChange = "getproductidcode();"  id="itemcode" name="itemcode"  class="text4" >
        
           <option value="0" "selected">[Select By  Model]</option>

            <%
        
        if rsitemFModel.recordcount>0 then
        rsitemFModel.movefirst
        end if
        
        Do While Not rsitemFModel.EOF%>
            <%
            if request.QueryString("itcd")=rsitemFModel("itcd") then
             %>
            <option value="itcd=<%=rsitemFModel("itcd")%>&mData=<%=rsitemFModel("cats")%>&mcode=<%=rsitemFModel("catm")%>&head1=<%=rsitemFModel("category")%>&head2=<%=rsitemFModel("groupdesc")%>" selected  ><%=rsitemFModel("itemcode")%></option>
            <%else %>
            <option value="itcd=<%=rsitemFModel("itcd")%>&mData=<%=rsitemFModel("cats")%>&mcode=<%=rsitemFModel("catm")%>&head1=<%=rsitemFModel("category")%>&head2=<%=rsitemFModel("groupdesc")%>"><%=rsitemFModel("itemcode")%></option>
            <%end if %>

            <% rsitemFModel.MoveNext 
        Loop %>

           </select>
           
           
           </td>
      </tr>
      </table>
       
      
       
    </td>
    <td style="padding:2px">
    <a href="#" class="btn btn-success">
    Find
         </a>
    </td>
    
    
    </tr>
    </table>
    

   </div>
   </td>
   </tr>

    
   
</table>


</div>
 