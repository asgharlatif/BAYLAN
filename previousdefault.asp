<%@ Language=VBScript %>
<!--#INCLUDE FILE="header.asp"-->
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

if Request.Cookies("isDealer") <> ""  then
  SQLtxt="select *,[dbo].Editprice(pga * "& price &") as pga,[dbo].EditPrice(retail* " & price &" ) as retail from vAlldata where status='Features Products' order by sortby " 
else
 SQLtxt="select * from vAlldata where status='Features Products'"
end if
rsOffer.CursorLocation = adUseClient
rsOffer.Open SQLtxt,conn,adOpenForwardOnly,adLockReadOnly,adcmdtext
set rsOffer.ActiveConnection = nothing

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
<html><head><title>Arshi Computer Home Page</title>
<table width="100%" border="0">
  <tr>
  <td><div align="center"> 
  <%  
   Dim objAd
   Set objAd = Server.CreateObject("MSWC.AdRotator") 'Create the component
   'Response.Write  objAd.GetAdvertisement ("/abanners/adrot.txt") 
   Set objAd = Nothing 'Destroy the object
   %>
</div>
 </td></tr></table>
 <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
       <td width="170" class="tborder" valign="top">&nbsp;<img src="aimages/news.gif" width="106" height="25"></td>
       <td width="170" valign="top"  class="tborder"  align="left" height="148">
       <%
       if rsoffer.RecordCount >0 then%>
       <table width="170" border="0" cellspacing="0" cellpadding="0" height="132">
       <tr> 
		<td class="tborder" height="132"> 
                <div align="center">
				<a href="apage2.asp?a=<%="ars"%>&Mcode=<%=rsoffer("catID")%>&mscat=<%=rsoffer("CatCode")%>&pID=<%=rsoffer("productID")%>&name=<%=rsoffer("Item")%>
<%if Request.Cookies("isDealer") <> "" AND Request.Cookies("TypeName") ="Reseller"  OR Request.Cookies("TypeName") ="Cable Net" then %>		
&price=<%=trim(rsoffer("pga"))%>
<%else if Request.Cookies("isDealer") <> "" then%> 
&price=<%=trim(rsoffer("retail"))%>
<%
end if %>
<%end if %>
&pic=<%=rsoffer("picturelarge")%>&stimage=<%=rsoffer("image")%>&mbrand=<%=rsoffer("brand")%>&mwar=<%=rsoffer("warranty")%>>
<img src=<%=rsoffer("picturefeature")%>"width="120" height="130" border="0"></a>	  
	<% response.Write rsoffer("picturefeature") %>
			    </div>
              </td>
            </tr>
          </table>
       <%end if%>
        </td>
<td width="416" height="148" valign="top" class="tborder"> 
         <% if rsoffer.RecordCount >0 then%>
		  <table width="100%" border="0" cellspacing="0" cellpadding="2">
            <tr>
              <td height="23">
                <p class="text-heading"> <%=rsoffer("itemcode")%>&nbsp;<b><a href="apage2.asp?a=<%="ars"%>&Mcode=<%=rsoffer("catID")%>&mscat=<%=rsoffer("CatCode")%>&pID=<%=rsoffer("productID")%>&name=<%=rsoffer("Item")%>
<%if Request.Cookies("isDealer") <> "" AND Request.Cookies("TypeName") ="Reseller"  OR Request.Cookies("TypeName") ="Cable Net" then %>		
&price=<%=trim(rsoffer("pga"))%>

<%else if Request.Cookies("isDealer") <> "" then%> 

&price=<%=trim(rsoffer("retail")) %>
<%
end if %>

<%end if 
%>&pic=<%=rsoffer("picturelarge")%>&stimage=<%=rsoffer("image")%>&mbrand=<%=rsoffer("brand")%>&mwar=<%=rsoffer("warranty")%>" class="text-heading"><%=rsoffer("item")%></a> 
                  </b>
              </td>
            </tr>
            <tr>
              <td><span class="normtext"> <%=rsoffer("description")%><br>
              </span></td>
            </tr>
          </table>
          <br>
            <br>
        </td>
		<%end if%>
	   <td valign="top" bgcolor="#C6CBD6">&nbsp;</td>
      </tr>
</table>
</td>
</tr>
  <script language=JavaScript>
 function getunits(form)
{
aa=form.ICODE.options.value ;
window.location.href;
if (aa == "") {}
					else window.location.href="apage2.asp?"+aa;
}
</script>	
<%
set RSoffer = server.CreateObject("ADODB.Recordset")
sqltxt="select * from valldata order by itemcode" 
RSoffer.CursorLocation = adUseClient
RSoffer.Open Sqltxt,conn,adOpenForwardOnly,adLockReadOnly,adcmdtext
set RSoffer.ActiveConnection = nothing
%>

<table width="100%" border="0" cellspacing="0" cellpadding="0">
<form name=form>

<tr>
<td  WIDTH="756" ><font size="1" face="Verdana, Arial">Quick Find :</font>
      <select name="ICODE" class="text4" onChange = 'getunits(this.form)'>
        <option value="" selected>[Select Product]</option>
        <font size="1" face="Verdana, Arial">
        <%
          Do While Not RSoffer.EOF%>
        <option value="a=ars&Mcode=<%=rsoffer("CATID")%>&mData=<%=catcode%>&head1=<%=RSoffer("groupdesc")%>&head2=<%=RSoffer("category")%>&pID=<%=RSoffer("productID")%>
		<%if Request.Cookies("isDealer") <> "" AND Request.Cookies("TypeName") ="Reseller"  OR Request.Cookies("TypeName") ="Cable Net" then %>		
&price=<%=trim(rsoffer("pga"))%>

<%else if Request.Cookies("isDealer") <> "" then%> 

&price=<%=trim(rsoffer("retail")) %>
<%
end if %>

<%end if %>

		">
        
		
		<!-- This is what will appear in the combo box -->
        <%=RSoffer("itemcode")%> </option>
        <% RSoffer.MoveNext 
             Loop ' keep the loop%>
      </select>
      <font face="Verdana, Arial" size="2"> <a href="javascript: getunits(form);"><img src="/aimages/go.gif" 
		width="15" border="0" alt="Click here to go directly to the product you selected" name="gobutton"></a></font></td>
    <td valign="top" bgcolor="#C6CBD6">&nbsp;</td>
    </form></table>
<td valign="top" height="449" width="100%">
    <table width="100%" border="0" cellspacing="0" cellpadding="0" height="162">
      <tr>
        <td width="39%" valign="top"> 
          <table width="756" border="0" cellspacing="0" cellpadding="0">
            <td colspan="2" height="1" bgcolor="#6666CC"></td>
            <td colspan="2" height="1" bgcolor="#6666CC"></td>
            
            <% 
              y = (rscat.RecordCount)/2 
            
            
            rscat.Movefirst
            bb=rscat("catid")
	
            %>
            <% for x = 1 to y 
              
              SQLt2="select * from vsubCat where catid='"&bb&"'" 
			  set rsdumm2=conn.Execute(sqlt2)
			  msid2=""
			  morecatid_=0
			  if rsdumm2.EOF=false then
			  morecatid_=rsdumm2("catid")
			  morecatcode_=rsdumm2("CATcode")
	          morehead1_=rscat("groupdesc")
	          morehead2_=rsdumm2("category")
			  end if
			  %>
            <tr> 
              <td width="10%" height="2"><img src="aimages/spacer.gif" width="65" height="8"></td>
              <td class=tborder4 width="40%" height="2"><img src="aimages/spacer.gif" width="220" height="8"></td>
              <td class=tborder5 width="10%" height="2"><img src="aimages/spacer.gif" width="65" height="1"></td>
              <td width="40%" height="2"><img src="aimages/spacer.gif" width="221" height="8"></td>
             
            </tr>
            <tr> 
              <%if rscat("myimage") <> "" and trim(rscat("myimage")) <> "PIMAGES\" then%>
                 <td width="10%">&nbsp;<img src="<%=rscat("myimage")%>">&nbsp; 
				  <%else%>
				<td width="10%">&nbsp;<img src="pimages/C/NIMAGE.GIF" border=0>&nbsp;
                  <%end if%>
			  </td>
              <td class=tborder4 width="40%" valign="top"><a href="apage2.asp?pgt=tt&Mcode=<%=rscat("CATID")%>&mData=<%=msID2%>&head1=<%=rscat("groupdesc")%>" class="text-heading"> 
                <%=rscat("groupdesc")%></a>
                
                <br>
                <% for j=1 to 6%>
                   <%
                   if rsdumm2.EOF=false then 
                   %>
				<a href="apage2.asp?Mcode=<%=rsdumm2("CATID")%>&mData=<%=rsdumm2("CATcode")%>&head1=<%=rscat("groupdesc")%>&head2=<%=rsdumm2("category")%>"class="text"> 
                <%Response.Write rsdumm2("category")%></a><BR>
				<%Response.Write " "
                  rsdumm2.MoveNext
                  end if
			  	  next
               
				 if morecatid_ <> "A" & 0 then
                %>
                <a href="apage2.asp?Mcode=<%=morecatid_%>&mData=<%=morecatcode_%>&head1=<%=morehead1_%>"class="desc">More...</a> 
                <%end if%>
                <br>
               
              </td>
                <%
rscat.MoveNext
aa=rscat("catid")

SQLt="select * from vsubcat where catid='"&trim(aa)&"'" 
set rsdumm=conn.Execute(sqlt)
    morecatid=0
	if rsdumm.EOF=false then
              morecatid=rsdumm("catid")
			  morecatcode=rsdumm("CATcode")
	          morehead1=rscat("groupdesc")
	          morehead2=rsdumm("category")
	end if
%>
             <%if rscat("myimage") <> "" and trim(rscat("myimage")) <> "PIMAGES\" then%>
                 <td class=tborder5 width="10%">&nbsp;<img src=<%=rscat("myimage")%>>&nbsp;</td>
				  <%else%>
				   <td class=tborder5 width="10%">&nbsp;<img src="pimages/C/NIMAGE.GIF" border=0>&nbsp;</td>
                  <%end if%> 
			    <td width="40%" valign="top"><a href="apage2.asp?Mcode=<%=rscat("CATID")%>&mData=<%=msID2%>&head1=<%=rscat("groupdesc")%>" class="text-heading"><%=rscat("groupdesc")%></a><br>
                <%
                for i=1 to 6 
                if rsdumm.EOF=false then
                %>
                <a href="apage2.asp?Mcode=<%=rsdumm("CATID")%>&mData=<%=rsdumm("CATcode")%>&head1=<%=rscat("groupdesc")%>&head2=<%=rsdumm("category")%>" class="text"> 
                <%Response.Write rsdumm("category")%></a><BR>
                <%Response.Write "   "
                  rsdumm.MoveNext%>
                <%
                end if
                next 
                %>
                <%
                if morecatid <> "A" & 0 then
                %>
                <a href="apage2.asp?Mcode=<%=morecatid%>&mData=<%=morecatcode%>&head1=<%=morehead1%>" class="desc">More...</a> 
                <%end if%>
                <br>
                
              </td>
               
            </tr>
            <tr> 
              <td colspan="2" height="1" bgcolor="#6666CC"></td>
              <td colspan="2" height="1" bgcolor="#6666CC"></td>
              
            </tr>
            <%
              if rscat.EOF=false then
              rscat.MoveNext
                 if rscat.EOF=false then
                 bb=rscat("catid")
                end if
              end if
              next 
              %>
          </table>
        </td>
        <td width="39%" valign="top" bgcolor="#C6CBD6">&nbsp;</td>
      </tr>
    </table>
</td>
</tr>
<br>
<!--#INCLUDE FILE="count.asp"-->
 <table width="100%" border="0" cellspacing="0" cellpadding="0" height="1" background="aimages/dot_grey1.gif">
    <tr> 
    <td height="4"></td>
    </tr>
    </table>
    <table border="0" cellspacing="0" cellpadding="0" width="100%">
    <hr color=#000000 size=1 width="650" align="center" noShade>
    <tr> 
    <td width="100%" height="21" id=msviRegionGradient2 
    style="FILTER: progid:DXImageTransform.Microsoft.Gradient(startColorStr='#FFFFFF', endColorStr='#6487DB', gradientType='1')">	<p align="center" class="font">ARSHI COMPUTER Copyright &copy; 2007</p></td>
	</td>
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
<SCRIPT language=JavaScript src="scripts/buzz.js" type=text/javascript></SCRIPT>
<SCRIPT language=JavaScript src="scripts/dynlayer.js" type=text/javascript></SCRIPT>
<SCRIPT language=JavaScript src="scripts/buzz_scroll.js" type=text/javascript></SCRIPT>
<%
'set rsMcat = server.CreateObject("ADODB.Recordset")
'SQLtext="select * from news order by newsID" 
'rsMcat.CursorLocation = adUseClient
'rsMcat.Open SQLtext,conn,adOpenForwardOnly,adLockReadOnly,adcmdtext
'set rsMcat.ActiveConnection = nothing
%>
 
<SCRIPT language=JavaScript type=text/javascript>
<!--
function init() {
	ticker.activate();
}
//NewsTicker(x,y,width,height)
 
ticker = new NewsTicker(0,160,160,90);//controls the location of the ticker window

ticker.addItem(i0);
ticker.addItem(i1);
ticker.addItem(i2);
ticker.addItem(i3);
ticker.addItem(i4);
ticker.addItem(i5);
ticker.addItem(i6);

ticker.build();
writeCSS(ticker.css);


document.write(ticker.div);
if (navigator.appName != "Opera") init();
-->
</SCRIPT>

</html>
</body>
