<%@ Language=VBScript %>
<link href="acss/newstyle.css" rel="stylesheet" type="text/css" />
<link href="assets/css/bootstrap.css" rel="stylesheet" type="text/css" />
<div class="SearchPage" >

<!--#INCLUDE FILE="header.asp"-->
<br>
<%
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
set conn = server.CreateObject ("ADODB.Connection")
set rsSubCat = server.CreateObject("ADODB.Recordset")
set rsMCat = server.CreateObject("ADODB.Recordset")
set rsbrand = server.CreateObject("ADODB.Recordset")
if Request.Cookies("isDealer") <> "" then
 xmsg="You are Logged in as " & Request.Cookies("typename") & " : [" & Request.Cookies("dealer_id") & "]"
 end if%>
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
SQLtxt2="select * from brand" 
rsbrand.CursorLocation = adUseClient
rsbrand.Open SQLtxt2,conn,adOpenForwardOnly,adLockReadOnly,adcmdtext
set rsbrand.ActiveConnection = nothing



%>
<html><head><title>Baylan Technology Inc. Search Page</title> 
 <table width="100%" border="0">
  <tr>
  <td><div align="center"> 
  <%  
   'Dim objAd
   'Set objAd = Server.CreateObject("MSWC.AdRotator") 'Create the component
   'Response.Write  objAd.GetAdvertisement ("/banners/adrot.txt") 
   'Set objAd = Nothing 'Destroy the object
   %>
</div>
 </td></tr></table>

<head>
<link rel="stylesheet" href="admin/arshilogin.css">


<body>

<form action="process_search.asp" method=post id=form1 name=form1 class="navbar-form pull-left">
  <table class=TableBorder cellspacing=2 cellpadding=4 width=100% border=0>
    <tbody>
      <tr class=TableHeading>
        <td class=TableHeading valign=top colspan=3>Search Page </td>
      </tr>
      <tr class=TableText>
        <td class=TableText align=right width="184"> &nbsp;</td>
        <td class=TableText width="254">
            &nbsp;</td>
        <td class=text align=center rowspan="6"><font color="#0000FF"><b>Select Your desired options to Search Products.</b></font><br>
                <br>
                <br>
&quot;Don't forget to Select <br>
        at least one Option&quot;</td>
      <tr class=TableText>
        <td class=TableText align=right width="184"> <span class="text"> Look for</span>
            <input type="checkbox" name="p_usetext" value="yes">
        </td>
        <td  width="254">
          <input size=25 name=p_text >
        </td>
        <tr class=TableText>
        <td class=TableText align=right width="184"> &nbsp;</td>
        <td class=TableText width="254">
            &nbsp;</td>
        <tr class=TableText>
        <td class=TableText align=right width="184"> <span class="text">in Group Category</span>
            <input type="checkbox" name="p_usemcat" value="yes">
            <span class="text"> </span></td>
        <td class=TableText width="254">
          <select name="p_MCat">
            <%Do While Not rsMCat.EOF%>
            <option value="'<%=rsMCat("catid")%>'"> <%=rsMCat("GroupDesc")%> </option>
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
            <option value="'<%=rsSubCat("CatCode")%>'"> <%=rsSubCat("Category")%> </option>
            <%rsSubCat.MoveNext 
            Loop%>
          </select>
        </td>
      <tr class=TableText>
        <td class=text align=right width="184"> in Brand
            <input type="checkbox" name="p_usebrand" value="yes">
        </td>
        <td><select name="p_brand">
            <%Do While Not rsbrand.EOF%>
            <option value="'<%=rsbrand("brandid")%>'"> <%=rsbrand("sdesc")%> </option>
            <%rsbrand.MoveNext 
            Loop%>
          </select>
        <td class=TableBORDER width="254"></td>
      <tr class=TableText>
        <td class=TableText align=right colspan=3>
          <input id=IMAGE 
            onClick=submit(); type=image 
            alt="Click to Continue" src="aimages/search2.gif" value=Login 
            border=0 name=IMAGE>
        </td>
      </tr>
    </tbody>
  </table>
</form>
	<br /><br />

   <!--#INCLUDE FILE="footer.asp"-->

</html>


</div>