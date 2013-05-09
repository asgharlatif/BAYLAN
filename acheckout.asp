<%@ Language=VBScript %>
<html>
<head>
<title>Checkout Step-1</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script Language="JavaScript">

                     <!--
                     function Form1_Validator(theForm)
                     {

                     var alertsay = ""; 
                     
                     if (theForm.loginname.value == "")
                     {
                     alert("You must Enter a User ID");
                     theForm.loginname.focus();
                     return (false);
                      }
                     if (theForm.p_password.value == "")
                     {
                     alert("You must Enter Password");
                     theForm.p_password.focus();
                     return (false);
                      }

                     return (true);
                     }

                     //-->
                     
           
                     </script>
                     
                     



                     
                     

<link rel="stylesheet" href="/admin/arshilogin.css" type="text/css">
</head>
<%
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3


set conn = server.CreateObject ("ADODB.Connection")
set rsUser = server.CreateObject("ADODB.Recordset")
set rsUserAdd = server.CreateObject("ADODB.Recordset")
set rsUserAdd2 = server.CreateObject("ADODB.Recordset")
bid=request("bid")


%>
<!--#INCLUDE FILE="dbconnect.asp"-->

<%
conn.Open con
zcode=request("mIDs") 
mrange=split(zcode,"-")
for i =0 to ubound(mrange)
	checkID=mrange(0)   'Customer id who is not a dealer
	bid2=mrange(1)  'Basketid
next


'checkID=request("checkID")
if Request.Cookies("isDealer") <> "" then
chkID=Request.Cookies("isDealer") 


if chkid <> "" then
session("msessiona")=chkid
end if


'Response.Write chkID
sqltext="select * from Adealers where DealerCode='"&chkid&"'" 
rsUser.CursorLocation = adUseClient
rsUser.Open SQLtext,conn,adOpenForwardOnly,adLockReadOnly,adcmdtext
set rsUser.ActiveConnection = nothing
mChkid=rsuser("dealerID")  'DealerID

sqltext2="select * from Ashippingaddress where Dealerid='"&mChkid&"'" 
rsUserAdd.CursorLocation = adUseClient
rsUserAdd.Open SQLtext2,conn,adOpenForwardOnly,adLockReadOnly,adcmdtext
set rsUserAdd.ActiveConnection = nothing

if (not rsUserAdd.EOF) then
	sShipAdd=rsUserAdd("ShippingAddressID") 'to be used in cartchekout.asp to store in order file
	scompany=rsUserAdd("company")
	scontact=rsUserAdd("contactp")
	sAdd1=rsUserAdd("add1")
	sAdd2=rsUserAdd("add2")
	sEmail=rsUserAdd("email")
	sCity=rsUserAdd("city")
	sPh=rsUserAdd("ph")
	sFax=rsUserAdd("fax")
	sZip=rsUserAdd("postalcode")
	sCountry=rsUserAdd("country")
call addresses()
end if


if rsUserAdd.EOF then
	sameAddress="same"
	regsql2="insert into shippingAddress values('"& mchkid &"','0','"& rsUser("Email") &"','"& rsuser("Add1") &"','" & rsUser("Add2") &"','"& rsUser("city") &"','" &rsUser("ph") &"','"& rsuser("fax") &"','"& rsuser("PostalCode") &"','"& rsUser("country") &"')"
	Response.Write regsql2
'	conn.Execute regsql2
   
	sqltext2="select * from Ashippingaddress where customerid='"&checkid&"'" 
	rsUserAdd2.CursorLocation = adUseClient
	rsUserAdd2.Open SQLtext2,conn,adOpenForwardOnly,adLockReadOnly,adcmdtext
	set rsUserAdd2.ActiveConnection = nothing

	sShipAdd=rsUserAdd2("ShippingAddressID") 'to be used in cartchekout.asp to store in order file
	scompany=rsUserAdd2("company")
	scontact=rsUserAdd("contactp")
	sAdd1=rsUserAdd2("add1")
	sAdd2=rsUserAdd2("add2")
	sEmail=rsUserAdd2("email")
	sCity=rsUserAdd2("city")
	sPh=rsUserAdd2("ph")
	sFax=rsUserAdd2("fax")
	sZip=rsUserAdd2("postalcode")
	sCountry=rsUserAdd2("country")
call addresses()

end if
              
'----------------------
else

'checkid=4

'Response.Write "memberid is = " & Request.Cookies("isMember") 
checkid=Request.Cookies("isMember") 
if checkid <> "" then
'Response.Write "aaaaaaaaaaaaaaaaaaaaaaa"

sqltext="select * from Acustomers where customerid='"&checkid&"'" 
rsUser.CursorLocation = adUseClient
rsUser.Open SQLtext,conn,adOpenForwardOnly,adLockReadOnly,adcmdtext
set rsUser.ActiveConnection = nothing
mChkid=rsuser("CustomerID")  'CustomerID

if mchkid <> "" then
session("msessiona")=mchkid
end if


sqltext2="select * from Ashippingaddress where customerid='"&checkid&"'" 
rsUserAdd.CursorLocation = adUseClient
rsUserAdd.Open SQLtext2,conn,adOpenForwardOnly,adLockReadOnly,adcmdtext
set rsUserAdd.ActiveConnection = nothing

if (not rsUserAdd.EOF) then
sShipAdd=rsUserAdd("ShippingAddressID") 'to be used in cartchekout.asp to store in order file
scompany=rsUserAdd("company")
scontact=rsUserAdd("contactp")
sAdd1=rsUserAdd("add1")
sAdd2=rsUserAdd("add2")
sEmail=rsUserAdd("email")
sCity=rsUserAdd("city")
sPh=rsUserAdd("ph")
sFax=rsUserAdd("fax")
sZip=rsUserAdd("postalcode")
sCountry=rsUserAdd("country")
end if

if rsUserAdd.EOF then
sameAddress="same"
regsql2="insert into AshippingAddress values('0','"& checkid &"','"& rsUser("Email") &"','"& rsuser("Add1") &"','" & rsUser("Add2") &"','"& rsUser("city") &"','" &rsUser("ph") &"','"& rsuser("fax") &"','"& rsuser("zip") &"','"& rsUser("country") &"')"
'Response.Write regsql2
conn.Execute regsql2

sqltext2="select * from Ashippingaddress where dealerid='"&checkid&"'" 
rsUserAdd2.CursorLocation = adUseClient
rsUserAdd2.Open SQLtext2,conn,adOpenForwardOnly,adLockReadOnly,adcmdtext
set rsUserAdd2.ActiveConnection = nothing

sShipAdd=rsUserAdd2("ShippingAddressID") 'to be used in cartchekout.asp to store in order file
scompany=rsUserAdd2("company")
scontact=rsUserAdd("contactp")
sAdd1=rsUserAdd2("add1")
sAdd2=rsUserAdd2("add2")
sEmail=rsUserAdd2("email")
sCity=rsUserAdd2("city")
sPh=rsUserAdd2("ph")
sFax=rsUserAdd2("fax")
sZip=rsUserAdd2("postalcode")
sCountry=rsUserAdd2("country")
end if
end if
end if ' enf of cookie check

%>

<body bgcolor="#FFFFFF" text="#000000">
<%if checkid = "" then%>
<form action="acheckoutID.asp?bid=<%=request("bid")%>" method="post" onSubmit="return Form1_Validator(this)" name="Form1">




<table width="80%" border="1" cellspacing="0" cellpadding="0" height="114">
  
  
<%if chkID ="" then%>
  <tr> 
    <td bgcolor="#336699" width="44%" height="19">&nbsp;</td>
    <td bgcolor="#336699" width="56%" align="right" height="19"><font color="#FFFFFF"><span class="text">If 
      not member</span> <a href="aimages/doublearrow.gif?act=signup" class="links1">SIGNUP</a> 
      <span class="text">here for free</span>.</font></td>
  </tr>
  <tr> 
    <td width="44%" height="23" bordercolor="#336699" style="border-left-style: solid; border-left-width: 1"><font face="Verdana" size="2"><strong>Enter User ID:</strong></font></td>
    <td width="56%" height="23" bordercolor="#336699" style="border-right-style: solid; border-right-width: 1"> 
      <input maxlength=50 name=loginname size="20">
    </td>
  </tr>
  <tr>
    <td width="44%" height="27" bordercolor="#336699" style="border-left-style: solid; border-left-width: 1"><font face="Verdana" size="2"><strong>Password:</strong></font></td>
    <td width="56%" height="27" bordercolor="#336699" style="border-right-style: solid; border-right-width: 1"> 
      <input type=password maxlength=30 name=p_password size="20">
      <input type="submit" name="login1" value="Login">
    </td>
  </tr>
  <tr> 
    <td width="44%" height="19" bordercolor="#336699" style="border-left-style: solid; border-left-width: 1; border-bottom-style: solid; border-bottom-width: 1">&nbsp;</td>
    <td width="56%" height="19" bordercolor="#336699" style="border-right-style: solid; border-right-width: 1; border-bottom-style: solid; border-bottom-width: 1">
    <span class="text">If you are not member </span><a href="aimages/doublearrow.gif?act=signup" class="links1">SIGNUP</a> 
    <span class="text">First for free.</span> </td>
  </tr>
  <%end if%>
</table>
</form>
<%else%>
<br>
<form action="acartCheckout.asp?cartida=<%=bid2%>&customerID=<%=mChkid%>&shipID=<%=sShipAdd%>&ctype=<%="Member"%>" method=Post>
<input type="hidden" name="bids" value=<%=Request.QueryString("bid")%>>

<table cellspacing=0 cellpadding=3 width=522 border=0 height="547">
  <tbody> 
  <tr> 
    <td  valign=center align=middle width="516" 
                height="1" bgcolor="#336699" colspan="2"></td>
  </tr>
  <tr> 
    <td  valign=center class="links1" width="258" 
                height="1"> </td>
    <td  valign=center align=right width="258" 
                height="7"> </td>
  </tr>
 
    <td  width="258" height="20" bgcolor="#336699"> 
      <p><b><font face="Verdana" size="2" color="#FFFFFF">About You:</font></b></p>
    </td>
    <td  width="258" height="20" bgcolor="#336699" align="right">&nbsp;</td>
  </tr>
  <tr> 
    <td width="516" height="431" colspan="2"> 
      <table width="100%" cellspacing="0" height="227" style="border-left-width: 0">
       
          <td class=cleartext width=172 height="13" style="border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1"></td>
          <td width=337 height="13" style="border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1"></td>
        </tr>
        <tr valign=top> 
          <td class=cleartext width=172 height="23" bordercolor="#336699" style="border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1"><font face="Verdana" size="2"><strong>Company 
            name: </strong></font></td>
          
          <td width=337 height="23" bordercolor="#336699" style="border-right-style: solid; border-right-width: 1; border-top-style: solid; border-top-width: 1"> 
         <input maxlength=50 
                        name=company size="45" readonly value="<%=rsUser("company")%>">
          </td>
        </tr>
        <tbody> 
        <tr valign=top> 
          <td class=cleartext width=172 height="23" bordercolor="#336699" style="border-left-style: solid; border-left-width: 1"><font face="Verdana" size="2"><strong>First 
            name: </strong> </font></td>
          <td width=337 height="23" bordercolor="#336699" style="border-right-style: solid; border-right-width: 1"> 
            <input maxlength=50 name=firstname readonly size="45" value="<%=rsUser("fname")%>">
          </td>
        </tr>
        <tr valign=top> 
          <td class=cleartext width=172 height="23" bordercolor="#336699" style="border-left-style: solid; border-left-width: 1"><font face="Verdana" size="2"><strong>Last 
            name: </strong> </font></td>
          <td width=337 height="23" bordercolor="#336699" style="border-right-style: solid; border-right-width: 1"> 
            <input maxlength=50  name=lastname readonly size="45" value="<%=rsUser("lname")%>">
          </td>
        </tr>
            <tr valign=top> 
          <td class=cleartext width=172 height="23" bordercolor="#336699" style="border-left-style: solid; border-left-width: 1"><font face="Verdana" size="2"><strong>GST Number: </strong> </font></td>
          <td width=337 height="23" bordercolor="#336699" style="border-right-style: solid; border-right-width: 1"> 
            <input maxlength=50  name=gstnumber readonly size="45" value="<%=rsUser("gstnumber")%>">
          </td>
        </tr>
		 <tr valign=top> 
          <td class=cleartext width=172 height="23" bordercolor="#336699" style="border-left-style: solid; border-left-width: 1"><font face="Verdana" size="2"><strong>NTN Number: </strong> </font></td>
          <td width=337 height="23" bordercolor="#336699" style="border-right-style: solid; border-right-width: 1"> 
            <input maxlength=50  name=ntnnumber readonly size="45" value="<%=rsUser("ntnnumber")%>">
          </td>
        </tr>
		
		<tr valign=top> 
          <td class=cleartext width=172 bgcolor="#336699" height="19" style="border-left-style: solid; border-left-width: 1" bordercolor="#336699"> 
            <strong><font face="Verdana" size="2" color="#FFFFFF">Billing Address</font></strong></td>
          <td width=337 bgcolor="#336699" height="19" style="border-right-style: solid; border-right-width: 1" bordercolor="#336699">&nbsp;</td>
        </tr>
        <tr valign=top> 
          <td class=cleartext width=172 height="3" bordercolor="#336699" style="border-left-style: solid; border-left-width: 1" bordercolorlight="#336699"> 
            <p align="left"><font face="Verdana" size="2"><b>Address:</b></font> 
          </td>
          <td width=337 height="3" style="border-right-style: solid; border-right-width: 1" bordercolor="#336699"><font face="Verdana" size="2"> 
            <input class=mono maxlength=45 
                        name=add1 size="45" readonly value="<%=rsUser("add1")%>">
            </font></td>
        </tr>
        <tr valign=top> 
          <td class=cleartext width=172 height="12" style="border-left-style: solid; border-left-width: 1" bordercolor="#336699"> 
            <p align="left"><font face="Verdana" size="2"><b>Street:</b></font> 
          </td>
          <td width=337 height="12" style="border-right-style: solid; border-right-width: 1" bordercolor="#336699"><font face="Verdana" size="2"> 
            <input class=mono maxlength=45 
                        name=add2 size="45" readonly value="<%=rsUser("add2")%>">
            </font></td>
        </tr>
        <tr valign=top> 
          <td class=cleartext width=172 height="18" style="border-left-style: solid; border-left-width: 1" bordercolor="#336699"> 
            <p align="left"><font face="Verdana" size="2"><b>Postal/Zip Code:</b></font> 
          </td>
          <td width=337 height="18" style="border-right-style: solid; border-right-width: 1" bordercolor="#336699"><font face="Verdana" size="2"> 
            <input class=mono maxlength=30 
                        name=zip size="20" readonly value="<%=rsUser("zip")%>">
            </font></td>
        </tr>
        <tr valign=top> 
          <td class=cleartext width=172 height="1" style="border-left-style: solid; border-left-width: 1" bordercolor="#336699"> 
            <p align="left"><strong><font face="Verdana" size="2">City:</font></strong> 
          </td>
          <td width=337 height="1" style="border-right-style: solid; border-right-width: 1" bordercolor="#336699" ><font face="Verdana" size="2">
            <input class=mono maxlength=30 
                        name=city size="20" readonly value="<%=rsUser("city")%>">
            </font> </td>
        </tr>
        
		<tr valign=top> 
          <td class=cleartext width=172 height="1" style="border-left-style: solid; border-left-width: 1" bordercolor="#336699"> 
            <p align="left"><strong><font face="Verdana" size="2">Country:</font></strong> 
          </td>
          <td width=337 height="1" style="border-right-style: solid; border-right-width: 1" bordercolor="#336699" ><font face="Verdana" size="2">
            <input class=mono maxlength=30 
                        name=country size="20" readonly value="<%=rsUser("country")%>">
            </font> </td>
        </tr>
		
		<tr valign=top> 
          <td class=cleartext width=172 height="10" style="border-left-style: solid; border-left-width: 1" bordercolor="#336699"> 
            <p align="left"><strong><font face="Verdana" size="2">Email:</font></strong> 
          </td>
          <td width=337 height="10" style="border-right-style: solid; border-right-width: 1" bordercolor="#336699"><font face="Verdana" size="2" > 
            <input class=mono maxlength=30 
                        name=email size="20" readonly value="<%=rsUser("email")%>">
            </font></td>
        </tr>
        <tr valign=top> 
          <td class=cleartext width=172 height="2" style="border-left-style: solid; border-left-width: 1" bordercolor="#336699"> 
            <p align="left"><strong><font face="Verdana" size="2">Mobile.: </font></strong> 
          </td>
          <td width=337 height="2" style="border-right-style: solid; border-right-width: 1" bordercolor="#336699"><font face="Verdana" size="2" > 
            <input class=mono maxlength=30 
                        name=mobilet size="20" readonly value="<%=rsUser("mobilet")%>">
            </font></td>
        </tr>
		<tr valign=top> 
          <td class=cleartext width=172 height="2" style="border-left-style: solid; border-left-width: 1" bordercolor="#336699"> 
            <p align="left"><strong><font face="Verdana" size="2">Phone NO.: </font></strong> 
          </td>
          <td width=337 height="2" style="border-right-style: solid; border-right-width: 1" bordercolor="#336699"><font face="Verdana" size="2" > 
            <input class=mono maxlength=30 
                        name=ph size="20" readonly value="<%=rsUser("ph")%>">
            </font></td>
        </tr>
        <tr valign=top> 
          <td class=cleartext width=172 height="4" style="border-left-style: solid; border-left-width: 1; border-bottom-style: solid; border-bottom-width: 1" bordercolor="#336699"> 
            <p align="left"><strong><font face="Verdana" size="2">Fax No. :</font></strong> 
          </td>
          <td width=337 height="4" style="border-right-style: solid; border-right-width: 1; border-bottom-style: solid; border-bottom-width: 1" bordercolor="#336699"><font face="Verdana" size="2"> 
            <input class=mono maxlength=30 
                        name=fax size="20" readonly value="<%=rsUser("fax")%>">
            </font></td>
        </tr>
        <tr valign=top> 
          <td class=cleartext width=172 height="1" style="border-top-style: solid; border-top-width: 1" bordercolor="#336699"></td>
          <td width=337 height="1" style="border-top-style: solid; border-top-width: 1" bordercolor="#336699"></td>
        </tr>
        <tr valign=top> 
          <td class=cleartext width=172 height="5"></td>
          <td width=337 height="3"></td>
        </tr>
        <tr> 
          <td class=cleartext width=172 bgcolor="#336699" height="19" style="border-left-style: solid; border-left-width: 1" bordercolor="#336699"> 
            <strong><font face="Verdana" size="2" color="#FFFFFF">Shipping Address</font></strong></td>
          <td width=337 bgcolor="#336699" height="19" style="border-right-style: solid; border-right-width: 1" bordercolor="#336699"> 
          <%if sameaddress <> "" then%>  
            <font color="#FFFFFF"><b class="text"> 
            <input type="checkbox" name="sameAddress" value="SameAddress" checked>
            I've same Shipping Address as my Billing Address.</b></font>
           <%else%>
		   <font color="#FFFFFF"><b class="text">
              <input type="checkbox" name="sameAddress" value="SameAddress" checked>
            I've same Shipping Address as my Billing Address.</b></font>
            <%end if%>
            </td>
        </tr>
        <tr> 
          <td class=cleartext width=172 height="3" bordercolor="#336699" style="border-left-style: solid; border-left-width: 1" bordercolorlight="#336699"><font face="Verdana" size="2"><strong>Company 
              name:</strong></font></td>
            <td width=337 height="3" style="border-right-style: solid; border-right-width: 1" bordercolor="#336699"><font face="Verdana" size="2">
              <input class=mono maxlength=45 
                        name=scompany size="45" value="<%=scompany%>">
              </font></td>
          </tr>
		  <tr>
            <td class=cleartext width=172 height="3" bordercolor="#336699" style="border-left-style: solid; border-left-width: 1" bordercolorlight="#336699"><font face="Verdana" size="2"><b>Name</b></font></td>
            <td width=337 height="3" style="border-right-style: solid; border-right-width: 1" bordercolor="#336699"><font face="Verdana" size="2">
              <input class=mono maxlength=45 
                        name=scontact size="45" value="<%=scontact%>">
              </font></td>
          </tr>
		 
		 
		  <td class=cleartext width=172 height="3" bordercolor="#336699" style="border-left-style: solid; border-left-width: 1" bordercolorlight="#336699"> 
            <p align="left"><font face="Verdana" size="2"><b>Address:</b></font> 
          </td>
          <td width=337 height="3" style="border-right-style: solid; border-right-width: 1" bordercolor="#336699"><font face="Verdana" size="2"> 
            <input class=mono maxlength=45 
                        name=sadd1 size="45" value="<%=sAdd1%>">
            </font></td>
        </tr>
        <tr> 
          <td class=cleartext width=172 height="12" style="border-left-style: solid; border-left-width: 1" bordercolor="#336699"> 
            <p align="left"><font face="Verdana" size="2"><b>Street:</b></font> 
          </td>
          <td width=337 height="12" style="border-right-style: solid; border-right-width: 1" bordercolor="#336699"><font face="Verdana" size="2"> 
            <input class=mono maxlength=45 
                        name=sadd2 size="45"  value="<%=sAdd2%>">
            </font></td>
        </tr>
        <tr> 
          <td class=cleartext width=172 height="18" style="border-left-style: solid; border-left-width: 1" bordercolor="#336699"> 
            <p align="left"><font face="Verdana" size="2"><b>Postal/Zip Code:</b></font> 
          </td>
          <td width=337 height="18" style="border-right-style: solid; border-right-width: 1" bordercolor="#336699"><font face="Verdana" size="2"> 
            <input class=mono maxlength=30 
                        name=szip size="20" value="<%=sZip%>">
            </font></td>
        </tr>
		 <tr> 
            <td class=cleartext width=172 height="1" style="border-left-style: solid; border-left-width: 1" bordercolor="#336699"> 
              <p align="left"><strong><font face="Verdana" size="2">City:</font></strong> 
            </td>
            <td width=337 height="1" style="border-right-style: solid; border-right-width: 1" bordercolor="#336699" > 
              <select name="scity" >
                <%
				set RCity=server.CreateObject("adodb.recordset")
                Rcity.CursorLocation=adUseClient
                set RCity=Conn.Execute("select destination from Adestinations")
                'set Rcity.ActiveConnection=nothing
                while Rcity.EOF=false
                %>
                %>
                <option value=<%=Rcity("destination")%> selected><%=Rcity("destination")%></option>
                <%
                Rcity.MoveNext
                wend
                %>
                <option value="<%=scity%>"selected><%=scity%></option>
              </select>
              <font face="Verdana" size="2" > </font> </td>
          </tr>
		 <td class=cleartext width=172 height="10" style="border-left-style: solid; border-left-width: 1" bordercolor="#336699"><strong><font face="Verdana" size="2">Country</font></strong></td>
            <td width=337 height="10" style="border-right-style: solid; border-right-width: 1" bordercolor="#336699"><font face="Verdana" size="2" > 
              <select name="scountry">
                <option value="[select one]" selected>[select one]</option>
                <%
			  set RCountry=server.CreateObject("adodb.recordset")
	RCountry.CursorLocation = adUseClient
	RCountry.Open   "select * from ACountry",conn,adOpenForwardOnly,adLockReadOnly,adcmdtext
   
              while RCountry.EOF=false
        
              %>
                <option value="<%=RCountry("country")%>"><%=RCountry("country")%></option>
                <%
              Rcountry.MoveNext
              wend
              %>
                <option value="<%=scountry%>"selected><%=scountry%></option>
              </select>
              </font></td>
          </tr>
		
		
                 <tr> 
          <td class=cleartext width=172 height="10" style="border-left-style: solid; border-left-width: 1" bordercolor="#336699"> 
            <p align="left"><strong><font face="Verdana" size="2">Email:</font></strong> 
          </td>
          <td width=337 height="10" style="border-right-style: solid; border-right-width: 1" bordercolor="#336699"><font face="Verdana" size="2" > 
            <input class=mono maxlength=30 
                        name=semail size="20"  value="<%=sEmail%>">
            </font></td>
        </tr>
        <tr> 
          <td class=cleartext width=172 height="2" style="border-left-style: solid; border-left-width: 1" bordercolor="#336699"> 
            <p align="left"><strong><font face="Verdana" size="2">Phone NO.: </font></strong> 
          </td>
          <td width=337 height="2" style="border-right-style: solid; border-right-width: 1" bordercolor="#336699"><font face="Verdana" size="2" > 
            <input class=mono maxlength=30 
                        name=sph size="20" value="<%=sph%>">
            </font></td>
        </tr>
        <tr> 
          <td class=cleartext width=172 height="4" style="border-left-style: solid; border-left-width: 1; border-bottom-style: solid; border-bottom-width: 1" bordercolor="#336699"> 
            <p align="left"><strong><font face="Verdana" size="2">Fax No. :</font></strong> 
          </td>
          <td width=337 height="3" style="border-right-style: solid; border-right-width: 1; border-bottom-style: solid; border-bottom-width: 1" bordercolor="#336699"><font face="Verdana" size="2"> 
            <input class=mono maxlength=30 
                        name=sfax size="20" value="<%=sfax%>">
            </font></td>
        </tr>
        <tr valign=top> 
          <td class=cleartext width=172 height="11" align="right" style="border-bottom-style: solid; border-bottom-width: 1"></td>
          <td width=337 height="1" class="text" style="border-bottom-style: solid; border-bottom-width: 1"></td>
        </tr>
        <tr valign=top bgcolor="#FFFFFF" align="right"> 
          <td class=cleartext colspan="2" height="26"> <font face="Verdana" size="2"> 
            <input type=submit value="Next &gt;&gt;" name=submit>
            </font></td>
        </tr>
        </tbody> 
      </table>
    </td>
  </tr>
  <tr name="spanimg"><img height=1 src="" width=1 border=0 
                name=sessionid> </tr>
</table>
</form>
<%end if%>
</body>
</html>

<%
' onSubmit="return Form1_Validator(this)" name="Form1"
sub addresses()%>



<br>
<form action="acartCheckout.asp?cartida=<%=request("bid")%>&dealerID=<%=mChkid%>&shipID=<%=sShipAdd%>&ctype=<%="Dealer"%>" method="post">

<input type="hidden" name="bids" value=<%=Request.QueryString("bid")%>>

<table cellspacing=0 cellpadding=3 width=522 border=0 height="547">
  <tbody> 
  <tr> 
    <td  valign=center align=middle width="516" 
                height="1" bgcolor="#336699" colspan="2"></td>
  </tr>
  <tr> 
    <td  valign=center class="links1" width="258" 
                height="1"> </td>
    <td  valign=center align=right width="258" 
                height="7"> </td>
  </tr>
  <tr valign=top> 
    <td  width="258" height="20" bgcolor="#336699"> 
      <p><b><font face="Verdana" size="2" color="#FFFFFF">About You:</font></b></p>
    </td>
    <td  width="258" height="20" bgcolor="#336699" align="right">&nbsp;</td>
  </tr>
  <tr> 
    <td width="516" height="431" colspan="2"> 
      <table width="100%" cellspacing="0" height="227" style="border-left-width: 0">
        <tr valign=top> 
          <td class=cleartext width=172 height="13" style="border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1"></td>
          <td width=337 height="13" style="border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1"></td>
        </tr>
        <tr valign=top> 
          <td class=cleartext width=172 height="23" bordercolor="#336699" style="border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1"><font face="Verdana" size="2"><strong>Company 
            name: </strong></font></td>
          <td width=337 height="23" bordercolor="#336699" style="border-right-style: solid; border-right-width: 1; border-top-style: solid; border-top-width: 1"> 
            <input maxlength=50 
                        name=company size="45" readonly value="<%=rsUser("company")%>">
          </td>
        </tr>
        <tbody> 
        <tr valign=top> 
          <td class=cleartext width=172 height="23" bordercolor="#336699" style="border-left-style: solid; border-left-width: 1"><font face="Verdana" size="2"><strong>First 
            name: </strong> </font></td>
          <td width=337 height="23" bordercolor="#336699" style="border-right-style: solid; border-right-width: 1"> 
            <input maxlength=50 name=firstname readonly size="45" value="<%=rsUser("firstname")%>">
          </td>
        </tr>
        <tr valign=top> 
          <td class=cleartext width=172 height="23" bordercolor="#336699" style="border-left-style: solid; border-left-width: 1"><font face="Verdana" size="2"><strong>Last 
            name: </strong> </font></td>
          <td width=337 height="23" bordercolor="#336699" style="border-right-style: solid; border-right-width: 1"> 
            <input maxlength=50  name=lastname readonly size="45" value="<%=rsUser("lastname")%>">
          </td>
        </tr>
         <tr valign=top> 
          <td class=cleartext width=172 height="23" bordercolor="#336699" style="border-left-style: solid; border-left-width: 1"><font face="Verdana" size="2"><strong>GST Number: </strong> </font></td>
          <td width=337 height="23" bordercolor="#336699" style="border-right-style: solid; border-right-width: 1"> 
            <input maxlength=50  name=gstnumber readonly size="45" value="<%=rsUser("gstnumber")%>">
          </td>
        </tr>
		 <tr valign=top> 
          <td class=cleartext width=172 height="23" bordercolor="#336699" style="border-left-style: solid; border-left-width: 1"><font face="Verdana" size="2"><strong>NTN Number: </strong> </font></td>
          <td width=337 height="23" bordercolor="#336699" style="border-right-style: solid; border-right-width: 1"> 
            <input maxlength=50  name=ntnnumber readonly size="45" value="<%=rsUser("ntnnumber")%>">
          </td>
        </tr>
		
		<tr valign=top> 
          <td class=cleartext width=172 bgcolor="#336699" height="19" style="border-left-style: solid; border-left-width: 1" bordercolor="#336699"> 
            <strong><font face="Verdana" size="2" color="#FFFFFF">Billing Address</font></strong></td>
          <td width=337 bgcolor="#336699" height="19" style="border-right-style: solid; border-right-width: 1" bordercolor="#336699">&nbsp;</td>
        </tr>
        <tr valign=top> 
          <td class=cleartext width=172 height="3" bordercolor="#336699" style="border-left-style: solid; border-left-width: 1" bordercolorlight="#336699"> 
            <p align="left"><font face="Verdana" size="2"><b>Address:</b></font> 
          </td>
          <td width=337 height="3" style="border-right-style: solid; border-right-width: 1" bordercolor="#336699"><font face="Verdana" size="2"> 
            <input class=mono maxlength=45 
                        name=add1 size="45" readonly value="<%=rsUser("add1")%>">
            </font></td>
        </tr>
        <tr valign=top> 
          <td class=cleartext width=172 height="12" style="border-left-style: solid; border-left-width: 1" bordercolor="#336699"> 
            <p align="left"><font face="Verdana" size="2"><b>Street:</b></font> 
          </td>
          <td width=337 height="12" style="border-right-style: solid; border-right-width: 1" bordercolor="#336699"><font face="Verdana" size="2"> 
            <input class=mono maxlength=45 
                        name=add2 size="45" readonly value="<%=rsUser("add2")%>">
            </font></td>
        </tr>
        <tr valign=top> 
          <td class=cleartext width=172 height="18" style="border-left-style: solid; border-left-width: 1" bordercolor="#336699"> 
            <p align="left"><font face="Verdana" size="2"><b>Postal/Zip Code:</b></font> 
          </td>
          <td width=337 height="18" style="border-right-style: solid; border-right-width: 1" bordercolor="#336699"><font face="Verdana" size="2"> 
            <input class=mono maxlength=30 
                        name=zip size="20" readonly value="<%=rsUser("postalCode")%>">
            </font></td>
        </tr>
        <tr valign=top> 
          <td class=cleartext width=172 height="1" style="border-left-style: solid; border-left-width: 1" bordercolor="#336699"> 
            <p align="left"><strong><font face="Verdana" size="2">City:</font></strong> 
          </td>
          <td width=337 height="1" style="border-right-style: solid; border-right-width: 1" bordercolor="#336699" ><font face="Verdana" size="2">
            <input class=mono maxlength=30 
                        name=city size="20" readonly value="<%=rsUser("city")%>">
            </font> </td>
        </tr>
		<tr valign=top> 
          <td class=cleartext width=172 height="1" style="border-left-style: solid; border-left-width: 1" bordercolor="#336699"> 
            <p align="left"><strong><font face="Verdana" size="2">Country:</font></strong> 
          </td>
          <td width=337 height="1" style="border-right-style: solid; border-right-width: 1" bordercolor="#336699" ><font face="Verdana" size="2">
            <input class=mono maxlength=30 
                        name=country size="20" readonly value="<%=rsUser("country")%>">
            </font> </td>
        </tr>
		
        <tr valign=top> 
          <td class=cleartext width=172 height="10" style="border-left-style: solid; border-left-width: 1" bordercolor="#336699"> 
            <p align="left"><strong><font face="Verdana" size="2">Email:</font></strong> 
          </td>
          <td width=337 height="10" style="border-right-style: solid; border-right-width: 1" bordercolor="#336699"><font face="Verdana" size="2" > 
            <input class=mono maxlength=30 
                        name=email size="20" readonly value="<%=rsUser("email")%>">
            </font></td>
        </tr>
		<tr valign=top> 
          <td class=cleartext width=172 height="2" style="border-left-style: solid; border-left-width: 1" bordercolor="#336699"> 
            <p align="left"><strong><font face="Verdana" size="2">Mobile.: </font></strong> 
          </td>
          <td width=337 height="2" style="border-right-style: solid; border-right-width: 1" bordercolor="#336699"><font face="Verdana" size="2" > 
            <input class=mono maxlength=30 
                        name=mobile size="20" readonly value="<%=rsUser("mobile")%>">
            </font></td>
        </tr>
        <tr valign=top> 
          <td class=cleartext width=172 height="2" style="border-left-style: solid; border-left-width: 1" bordercolor="#336699"> 
            <p align="left"><strong><font face="Verdana" size="2">Phone NO.: </font></strong> 
          </td>
          <td width=337 height="2" style="border-right-style: solid; border-right-width: 1" bordercolor="#336699"><font face="Verdana" size="2" > 
            <input class=mono maxlength=30 
                        name=ph size="20" readonly value="<%=rsUser("ph")%>">
            </font></td>
        </tr>
        <tr valign=top> 
          <td class=cleartext width=172 height="4" style="border-left-style: solid; border-left-width: 1; border-bottom-style: solid; border-bottom-width: 1" bordercolor="#336699"> 
            <p align="left"><strong><font face="Verdana" size="2">Fax No. :</font></strong> 
          </td>
          <td width=337 height="4" style="border-right-style: solid; border-right-width: 1; border-bottom-style: solid; border-bottom-width: 1" bordercolor="#336699"><font face="Verdana" size="2"> 
            <input class=mono maxlength=30 
                        name=fax size="20" readonly value="<%=rsUser("fax")%>">
            </font></td>
        </tr>
        <tr valign=top> 
          <td class=cleartext width=172 height="1" style="border-top-style: solid; border-top-width: 1" bordercolor="#336699"></td>
          <td width=337 height="1" style="border-top-style: solid; border-top-width: 1" bordercolor="#336699"></td>
        </tr>
        <tr valign=top> 
          <td class=cleartext width=172 height="5"></td>
          <td width=337 height="3"></td>
        </tr>
        <tr> 
          <td class=cleartext width=172 bgcolor="#336699" height="19" style="border-left-style: solid; border-left-width: 1" bordercolor="#336699"> 
            <strong><font face="Verdana" size="2" color="#FFFFFF">Shipping Address</font></strong></td>
          <td width=337 bgcolor="#336699" height="19" style="border-right-style: solid; border-right-width: 1" bordercolor="#336699"> 
          <%if sameaddress <> "" then%>  
            <font color="#FFFFFF"><b class="text"> 
            <input type="checkbox" name="sameAddress" value="SameAddress" checked>
            <font face="Verdana, Arial" size="2">I've same Shipping Address as 
            my Billing Address.</font></b></font> <font face="Verdana, Arial" size="2"> 
            <%else%>
            <font color="#FFFFFF"><b class="text"> 
            <input type="checkbox" name="sameAddress" value="SameAddress" checked>
            Please Select same Shipping Address as below Address. You want Change 
            Address Do Not Select</b></font></font> 
            <%end if%>
          </td>
        </tr>
        <tr> 
         <td class=cleartext width=172 height="3" bordercolor="#336699" style="border-left-style: solid; border-left-width: 1" bordercolorlight="#336699"><font face="Verdana" size="2"><strong>Company 
              name:</strong></font></td>
            <td width=337 height="3" style="border-right-style: solid; border-right-width: 1" bordercolor="#336699"><font face="Verdana" size="2">
              <input class=mono maxlength=45 
                        name=scompany size="45" value="<%=scompany%>">
              </font></td>
          </tr>
		  <tr>
            <td class=cleartext width=172 height="3" bordercolor="#336699" style="border-left-style: solid; border-left-width: 1" bordercolorlight="#336699"><font face="Verdana" size="2"><b>Name</b></font></td>
            <td width=337 height="3" style="border-right-style: solid; border-right-width: 1" bordercolor="#336699"><font face="Verdana" size="2">
              <input class=mono maxlength=45 
                        name=scontact size="45" value="<%=scontact%>">
              </font></td>
          </tr>
		 
		 
		  <td class=cleartext width=172 height="3" bordercolor="#336699" style="border-left-style: solid; border-left-width: 1" bordercolorlight="#336699"> 
            <p align="left"><font face="Verdana" size="2"><b>Address:</b></font> 
          </td>
          <td width=337 height="3" style="border-right-style: solid; border-right-width: 1" bordercolor="#336699"><font face="Verdana" size="2"> 
            <input class=mono maxlength=45 
                        name=sadd1 size="45" value="<%=sAdd1%>">
            </font></td>
        </tr>
        <tr> 
          <td class=cleartext width=172 height="12" style="border-left-style: solid; border-left-width: 1" bordercolor="#336699"> 
            <p align="left"><font face="Verdana" size="2"><b>Street:</b></font> 
          </td>
          <td width=337 height="12" style="border-right-style: solid; border-right-width: 1" bordercolor="#336699"><font face="Verdana" size="2"> 
            <input class=mono maxlength=45 
                        name=sadd2 size="45"  value="<%=sAdd2%>">
            </font></td>
        </tr>
        <tr> 
          <td class=cleartext width=172 height="18" style="border-left-style: solid; border-left-width: 1" bordercolor="#336699"> 
            <p align="left"><font face="Verdana" size="2"><b>Postal/Zip Code:</b></font> 
          </td>
          <td width=337 height="18" style="border-right-style: solid; border-right-width: 1" bordercolor="#336699"><font face="Verdana" size="2"> 
            <input class=mono maxlength=30 
                        name=szip size="20" value="<%=sZip%>">
            </font></td>
        </tr>
		 <tr> 
            <td class=cleartext width=172 height="1" style="border-left-style: solid; border-left-width: 1" bordercolor="#336699"> 
              <p align="left"><strong><font face="Verdana" size="2">City:</font></strong> 
            </td>
            <td width=337 height="1" style="border-right-style: solid; border-right-width: 1" bordercolor="#336699" > 
              <select name="scity" >
                <%
				set RCity=server.CreateObject("adodb.recordset")
                Rcity.CursorLocation=adUseClient
                set RCity=Conn.Execute("select destination from Adestinations")
                'set Rcity.ActiveConnection=nothing
                while Rcity.EOF=false
                %>
                %>
                <option value=<%=Rcity("destination")%> selected><%=Rcity("destination")%></option>
                <%
                Rcity.MoveNext
                wend
                %>
                <option value="<%=scity%>"selected><%=scity%></option>
              </select>
              <font face="Verdana" size="2" > </font> </td>
          </tr>
		 <td class=cleartext width=172 height="10" style="border-left-style: solid; border-left-width: 1" bordercolor="#336699"><strong><font face="Verdana" size="2">Country</font></strong></td>
            <td width=337 height="10" style="border-right-style: solid; border-right-width: 1" bordercolor="#336699"><font face="Verdana" size="2" > 
              <select name="scountry">
                <option value="[select one]" selected>[select one]</option>
                <%
			  set RCountry=server.CreateObject("adodb.recordset")
	RCountry.CursorLocation = adUseClient
	RCountry.Open   "select * from ACountry",conn,adOpenForwardOnly,adLockReadOnly,adcmdtext
   
              while RCountry.EOF=false
        
              %>
                <option value="<%=RCountry("country")%>"><%=RCountry("country")%></option>
                <%
              Rcountry.MoveNext
              wend
              %>
                <option value="<%=scountry%>"selected><%=scountry%></option>
              </select>
              </font></td>
          </tr>
		
		
                 <tr> 
          <td class=cleartext width=172 height="10" style="border-left-style: solid; border-left-width: 1" bordercolor="#336699"> 
            <p align="left"><strong><font face="Verdana" size="2">Email:</font></strong> 
          </td>
          <td width=337 height="10" style="border-right-style: solid; border-right-width: 1" bordercolor="#336699"><font face="Verdana" size="2" > 
            <input class=mono maxlength=30 
                        name=semail size="20"  value="<%=sEmail%>">
            </font></td>
        </tr>
        <tr> 
          <td class=cleartext width=172 height="2" style="border-left-style: solid; border-left-width: 1" bordercolor="#336699"> 
            <p align="left"><strong><font face="Verdana" size="2">Phone NO.: </font></strong> 
          </td>
          <td width=337 height="2" style="border-right-style: solid; border-right-width: 1" bordercolor="#336699"><font face="Verdana" size="2" > 
            <input class=mono maxlength=30 
                        name=sph size="20" value="<%=sph%>">
            </font></td>
        </tr>
        <tr> 
          <td class=cleartext width=172 height="4" style="border-left-style: solid; border-left-width: 1; border-bottom-style: solid; border-bottom-width: 1" bordercolor="#336699"> 
            <p align="left"><strong><font face="Verdana" size="2">Fax No. :</font></strong> 
          </td>
          <td width=337 height="3" style="border-right-style: solid; border-right-width: 1; border-bottom-style: solid; border-bottom-width: 1" bordercolor="#336699"><font face="Verdana" size="2"> 
            <input class=mono maxlength=30 
                        name=sfax size="20" value="<%=sfax%>">
            </font></td>
        </tr>
        <tr valign=top> 
          <td class=cleartext width=172 height="11" align="right" style="border-bottom-style: solid; border-bottom-width: 1"></td>
          <td width=337 height="1" class="text" style="border-bottom-style: solid; border-bottom-width: 1"></td>
        </tr>
        <tr valign=top bgcolor="#FFFFFF" align="right"> 
          <td class=cleartext colspan="2" height="26"> <font face="Verdana" size="2"> 
            <input type=submit value="Next &gt;&gt;" name=submit>
            </font></td>
        </tr>
        </tbody> 
      </table>
</form>
    </td>
  </tr>
  <tr name="spanimg"><img height=1 src="" width=1 border=0 
                name=sessionid> </tr>
</table>


<%end sub%>