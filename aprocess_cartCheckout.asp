<%@ Language=VBScript %>
<html>
<head>
<link rel="stylesheet" href="admin/arshilogin.css" type="text/css">
</head>
</html>

<%
'strHost = "192.168.0.171"
'strHost = "63.99.209.17"
if session("msessiona")= "" then
	Response.Write "<font face =arial size=2 color=red><b>Your Session has expired</b></font>"
else

Set objMessage  = Server.CreateObject("CDO.Message")

Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3

cartida=request("cartida")


'customerid=Request.QueryString("custid")
dealerid=request("dealid")

xtoday=now()

'if customerid <> "" then
'	dealerID=0
'else
'	customerID=0
'end if

set conn = server.CreateObject ("ADODB.Connection")
set rsCart = server.CreateObject("ADODB.Recordset")
set rsMAX = server.CreateObject("ADODB.Recordset")
set rsOrder= server.CreateObject("ADODB.Recordset")
set rsAddr2= server.CreateObject("ADODB.Recordset")
set rsAddr3= server.CreateObject("ADODB.Recordset")
set rsEmail= server.CreateObject("ADODB.Recordset")
%>
<!--#INCLUDE FILE="dbconnect.asp"-->
<%
conn.Open con 

''''for shippping address 

'Response.Write "<br>address match = " & trim(session("addressmatch")) & "<br>"




if trim(session("addressmatch")) = 1 then


if Request.Cookies("isdealer")	 <> "" or Request.Cookies("ismember") <> ""  then

if Request.Cookies("isdealer") <> "" then
regsql2="insert into shippingAddress " 
regsql2=regsql2+"values('"&Request.Cookies("dealerid")&"','0',0,'"
regsql2=regsql2+""&session("semail")&"','"&session("sadd1")&"','"
regsql2=regsql2+""&session("sadd2")&"','"&session("scity")&"','"&session("sph")&"','"
regsql2=regsql2+""&session("sfax")&"','"&session("szip")&"','"
regsql2=regsql2+""&session("scountry")&"','"&session("scontact")&"','"
regsql2=regsql2+""&session("scompany")&"')"
conn.Execute regsql2
end if

'if Request.Cookies("isMember") <> "" then
'regsql2="insert into shippingAddress " 
'regsql2=regsql2+"values('0','"&Request.Cookies("actid")&"',0,'"
'regsql2=regsql2+""&session("semail")&"','"&session("sadd1")&"','"
'regsql2=regsql2+""&session("sadd2")&"','"&session("scity")&"','"&session("sph")&"','"
'regsql2=regsql2+""&session("sfax")&"','"&session("szip")&"','"
'regsql2=regsql2+""&session("scountry")&"','"&session("scontact")&"','"
'regsql2=regsql2+""&session("scompany")&"')"
'Response.Write "<br> shipping address is = "& regsql2 & "<br>"
'conn.Execute regsql2
'Response.Write "in member"
'end if

set rsmax_sid=server.CreateObject("adodb.recordset")
sqlmax="SELECT  MAX(shippingaddressid) AS Expr1 FROM Ashippingaddress"
rsmax_sid.CursorLocation = adUseClient
rsmax_sid.Open SQLmax,conn,adOpenForwardOnly,adLockReadOnly,adcmdtext
set rsmax_sid.ActiveConnection = nothing
max_sid=rsmax_sid("Expr1")
session("shipaddID")=trim(max_sid)

'Response.Write "<br>max shipping id = " & session("shipaddID") & "<br>"
end if


else

		'if Request.Cookies("isdealer") <> "" then
		set RsDealerSA=server.CreateObject("adodb.recordset")
        set RsDealerSA=conn.Execute("select * from Ashippingaddress where dealerid="&request("dealid")&" and defaulty=1")
        session("scontact")=RsDealerSA("contactp")
        session("sadd1")=RsDealerSA("add1")
        session("sadd2")=RsDealerSA("add2")
        session("scity")=RsDealerSA("city")
		session("szip")=RsDealerSA("postalcode")
		session("scountry")=RsDealerSA("country")
		session("scompany")=RsDealerSA("company")
       
        
		'else
		'set RsDealerSA=server.CreateObject("adodb.recordset")
        'set RsDealerSA=conn.Execute("select * from shippingaddress where customerid="&request("custid")&" and defaulty=1")
        'session("scontact")=RsDealerSA("contactp")
        'session("sadd1")=RsDealerSA("add1")
        'session("sadd2")=RsDealerSA("add2")
        'session("scity")=RsDealerSA("city")
		'session("szip")=RsDealerSA("postalcode")
		'session("scountry")=RsDealerSA("country")
		'session("scompany")=RsDealerSA("company")
		'end if
		RsDealerSA.close
		set RsDealerSA=nothing

end if

'''end shipping address



'if customerid <> "" then
'	Stxt1="select * from Acustomers where Customerid='"&customerid&"'" 
'	rsAddr2.CursorLocation = adUseClient
'	rsAddr2.Open Stxt1,conn,adOpenForwardOnly,adLockReadOnly,adcmdtext
	
'	set rsAddr2.ActiveConnection = nothing
'	if not rsaddr2.EOF then
'	company=rsAddr2("company")
'	fname=rsAddr2("fname")
'	lname=rsAddr2("lname")
'	add1=rsAddr2("add1")
'	add2=rsAddr2("add2")
'	city=rsAddr2("city")
'	zip=rsAddr2("zip")
'	country=rsAddr2("country")
'	memail=rsAddr2("email")
'	end if
'end if

if Dealerid <> "" then
	Stxt1="select * from Adealers where Dealerid='"&Dealerid&"'" 
	rsAddr3.CursorLocation = adUseClient
	rsAddr3.Open Stxt1,conn,adOpenForwardOnly,adLockReadOnly,adcmdtext
	set rsAddr3.ActiveConnection = nothing
	if not rsaddr3.EOF then
	company=rsAddr3("company")
	fname=rsAddr3("firstname")
	lname=rsAddr3("lastname")
	add1=rsAddr3("add1")
	add2=rsAddr3("add2")
	city=rsAddr3("city")
	zip=rsAddr3("postalcode")
	country=rsAddr3("country")
	memail=rsAddr3("email")

   end if
end if

	
	SQLtxt="select * from Avcart where basketid='"&cartida&"'" 
	rsCart.CursorLocation = adUseClient
	rsCart.Open SQLtxt,conn,adOpenForwardOnly,adLockReadOnly,adcmdtext
	set rsCart.ActiveConnection = nothing

   ' Response.Write "<br>shipiing id = " & session("shipaddID") & "<br>"
  'on error resume next
	
	st="insert into Aorders ( dealerid,shippingaddressid,taxrate,taxCharged, "
	st=st + "ptotal,subtotal,total,courier,destination,freight,totalweight,created,completed, "
	st=st + "status,customertype,slipnot,htax,chtax,couriercharges ) "
	
	st=st +  " values ("&dealerID&", "
	st=st +  " "&session("shipaddID")&","&session("mGST")&", "
	st=st +  " "&session("xGST")&","&session("t_tot")&", "
	st=st +  " "&session("ntotal")&","&session("grandTotal")&", "
	st=st +  " '"&session("mCour")&"','"&session("mDest")&"',"&session("freight")&", "
	st=st +  " "&session("wt1")&",'"&xToday&"','"&xToday&"', "
	st=st +  " 1,'"&session("CustomerType")&"',' ', "
	st=st +  " "&session("htax")&","&session("Chtax")&","&session("mfreight")&" )"
	'Response.Write "--------- xgst =  " &  st & "<br>"
	conn.Execute(st)

'Response.Write st

	sqlmax="SELECT  MAX(orderid) AS Expr1 FROM Aorders"
	rsMax.CursorLocation = adUseClient
	rsMax.Open SQLmax,conn,adOpenForwardOnly,adLockReadOnly,adcmdtext
	set rsMax.ActiveConnection = nothing
	mlast=rsMax("Expr1") 'pick last productID to insert in QtyUnit table
'on error resume next
rscart.MoveFirst 
do while (not rsCart.EOF)
ntot=cdbl(rscart("quantity"))*cdbl(rscart("priceEach"))
sqltxt2="insert into AorderDetail values("&mlast&", "&rscart("productid")&","&rscart("quantity")&","&rscart("priceEach")&","&ntot&",'"&rscart("taxable")&"','"&rscart("freeshipping")&"',"&rscart("weight")&")"
conn.Execute(sqltxt2)
rscart.MoveNext
loop
Setxt="select * from Aextras" 
	rsemail.CursorLocation = adUseClient
	rsemail.Open Setxt,conn,adOpenForwardOnly,adLockReadOnly,adcmdtext
	set rsemail.ActiveConnection = nothing
	aa=session("mgst")
	call sendTOclient()
	
	session.Abandon 
	%>
	<div align="center">
	<p>&nbsp;</p>
	<p><img src="admin/images/spacer.gif" width="500" height="1"><br>
    <font color="#0066CC"><br>
    <span class="text"><b>Thank you for Placing Order<br>
    We will contact you as soon as possible. </b></span></font><br>
    <img src="admin/images/spacer.gif" width="500" height="1"> 
    <br>
    <span class="text">|</span> <a href ="computer.asp" class="links1">HOME</a> <span class="text">|</span></p>
</div>

	<%
end if 'end of session check	

sub sendTOclient()

	SQLtxt3="select * from Avorders where orderid='"&mlast&"'" 
	rsOrder.CursorLocation = adUseClient
	rsOrder.Open SQLtxt3,conn,adOpenForwardOnly,adLockReadOnly,adcmdtext
	set rsOrder.ActiveConnection = nothing
 
 
    conn.Execute("delete from Abasket where basketid="&trim(Request.QueryString("cartida"))&"")
    conn.Execute("delete from Abasketitems where basketid="&trim(Request.QueryString("cartida"))&"")
    
    'Response.Write "basket is empty now<br>" & Request.QueryString("cartid")


strHTML = strHTML & "<body style=font-family:Arial;font-size:12px>"
strHTML = strHTML & "<p>You placed an order request with us. Details are as under.</p>"
strHTML = strHTML & "<table width=675 border=0 cellpadding=0>"
strHTML = strHTML & " <tr bgcolor=#999999>" 




strHTML = strHTML & "<tr><td width=200><font face=Arial size=2><b>Name and Address:</font></td>"
strHTML = strHTML & "<td width=275><font face=Arial size=2><b>Shipping Address:</font></td>"
strHTML = strHTML & "<td width=200><font face=Arial size=2><b>Order Detail:</font></td>"
strHTML = strHTML & "</tr>"

strHTML = strHTML & "<tr><td width=200><font face=Arial size=2><b>"&company&"</b></font>"
strHTML = strHTML & "<td width=275><font face=Arial size=2><b>"&session("scompany")&"</b></font>"
' if Request.Cookies("isMember") <> "" then
'if customerid <> "" then
'strHTML = strHTML & "<td width=200><font face=Arial size=2><b>Customer</b></font></td>"
' ID :</b>"&customerid&"
'end if
if Request.Cookies("isDealer") <> "" then
'if dealerid <> "" then
strHTML = strHTML & "<td width=200><font face=Arial size=2><b>User</b></font></td>"
' ID :</b>"&dealerID&"
end if
strHTML = strHTML & "</tr>"

strHTML = strHTML & "<td width=200><font face=Arial size=1>"&fname&"&nbsp;"&lname&"</font></td>"
strHTML = strHTML & "<td width=275><font face=Arial size=1>"&session("scontact")&"</font></td>"
strHTML = strHTML & "<td width=200><font face=Arial size=1><b>Order No :</b>"&rsOrder("orderID")&"</font></td>"
strHTML = strHTML & "</tr><tr>"
strHTML = strHTML & "<td width=200><font face=Arial size=1>"&add1&"</font></td>"
strHTML = strHTML & "<td width=275><font face=Arial size=1>"&session("sadd1")&"</font></td>"
strHTML = strHTML & "<td width=200><font face=Arial size=1><b>Order Date :</b>"&rsOrder("created")&"</font></td>"
strHTML = strHTML & "</tr><tr>"
strHTML = strHTML & "<td width=200><font face=Arial size=1>"&add2&"</font></td>"
strHTML = strHTML & "<td width=275><font face=Arial size=1>"&session("sadd2")&"</font></td>"
strHTML = strHTML & "<td width=200>&nbsp;</td>"
strHTML = strHTML & "</tr><tr>"
strHTML = strHTML & "<td width=200><font face=Arial size=1>"&city&"-"&zip&","&country&"</font>"
strHTML = strHTML & "<td width=275><font face=Arial size=1>"&session("scity")&"-"&session("szip")&","&session("scountry")&"</font>"
strHTML = strHTML & "<td>&nbsp;</td>"
strHTML = strHTML & "</tr><tr>"
strHTML = strHTML & "</tr>"
strHTML = strHTML & "<td colspan=4 bgcolor=#999999></td>"
strHTML = strHTML & "</tr></table><br>"

strHTML = strHTML & " <table width=675 border=0 cellpadding=1>"
strHTML = strHTML & " <tr bgcolor=#999999>" 
strHTML = strHTML & "<td width=75><font face=Arial size=2><b>Product Id</b></font></td>"
strHTML = strHTML & "<td width=300><font face=Arial size=2><b>Item/Product</b></font></td>"
strHTML = strHTML & "<td width=100 align=center><font face=Arial size=2><b>Quanity</b></font></td>"
strHTML = strHTML & "<td width=100 align=right><font face=Arial size=2><b>Price</b></font></td>"
strHTML = strHTML & "<td width=100 align=right><font face=Arial size=2><b>Total</b></font></td>"
strHTML = strHTML & "</tr>"
do WHILE not RSORDER.EOF
strHTML = strHTML &"<tr>"
strHTML = strHTML & "<td width=75><font face=Arial size=1>"&RSORDER("itemcode")&"</font></td><td width=300><font face=Arial size=1>"&RSORDER("Description")&"</font></td><td width=100 align=center><font face=Arial size=1>"&RSORDER("QUANTITY")&"<font face=Arial size=1></td>"
strHTML = strHTML & "<td width=100 align=right><font face=Arial size=1>"& formatnumber(RSORDER("PRICEeACH"),2) &"</font></td><td width=100 align=right><font face=Arial size=1>"& formatnumber(RSORDER("LINETOTAL"),2) &"</font></td>"
RSORDER.MoveNext
strHTML = strHTML & "</tr>"
loop
RSORDER.MoveFirst 
strHTML = strHTML & "<tr bgcolor=#999999>" 
strHTML = strHTML & "<td width=75>&nbsp;</td>"
strHTML = strHTML & "<td width=300>&nbsp;</td>"
strHTML = strHTML & "<td width=100>&nbsp;</td>"
strHTML = strHTML & "<td width=100 align=right><font face=Arial size=2><b>Total</b></font></td>"
strHTML = strHTML & "<td width=100 align=right><font face=Arial size=2><b>"& formatnumber(RSORDER("ptotal"),2) &"</b></font></td>"
strHTML = strHTML & "</tr><tr>" 
strHTML = strHTML & "<td width=75>&nbsp;</td>"
strHTML = strHTML & "<td width=300><font face=arial size=2>Courier:&nbsp;<b>"&RSORDER("courier")&"</b></font></td>"
strHTML = strHTML & "<td width=100>&nbsp;</td>"
strHTML = strHTML & "<td width=100 align=right><font face=arial size=1>Total&nbsp;Freight&nbsp;Charged:</font></td>"
strHTML = strHTML & "<td width=100 align=right><font face=arial size=2><b>"& formatnumber(RSORDER("freight"),2) &"</b></font></td>"
strHTML = strHTML & "</tr><tr>"
strHTML = strHTML & "<td width=75>&nbsp;</td>"
strHTML = strHTML & "<td width=300><font face=arial size=2>Weight:&nbsp;<b>"&RSORDER("totalweight")&"</b>&nbsp;Kg.</font></td>"
strHTML = strHTML & "<td width=100>&nbsp;</td>"
strHTML = strHTML & "<td width=100>&nbsp;</td>"
strHTML = strHTML & "<td width=100>&nbsp;</td>"
strHTML = strHTML & "</tr><tr>"
strHTML = strHTML & "<td width=75>&nbsp;</td>"
strHTML = strHTML & "<td width=300><font face=arial size=2>Destination:&nbsp;<b>"&RSORDER("destination")&"</b></font></td>"
strHTML = strHTML & "<td width=100>&nbsp;</td>"
strHTML = strHTML & "<td width=100 align=right><font face=arial size=1>Sub-Total</font></td>"
strHTML = strHTML & "<td width=100 align=right><font face=arial size=2><b>"& formatnumber(rsorder("subTotal"),2) &"</b></font></td>"
strHTML = strHTML & "</tr><tr>"
if trim(rsorder("taxrate"))>0 then
strHTML = strHTML & "</tr><tr bgcolor=#999999>"
strHTML = strHTML & "<td colspan=4 align=right><font face=arial size=1><b>General Sales Tax(GST) @:"&rsorder("taxRate")&"%</b></font></td><td width=100 align=right><font face=arial size=2><b>"& formatnumber(rsorder("taxcharged"),2) &"</b></font></td>"
end if
if trim(session("htax"))>0 then
strHTML = strHTML & "</tr><tr bgcolor=#999999>"
strHTML = strHTML & "<td colspan=4 align=right><font face=arial size=1><b>With Holding Tax @:"&session("htax")&"%</b></font></td><td width=100 align=right><font face=arial size=2><b>"& formatnumber(session("chtax"),2) &"</b></font></td>"
end if
strHTML = strHTML & "</tr><tr bgcolor=#999999>"
strHTML = strHTML & "<td colspan=4 align=right><font face=arial size=2><b>Total Invoice Value</b></font></td><td width=100 align=right><font face=arial size=2><b>"& formatnumber(rsorder("Total"),2) &"</b></font></td>"
strHTML = strHTML & "</tr></table>"


objMessage.From =rsemail("ccemail")  'From bbl
objMessage.To = mEmail                 ' To client   
objMessage.Bcc = rsemail("ccemail") 
objMessage.Subject ="Order Information"
objMessage.HTMLBody = "<HTML><BODY>" & strHTML & "</BODY></HTML>"
'objMessage.Body = strHTML
objMessage.Send
Set objMessage = Nothing

end sub

'Response.Write "cartid=" & Request.QueryString("cartid") & "<br>"
'Response.Write conn.State 
%>
