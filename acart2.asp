<%@ Language=VBScript %>
<%
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
a=request("a")
xtoday=now
pid=request("xpid")
mqty=cdbl(request("mqty"))
'Response.Write xtoday
bb=60+xtoday
'Response.Write bb
mact=request("mAct")
mcd=request("mcode")
msCat=request("mdata")

Response.Write "<br>msCat=" & request("msCat") & "<BR>"

set conn = server.CreateObject ("ADODB.Connection")
set rsCart = server.CreateObject("ADODB.Recordset")
set RScartid = server.CreateObject("ADODB.Recordset")
set RScartitems = server.CreateObject("ADODB.Recordset")
set rsQtys = server.CreateObject("ADODB.Recordset")

%>
<!--#INCLUDE FILE="dbconnect.asp"-->
<%
conn.Open con 
bid=request("bid")
'mcd=request("mcode")
brand=request("brand")
if bid <> "" then
Response.Redirect "acart.asp?cids="&bid+"-"+mcd+"-"+mscat
else

if  session("msta")="" then
	session("msta")=1	
	
	sqltxt="insert into Abasket values('"&bb&"','"&xtoday&"')"
	set rsCartID=conn.execute(sqltxt)

	mysql="select * from Abasket order by basketid desc"
	set rscartid=conn.execute(mysql)
	session("xscartIDa")=rscartid.Fields("basketid")
	scartid=session("xscartIDa")
	'Response.Write scartid	
end if

if session("msta")<> "" then 
SQLtxt="select * from AvPQtyUnit where productid='"&pid&"'" 

'Response.Write sqltxt

rsQtys.CursorLocation = adUseClient
rsQtys.Open SQLtxt,conn,adOpenForwardOnly,adLockReadOnly,adcmdtext
set rsQtys.ActiveConnection = nothing


'Response.Write "weight = " & rsQtys("weight") & "<br>"



if not rsqtys.EOF then

'if Request.Cookies("isMember")= "" then
'mprice=rsqtys("retail")
'end if

'Response.Write "<br> shipping c = " & rsQtys("shippingc") & "<br>"
'Response.Write "<br> taxable c = " & rsQtys("taxable") & "<br>"



if Request.Cookies("isDealer") <> "" AND Request.Cookies("TypeName") ="Reseller"  OR Request.Cookies("TypeName") ="Cable Net" then               
Response.Write "cookie found" & "<br>"
pa=rsqtys("qtyA")
pb=rsqtys("qtyB")  
pc=rsqtys("qtyC")
pd=rsqtys("qtyD")

Response.Write "pa=" & pa & "<br>"

mrange1=cint(mid(pa,instr(1,pa,"-")+1,8)) '5
mrange2=cint(mid(pb,instr(1,pb,"-")+1,8)) '10
mrange3=cint(mid(pc,instr(1,pc,"-")+1,8)) '15
mrange4=cdbl(mid(pd,instr(1,pd,"-")+1,8)) '20

'marange1=1
'marange1=2
'marange1=3
'marange1=4
'marange1=5


mprice=""
if mqty <= mrange1 then
  mprice=trim(rsqtys("pgA")) 
  Response.Write "pga"
  Response.Write mprice
end if

if mqty > mrange1 and mqty <= mrange2 then

  mprice=trim(rsqtys("pgb"))
  Response.Write "pgb"
  Response.Write mprice
end if


if mqty > mrange2 and mqty <= mrange3 then
  mprice=trim(rsqtys("pgc"))
  Response.Write "pgc"
  Response.Write mprice
end if

if mqty > mrange3 and mqty <= mrange4 then
  mprice=trim(rsqtys("pgd"))
  Response.Write "pgd"
  Response.Write mprice
end if

if mqty > mrange4  then
  mprice=trim(rsqtys("pgd"))
  Response.Write "pgd"
  Response.Write mprice
end if


else if Request.Cookies("isDealer") <> "" then
Response.Cookies("isloggedinas").expires=#december, 31 1999 00:00:00#
Response.Write "NO Cookie"

mprice=rsqtys("retail")
Response.Write trim(mprice) * trim(session("price__"))
Response.Write "<br>mprice=" & trim(mprice) * trim(session("price__")) & "<br>"


end if 'end of cookie check
end if


end if
	psqltxt="insert into AbasketItems(basketid,productID,quantity ,priceEach,taxable,freeshipping,weight) values('"&session("xscartIDa")&"','"&pid&"','"&mqty&"', [dbo].Editprice('"&trim(mprice) * trim(session("price__"))&"'),'"&rsQtys("taxable")&"','"&rsQtys("shippingc")&"',"&rsQtys("weight")&")"
    Response.Write psqltxt
    
    
    set RScartitems=conn.execute(psqltxt)
end if

scid=session("xscartIDa")
	application("mid")=session("xscartIDa")
	aaa=application("mid")


	SQLtxt3="SELECT SUM(Quantity * priceEach)AS tota, COUNT(*) AS nitemsa FROM AvCart  where basketid='"&scid&"'"
	RScartitems.CursorLocation = adUseClient
	RScartitems.Open SQLtxt3,conn,adOpenForwardOnly,adLockReadOnly,adcmdtext
	
	
	'if Request.Cookies("isDealer") <> "" AND Request.Cookies("TypeName") ="Reseller"  OR Request.Cookies("TypeName") ="Cable Net" then               
	arshi=mcd&"-"&mscat&"-"&RScartitems("tota")&"-"&RScartitems("nitemsa")&"-"&aaa
	'Response.Write "here in dealer"
    'else if Request.Cookies("isDealer") <> "" then
    'arshi=mcd&"-"&mscat&"-"&RScartitems("tota")&"-"&RScartitems("nitemsa")&"-"&aaa
    'Response.Write "here in member"
    'end if
    'end if
    
   
    
  Response.redirect "apage2.asp?mcd="&arshi
end if

%>