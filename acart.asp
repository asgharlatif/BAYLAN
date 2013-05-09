<%@ Language=VBScript %>
<%
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
a=request("a")
xtoday=now
pid=request("xpid")
mqty=request("mqty")
'Response.Write xtoday
bb=60+xtoday
'Response.Write bb
zcode=request("cids")
if zcode <> "" then
mrange=split(zcode,"-")
for i =0 to ubound(mrange)
	cartida=mrange(0)   
	mcd=mrange(1)  'mcode to view same selected catgory as before
	mscat=mrange(2)  'subcat id to view same selected catgory as before

next
else
cartida=request("cartida")
mcd=request("mcd")


'mscat=

end if
'cartid=request("cartID")
'mcd=request("mcode")
'Response.Write mact

set conn = server.CreateObject ("ADODB.Connection")
set rsCart = server.CreateObject("ADODB.Recordset")
set RScartid = server.CreateObject("ADODB.Recordset")
set RScartitems = server.CreateObject("ADODB.Recordset")

%>
<!--#INCLUDE FILE="dbconnect.asp"-->
<%
conn.Open con 

if request("delete") <> ""  then 'Delete Item
for each vch in request("chkdel")
'Response.Write "me clicik tod delete" & vch & "<br>"
msql="delete from  AbasketItems where basketItemID="&trim(vch)&""
conn.Execute(msql)
next
end if



	SQLtxt="select * from Avcart where basketid='"&cartida&"'" 
	rsCart.CursorLocation = adUseClient
	rsCart.Open SQLtxt,conn,adOpenForwardOnly,adLockReadOnly,adcmdtext
	set rsCart.ActiveConnection = nothing
'	Response.Write sqltxt
'Response.Write rscart("item")
if not rscart.EOF then
rscart.MoveFirst
mbID=rscart("basketItemID")
trec=rscart.RecordCount
end if



if request("recalc") <> ""  then 'RECALCULATE QTY
mbid=request("mbid")
trec=request("trec")

if Request.Cookies("isDealer") <> "" AND Request.Cookies("TypeName") ="Reseller"  OR Request.Cookies("TypeName") ="Cable Net" then               
Response.Write "<div align=center> <b><font face=arial size=2 color=red>You are a Dealer you can not add in Qty here please Delete item and then add with desired qty.</font></b></div>"
 else if Request.Cookies("isDealer") <> "" then
dim qt(30)
tqt=0
xbid=mbid
for j=1 to trec
qt(j)=request("new_quantity"&j)
if qt(j)=0 then
  Response.Write "<div align=center> <b><font face=arial size=2 color=red> >>>>  Qty can not be a ZERO plus replace ZERO with valid QTY.</font></b></div>"
 else 
	msql="update AbasketItems set quantity="&qt(j)&" where basketItemID="&xbid&""
	conn.Execute(msql)
	'Response.Write msql
	xbid=mbid+j
end if	
next
Set rsCart.ActiveConnection = CONN
RSCART.Requery
end if ' end of cookie check
end if 
end if


if request("more") <> "" then  'More Shopping pressed
	SQLtxt3="SELECT SUM(Quantity * priceEach)AS tota, COUNT(*) AS nitemsa FROM AvCart  where basketid='"&cartida&"'"
	RScartitems.CursorLocation = adUseClient
	RScartitems.Open SQLtxt3,conn,adOpenForwardOnly,adLockReadOnly,adcmdtext
	arshi=mcd&"-"&mscat&"-"&RScartitems("tota")&"-"&RScartitems("nitemsa")&"-"&cartida
  ' Response.Write  RScartitems("tot")&"-"&RScartitems("nitems")
  Response.redirect "apage2.asp?mcd="&arshi
end if


%>
<HTML><HEAD><TITLE>Shopping Cart</TITLE>
<META content="text/html; charset=iso-8859-1" http-equiv=Content-Type>
<link rel="stylesheet" href="../admin/arshilogin.css" type="text/css">

<script language="javascript">
function RemoveCheck() {
var  i,j;

j=0;
for ( var i=0; i<document.myform.elements.length;i++)
{
	
 if (document.myform.elements[i].type == "checkbox" ){
		if (document.myform.elements[i].checked == true )
		{
		j=1;
		}
		}
}



if (j==1){
alert ('Page is going to checkout not delete')
}

if (j != 1) {
window.location.href="acheckout.asp?bid="+document.myform.cartidt.value;
}

}
</script>

</HEAD>
<BODY aLink=#990000 bgColor=#ffffff leftMargin=0 link=#990000 
text=#000000 topMargin=0 vLink=#999999 marginheight="0" 
marginwidth="0">

<TABLE border=0 cellPadding=0 cellSpacing=0 width=780>
  <TBODY> 
  <TR> 
    <TD align=left colSpan=5 height=228 vAlign=top> 
      <TABLE border=0 cellPadding=0 cellSpacing=0 width=650>
        <TBODY> 
        <TR> 
          <TD height=290 vAlign=top><BR>
            <b><font face="Verdana, Arial" size="5">Shopping Cart</font></b><BR>
            <BR>
            <FORM action="acart.asp?mact=rec&mbid=<%=mbid%>&cartida=<%=cartida%>&trec=<%=trec%>&mcd=<%=mcd%>" method="post" name="myform">
              <input type="hidden" name="cartidt" value =<%=cartida%>>
              <TABLE align=center bgColor=#cccccc border=0 cellPadding=1 
            cellSpacing=1 width=550>
                <INPUT name=pcode type=hidden>
                <INPUT 
              name=q type=hidden>
                <TBODY></TBODY> 
              </TABLE>
              <FONT 
            face="Verdana, Arial, Helvetica, sans-serif" size=2>Here is the detail 
              of the products you have added in the shopping cart so far. If you 
              wish to make any change in the shopping cart or selected items you 
              may do so by just using the various options below.<BR>
              </FONT><BR>
              <TABLE align=center bgColor=#cccccc border=0 cellPadding=1 
            cellSpacing=1 width=597>
                <TBODY> 
                <TR bgColor=#ffffff> 
                  <TD width=10% bgcolor="#0099FF"> 
                    <DIV align=center><FONT 
                  face="Verdana, Arial, Helvetica, sans-serif"><B><FONT 
                  color=#000000 size=2>Quantity</FONT></B></FONT></DIV>
                  </TD>
                  <TD width=33% bgcolor="#0099FF"> 
                    <DIV align=center><FONT 
                  face="Verdana, Arial, Helvetica, sans-serif"><B><FONT 
                  color=#000000 size=2><B>Item / Product</B></FONT></B></FONT></DIV>
                  </TD>
                  <TD width=15% bgcolor="#0099FF" colspan="2"> 
                    <DIV align=center><FONT 
                  face="Verdana, Arial, Helvetica, sans-serif"><B><FONT 
                  color=#000000 size=2>Price/Each</FONT></B></FONT></DIV>
                  </TD>
                  <TD width=15% bgcolor="#0099FF" colspan="2"> 
                    <DIV align=center><FONT 
                  face="Verdana, Arial, Helvetica, sans-serif"><B><FONT 
                  color=#000000 size=2>Total Price</FONT></B></FONT></DIV>
                  </TD>
                  <TD width=15% bgcolor="#0099FF" align="center"><font 
                  face="Verdana, Arial, Helvetica, sans-serif"><b><font 
                  color=#000000 size=2>Weight</font></b></font></TD>
                  <TD width=12% bgcolor="#0099FF"> 
                    <DIV align=center><FONT 
                  face="Verdana, Arial, Helvetica, sans-serif"><B><FONT 
                  color=#000000 size=2>Check to remove </FONT></B></FONT></DIV>
                  </TD>
                </TR>
                <% lp=1
               tqt=0
               t_tot=0
               wt1=0
               do while (not rscart.EOF)
               qt1=rscart("quantity")
               tt=rscart("priceEach")
               
               wt=cdbl(rscart("weight"))*trim(qt1)
               tot1=trim(tt)*trim(qt1)
               '--------------------
               'Response.Write "weight=" & cdbl(rscart("weight")) * trim(qt1)
               
               't_tot=100
               
               '---------------------
               
               t_tot=t_tot+tot1
              
               'Response.Write t_tot%>
                <TR bgColor=#ffffff> 
                  <TD width=10%> 
                    <DIV align=center><FONT face="Verdana, Arial, Helvetica, sans-serif"> 
                      <INPUT maxLength=8 name="new_quantity<%=lp%>" size=8 value=<%=rscart("quantity")%> STYLE="text-align:right">
                      </FONT></DIV>
                  </TD>
                  <TD width=33% class="text"> <%=rscart("Description")%> </TD>
                  <INPUT name=amount type=hidden value=<%=rscart("priceEach")%>>
                  <TD width=25> 
                    <div align="left"><FONT color=#000000 
                  face="Verdana, Arial, Helvetica, sans-serif" size=2>Rs. </FONT></div>
                  </TD>
                  <TD width=75>
                    <div align="right"><font color=#000000 
                  face="Verdana, Arial, Helvetica, sans-serif" size=2><%=formatnumber(rscart("priceEach"),2)%></font></div>
                  </TD>
                  <TD width=25> 
                    <div align="left"><FONT color=#000000 
                  face="Verdana, Arial, Helvetica, sans-serif" size=2>Rs. </FONT></div>
                  </TD>
                  <TD width=75>
                    <div align="right"><font color=#000000 
                  face="Verdana, Arial, Helvetica, sans-serif" size=2><%=formatnumber(trim(rscart("priceEach")) * trim(rscart("quantity")),2) %></font></div>
                  </TD>
                  <TD width=15%> 
                    <div align="right"><font color=#000000 
                  face="Verdana, Arial, Helvetica, sans-serif" size=2><%=formatnumber(wt,2)%></font></div>
                  </TD>
                  <TD width=12%> 
                    <div align="center">
                      <INPUT name="chkdel" type=checkbox value=<%=rscart("basketItemID")%>>
                    </div>
                  </TD>
                  <%rscart.MoveNext
                  lp=lp+1
                  tqt=cdbl(tqt)+cdbl(qt1)
                  wt1=wt1+wt
                '  Response.Write tqt
                  loop
                  if wt1 < 1 then
					  wt1=1
				  end if  
				  %>
                </TR>
                <TR bgColor=#ffffff> 
                  <TD width=63 bgcolor="#0099FF"> 
                    <DIV align=center><FONT color=#000000 
                  face="Verdana, Arial, Helvetica, sans-serif"> 
                      <INPUT maxLength=8 name=t_items  readOnly size=8 STYLE="text-align:right" value=<%=tqt%>>
                      </FONT></DIV>
                  </TD>
                  <TD width=33% bgcolor="#0099FF"> 
                    <DIV align=center><FONT color=#000000 
                  face="Verdana, Arial, Helvetica, sans-serif" size=2><B>Total 
                      Amount</B></FONT></DIV>
                  </TD>
                  <TD width=15% bgcolor="#0099FF" colspan="2"> 
                    <DIV align=right><font color=#000000 
                  face="Verdana, Arial, Helvetica, sans-serif" 
                  size=2><b>Rs.</b></font><FONT color=#000000 
                  face="Verdana, Arial, Helvetica, sans-serif"> </FONT></DIV>
                  </TD>
                  <TD width=15% bgcolor="#0099FF" class="text" colspan="2">
                    <div align="right"><font color=#000000 
                  face="Verdana, Arial, Helvetica, sans-serif"> 
                      <input type="text" name="t_amount" readonly="readOnly" size="13" value="<%=formatnumber(t_tot,2)%>" STYLE="text-align:right">
                      </font> </div>
                  </TD>
                  <TD width=15% bgcolor="#0099FF"> 
                    <div align="right"><b><font size="2">Kg</font></b><font size="2">.</font><FONT color=#000000 
                  face="Verdana, Arial, Helvetica, sans-serif"> 
                      <input maxlength=15 name=twt  align=right readOnly size=6 value=<%=formatnumber(wt1,2)%> STYLE="text-align:right">
                      </FONT></div>
                  </TD>
                  <TD width=12% bgcolor="#0099FF"> 
                    <DIV align=center><FONT 
                  face="Verdana, Arial, Helvetica, sans-serif"><B></B></FONT></DIV>
                  </TD>
                </TR>
                </TBODY> 
              </TABLE>
              <BR>
              <TABLE align=center border=0 cellPadding=0 cellSpacing=0 
              width=550>
                <TBODY> 
                <TR> 
                  <TD height=18 width=86> 
                    <DIV align=center> 
                      <input type="submit" name="recalc" value="Recalculate" >
                    </DIV>
                  </TD>
                  <TD width=204> 
                    <DIV align=center> 
                      <input type="submit" name="more" value="More Shopping" >
                    </DIV>
                  </TD>
                  <TD width=81> 
                    <DIV align=center> 
                      <input type="submit"  name="delete" value="Delete Selected">
                    </DIV>
                  </TD>
                  <TD height=18 width=179> 
                    <DIV align=center> 
                      <input type="button"  name="checkout" value="Checkout"   onclick="RemoveCheck()">
                    </DIV>
                  </TD>
                </TR>
                </TBODY> 
              </TABLE>
            </FORM>
            <P align=center>&nbsp;</P>
            <P align=center>&nbsp;</P>
          </TD>
        </TR>
        <TR> 
          <TD vAlign=top> 
            <MAP name=Map2Map> 
              <AREA coords=-5,3,62,72 
              href="computer.asp" 
        shape=RECT>
            </MAP>
          </TD>
        </TR>
        </TBODY> 
      </TABLE>
    </TD>
  </TR>
  <TR> 
    
  
  </TR>
  </TBODY>
</TABLE>
