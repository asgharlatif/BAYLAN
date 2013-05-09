<%@ Language=VBScript %>
<!--#INCLUDE FILE="dbconnect.asp"-->
<%
'Response.Write "<br>mcd=" & Request.Cookies("gstnumber") & "<br>"

set conn = server.CreateObject ("ADODB.Connection")
conn.Open con 
if session("msessiona")= "" then
	Response.Write "<font face =arial size=2 color=red><b>Your Session has expired!</b></font>"
else
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3

'Response.Write "----------------------------" & Request.form("sameAddress") & "<br>"
'Response.Write "Dealers----------------------------" & Request.Cookies("isDealer") & "<br>"
'Response.Write "Mem----------------------------" & Request.Cookies("isMember") & "<br>"

customerID=request("CustomerID") 'from checkout.asp
DealerID=request("dealerID") 'from checkout.asp
shipaddID=request("shipID") 'from checkout.asp
CustomerType=request("Ctype") 'from checkout.asp

'Response.Write "bid=" & dealerID & "<br>" & Request.Cookies("dealerid")


if customerID <> "" then
	session("customerID")=customerID
end if

if dealerID <> "" then
	session("dealerID")=DealerID
end if

if request("shipID") <> "" then

if ucase(Request.form("sameAddress"))= ucase("sameAddress") then
	session("shipaddid")=request("shipID")
	session("addressmatch")=0
	%>
<script language="javascript">
//alert ('address is same')
</script>
<%
	
end if

if ucase(Request.form("sameAddress")) <>  ucase("sameAddress") then
session("addressmatch")=1
session("shipaddid")=""
%>
<script language="javascript">
//alert ('select new address')
</script>
<%
if Request.Cookies("isDealer")<> "" then
if request("semail") <> "" then
session("semail")=trim(request("semail"))
session("sadd1")=trim(request("sadd1"))
session("sadd2")=trim(request("sadd2"))
session("scity")=trim(request("scity"))
session("sph")=trim(request("sph"))
session("sfax")=trim(request("sfax"))
session("szip")=trim(request("szip"))
session("scountry")=trim(request("scountry"))
session("scontact")=trim(request("scontact"))
session("scompany")=trim(request("scompany"))
end if
elseif Request.Cookies("isMember") <> "" then

				if request("semail") <> "" then
				session("semail")=trim(request("semail"))
				session("sadd1")=trim(request("sadd1"))
				session("sadd2")=trim(request("sadd2"))
session("scity")=trim(request("scity"))
session("sph")=trim(request("sph"))
session("sfax")=trim(request("sfax"))
session("szip")=trim(request("szip"))
session("scountry")=trim(request("scountry"))
session("scontact")=trim(request("scontact"))
session("scompany")=trim(request("scompany"))
end if


end if
end if	
	
end if


'Response.Write "session v is =" & session("addressmatch") & "-" & session("scompany")


if shipaddid <> "" then
	session("customerType")=customerType
end if

xfreight=request("xFreight")
a=request("a")
xtoday=now()
pid=request("xpid")
zcode=request("cids")
if zcode <> "" then
	mrange=split(zcode,"-")
	for i =0 to ubound(mrange)
		cartida=mrange(0)   
		mcd=mrange(1)  'mcode to view same selected catgory as before
		mscat=mrange(2)  'subcat id to view same selected catgory as before
	next
else
	
	if xFreight <> "" then 'input from Freight dropmenu
		mFr=split(xFreight,"-")
		for j =0 to ubound(mFr)
			mFreight=mfr(0)  'Freight/kg
			cartida=mfr(1)   
			mcour=mfr(2)  'courier name
			mdest=mfr(3)  'Destination
		next
		session("mcour")=mcour
		session("mdest")=mdest
	else
		cartida=request("cartida")
		mcd=request("mcd")
	if cartida<> "" then
	  session("cartida")=cartida
     end if
	end if
end if

set rsCart = server.CreateObject("ADODB.Recordset")
set rsFreight= server.CreateObject("ADODB.Recordset")
set rsgst= server.CreateObject("ADODB.Recordset")
set rsMAX= server.CreateObject("ADODB.Recordset")

if trim(Request("bids")) <> "" then
session("bids")=Request("bids")
end if


	'Response.Write "masghar"
	
	SQLtxt="select * from Avcart where basketid="&session("bids")&"" 
	rsCart.CursorLocation = adUseClient
	
	rsCart.Open SQLtxt,conn,adOpenForwardOnly,adLockReadOnly,adcmdtext
	set rsCart.ActiveConnection = nothing

	
	SQLtxt2="select * from AvFreights order by Courier" 
	rsFreight.CursorLocation = adUseClient
	rsFreight.Open SQLtxt2,conn,adOpenForwardOnly,adLockReadOnly,adcmdtext
	set rsFreight.ActiveConnection = nothing





	SQLtxt3="select gstcharge from Agst where gst='"&Request.Cookies("gstcat")&"'" 
	rsgst.CursorLocation = adUseClient
	rsgst.Open SQLtxt3,conn,adOpenForwardOnly,adLockReadOnly,adcmdtext
	set rsgst.ActiveConnection = nothing
	
	set rshtax=server.CreateObject("adodb.recordset")
	SQLtxt3="select htaxcharge from Ahtax where htax='"&Request.Cookies("htaxcat")&"'" 
	rshtax.CursorLocation = adUseClient
	rshtax.Open SQLtxt3,conn,adOpenForwardOnly,adLockReadOnly,adcmdtext
	set rshtax.ActiveConnection = nothing
	
	
	set Rfreights =server.CreateObject("adodb.recordset")
	Rfreights.CursorLocation=aduseclient
	Rfreights.Open "select * from ACouriers",conn,adOpenForwardOnly,adLockReadOnly,adcmdtext
	set Rfreights.ActiveConnection=nothing
	
	
	
	
	
	
	
	
	

if not rsgst.EOF then
if trim(Request.Cookies("isDealer")) <> "" or trim(Request.Cookies("isMember")) <> "" then
'if trim(Request.Cookies("gstnumber")) = "----" or trim(Request.Cookies("gstnumber"))="" then
'mgst=rsgst where ("gstcharge") '"& Request.Cookies("gstcat") &"'" 
'response.write "gstcat " & rsgst("gstcharge") &" "
mgst=rsgst("gstcharge")
end if


'musRate=rsExtras("usrate")


end if

if not rscart.EOF then
rscart.MoveFirst
mbID=rscart("basketItemID")
trec=rscart.RecordCount
end if





%>
<HTML><HEAD><TITLE>Checkout Step-2</TITLE>
<link rel="stylesheet" href="../admin/arshilogin.css" type="text/css">
<script language="javascript">
function getunits(form)
{



aa=form.MainCat.options.value;

alert (aa)

window.location.href="acartCheckout.asp?xFreight="+aa;
}


function findunit()
{
alert ('aw')
window.location.href="acartCheckout.asp"
}




function movetopage()
{
window.location="computer.asp"
}


 function myform_Validator(theForm)
                     {
                     var alertsay = ""; 
                     
                     if (theForm.MainCat.value == "Select Courier")
                     {
                     alert("You must Select a Courier");
                     theForm.MainCat.focus();
                     return (false);
                      }
                     
                      if (theForm.Destination.value == "Select Destination")
                     {
                     alert("You must Select a Destination");
                     theForm.Destination.focus();
                     return (false);
                      } 
                      
					  if (theForm.t_amount.value == "0")
                     {
                     alert("You must Purchase Some Item then Checkout");
                     theForm.t_amount.focus();
                     return (false);
                      }
					  
                     return (true);
                     }



</script>

</HEAD>
<BODY aLink=#990000 bgColor=#ffffff leftMargin=0 link=#990000 
text=#000000 topMargin=0 vLink=#999999 marginheight="0" 
marginwidth="0">

<TABLE border=0 cellPadding=0 cellSpacing=0 width=780>
  <TBODY> 
  <TR> 
    <TD height=228 
    vAlign=top width="61">&nbsp;</TD>
    <TD align=middle colSpan=5 height=228 vAlign=top> 
      <TABLE border=0 cellPadding=0 cellSpacing=0 width=708>
        <TBODY> 
        <TR> 
          <TD height=290 vAlign=top> <br>
            <%
if request("ChargeB") <> "" then
%>
            <script language="javascript">
alert ("yahhhhhhhhhh")
</script>
            <%
end if

%>
            <FORM action="acartCheckout.asp?cartida=<%=session("cartida")%>&custID=<%=session("CustomerID")%>&DealID=<%=session("dealerID")%>" method=post onSubmit="return myform_Validator(this)" name="myform">
              <table width="100%" border="0" cellspacing="0" cellpadding="0" height="42">
                <tr> 
                  <td bgcolor="#0099FF" class="text" height="31" align="center" valign="middle"> 
                    <h2><b><font color="#FFFFFF">Checkout Step-3</font></b></h2>
                  </td>
                  <td bgcolor="#0099FF" align="center" valign="middle" height="31"> 
                    <input type="button" name="moreb" value="More Shopping" onclick="movetopage()">
                  </td>
                </tr>
              </table>
              <TABLE align=center bgColor=#cccccc border=0 cellPadding=1 
            cellSpacing=1 width=700>
                <INPUT name=pcode type=hidden>
                <INPUT 
              name=q type=hidden>
                <TBODY></TBODY> 
              </TABLE>
              <FONT 
            face="Verdana, Arial, Helvetica, sans-serif" size=2>Here is the detail 
              of the products you have added in the shopping cart so far. Its 
              Final stage of Checkout processa.<BR>
              </FONT><BR>
              <TABLE align=center bgColor=#cccccc border=0 cellPadding=1 
            cellSpacing=1 width=691>
                <TBODY> 
                <TR> 
                  <TD width=89 bgcolor="#0099FF"> 
                    <DIV align=center><FONT 
                  face="Verdana, Arial, Helvetica, sans-serif"><B><FONT 
                  color=#000000 size=2>Quantity</FONT></B></FONT></DIV>
                  </TD>
                  <TD width=349 bgcolor="#0099FF" colspan="2"> 
                    <DIV align=center><FONT 
                  face="Verdana, Arial, Helvetica, sans-serif"><B><FONT 
                  color=#000000 size=2><B>Item / Product</B></FONT></B></FONT></DIV>
                  </TD>
                  <TD width=127 bgcolor="#0099FF"> 
                    <DIV align=center><FONT 
                  face="Verdana, Arial, Helvetica, sans-serif"><B><FONT 
                  color=#000000 size=2>Price</FONT></B></FONT></DIV>
                  </TD>
                  <TD width=113 bgcolor="#0099FF" align="center"><font 
                  face="Verdana, Arial, Helvetica, sans-serif"><b><font 
                  color=#000000 size=2>Weight/Kgs.</font></b></font></TD>
                  <TD width=113 bgcolor="#0099FF" align="center"><font 
                  face="Verdana, Arial, Helvetica, sans-serif"><b><font 
                  color=#000000 size=2>Total Price.</font></b></font></TD>
                </TR>
                <% lp=1
               tqt=0
               t_tot=0
               wt1=0
       totaltax=0
       TFS_Weight=0
       do while (not rscart.EOF)
               
               qt1=rscart("quantity")
               tt=rscart("priceEach")
               wt=cdbl(rscart("weight"))*cdbl(qt1)
               tot1=cdbl(tt)*cdbl(qt1)
               '---------------------
               
               if rscart("tax") <> "No" then
               totaltax=totaltax + ( tot1 * trim (mgst) / 100)
               end if
               
               '---------------
               
               
               t_tot=t_tot + tot1
               if rscart("tax") <> "No" then
               
               xGst0= tot1*(cdbl(mGST)/100)
               grandTotal0=xgst0+Totl
			   tax_tot=grandTotal0
			   else
			   t_tot3=t_tot+tot1
			   end if	
               '--------------------------Response.Write "<br>totaltax=" & tax_tot & "<br>"			
               %>
                <TR> 
                  <TD width=89 bgcolor="#ffffff"> 
                    <DIV align=center><FONT face="Verdana, Arial, Helvetica, sans-serif"> 
                      <INPUT maxLength=3 name="new_quantity<%=lp%>" size=3 readonly value=<%=rscart("quantity")%> STYLE="text-align:right">
                      </FONT></DIV>
                  </TD>
                  <TD width=349 class="text" bgcolor="#ffffff" colspan="2"> 
                    <%
                     if rscart("tax") <> "No" then
                     Response.Write rscart("Description")%>
                    <font color=blue> [Taxable]</font> 
                    <%else
                     Response.Write rscart("Description")
                     end if
                     
                     if rscart("ship") = "No" then
                     %>
                    <font color=red>[Free Shipping] </font> 
                    <%
                     end if
                     %>
                  </TD>
                  <INPUT name=amount type=hidden value=<%=rscart("priceEach")%>>
                  <TD width=127 bgcolor="#ffffff" align="right"> <FONT color=#000000 
                  face="Verdana, Arial, Helvetica, sans-serif" size=2>Rs. <%=formatnumber(rscart("priceEach"),2)%></FONT></TD>
                  <TD width=113 bgcolor="#ffffff" align="right"><font color=#000000 
                  face="Verdana, Arial, Helvetica, sans-serif" size=2><%=formatnumber(wt,2)%></font></TD>
                  <TD width=113 bgcolor="#ffffff" align="right"><font color=#000000 
                  face="Verdana, Arial, Helvetica, sans-serif" size=2><%=formatnumber(trim(rscart("priceEach")) * trim(rscart("quantity")),2)%></font></TD>
                  <%
                 
                  
                  if ucase(rscart("ship")) = ucase("yes") then
                  TFS_Weight=TFS_Weight + wt
                  end if
                  wt1=wt1+wt
                  rscart.MoveNext
                  lp=lp+1
                  tqt=cdbl(tqt)+cdbl(qt1)
                  loop
           
                  
     
     if TFS_Weight < 1 and TFS_Weight <> 0  then
     TFS_Weight=1
     elseif instr(1,TFS_Weight,".") then
     TFS_Weight =mid(TFS_Weight,1,instr(1,TFS_Weight,".")-1)+1 
     end if
    
                 
				
				'Response.Write nTotal  "<br>" 
                  %>
                </TR>
                <TR> 
                  <TD width=89 bgcolor="#0099FF"> 
                    <DIV align=center><FONT color=#000000 
                  face="Verdana, Arial, Helvetica, sans-serif"> 
                      <INPUT maxLength=4 name=t_items  readOnly size=3 value=<%=tqt%> STYLE="text-align:right">
                      </FONT></DIV>
                  </TD>
                  <TD width=349 bgcolor="#0099FF" align="right" colspan="2"> <FONT color=#000000 
                  face="Verdana, Arial, Helvetica, sans-serif" size=2><B>Total 
                    Amount</B></FONT></TD>
                  <TD width=127 bgcolor="#0099FF" align="right"> <FONT color=#000000 
                  face="Verdana, Arial, Helvetica, sans-serif" 
                  size=2>Rs.</FONT><FONT color=#000000 
                  face="Verdana, Arial, Helvetica, sans-serif"> 
                    <INPUT maxLength=15 name=t_amount  readOnly size=11 value=<%=formatnumber(t_tot,2)%> STYLE="text-align:right">
                    </FONT></TD>
                  <TD width=113 bgcolor="#0099FF" align="right"> <FONT color=#000000 
                  face="Verdana, Arial, Helvetica, sans-serif" size=1><B> Kgs. 
                    <INPUT maxLength=15 name=twt  readOnly size=6 value=<%=formatnumber(wt1,2)%> STYLE="text-align:right">
                    </B></FONT></TD>
                  <TD width=113 bgcolor="#0099FF" align="right"> <FONT color=#000000 
                  face="Verdana, Arial, Helvetica, sans-serif" size=1><B> &nbsp;</B></FONT></TD>
                  <%
                   if request("MainCat") <> ""  and ucase(request("MainCat")) <> ucase("Select Courier") then
                   mCour=request("MainCat")
                   end if 
                   if request("Destination") <> ""  and ucase(request("Destination")) <> ucase("Select Destination") then
                   mDest=request("Destination")
                   end if
                   
				    %>
                </TR>
                <TR> 
                  <TD width=89 bgcolor="#FFFFFF" height="2"></TD>
                  <TD width=349  bgcolor="#FFFFFF" align="right" colspan="2"></TD>
                  <TD width=127 bgcolor="#FFFFFF" height="2"></TD>
                  <TD width=113 bgcolor="#FFFFFF" height="2"></TD>
                </TR>
                <TR> 
                  <TD width=89 bgcolor="#FFFFFF" class="text">Freight Charges<br>
                    <br>
                    <br>
                  </TD>
                  <TD width=125 bgcolor="#FFFFFF"><span class="text">Courier :<br>
                    <br>
                    <br>
                    Destination:<br>
                    </span><br>
                    <br>
                    <br>
                    <span class="text"><font color="#FF0000"><b> </b></font></span></TD>
                  <TD width=224 bgcolor="#FFFFFF" class="text" align="right"> 
                    <p> 
                      <select name="MainCat"  onblur="(document.myform.Destination.value='Select Destination')" onchange="submit()">
                        <option value="Select Courier">Select Courier</option>
                        <% 
    set RCour =server.CreateObject("adodb.recordset")
	RCour.CursorLocation=aduseclient
    RCour.Open "select * from CourierWithFreight",conn,adOpenForwardOnly,adLockReadOnly,adcmdtext
	set RCour.ActiveConnection=nothing
	                  session("mcour")=""
	                  session("mdest")=""              
                      RCour.MoveFirst 
                      Do While Not RCour.EOF%>
                        <option value ="<%=RCour("courier")%>" ><%=RCour("courier")%></option></option> 
                        <%RCour.MoveNext 
					  loop
					  
					  if request("MainCat") <> ""  and ucase(request("MainCat")) <> ucase("Select Courier") then
					  %>
                        <option value ="<%= request("MainCat")%>" selected><%=request("MainCat")%></option>
                        <%  
					  RCour.MoveFirst
					  while RCour.EOF=false
					  if ucase(RCour("courier")) = ucase(request("MainCat")) then
					  courierid=RCour("couriersid")
					  session("mcour")=RCour("courier")
		           
					  end if
					  RCour.MoveNext
					  wend
					  end if					  
					  %>
                      </select>
                    </p>
                    <p> 
                      <%
    set RDest =server.CreateObject("adodb.recordset")
	RDest.CursorLocation=aduseclient
                      %>
                      <select name="Destination"  onChange = 'submit()'>
                        <option value="Select Destination">Select Destination</option>
                        <%
                     session("mdest")=""
                     if request("MainCat") <> "" and request("MainCat") <> "Select Courier" then
                     RDest.Open "select * from AvFreights where couriersid="&trim(courierid)&"",conn,adOpenForwardOnly,adLockReadOnly,adcmdtext
	                 set RDest.ActiveConnection=nothing
	                 
                     while RDest.EOF=false %>
                        <option value="<%=RDest("destination")%>"><%=RDest("destination")%></option>
                        <% RDest.MoveNext
                     wend
                     
                     if request("Destination") <> ""  and ucase(request("Destination")) <> ucase("Select Destination")  then
					 
					  RDest.MoveFirst
					  while RDest.EOF=false
					  if ucase(RDest("Destination")) = ucase(request("Destination")) then
					  %>
                        <option value ="<%= request("Destination")%>" selected><%=request("Destination")%></option>
                        <%  
					  Destinationid=RDest("Destinationsid")
					  session("mdest")=RDest("Destination")
					  end if
					  RDest.MoveNext
					  wend
					 end if	
					 end if 
					  %>
                      </select>
                      <%
          
          if trim(courierid) <> "" and trim(Destinationid) <> "" then
        '  Response.Write "courier id is = " & courierid & "<br>"
        '  Response.Write "destainations  id is = " & Destinationid & "<br>"
          
          rsFreight.MoveFirst
          while rsFreight.EOF=false
          
          if trim(rsFreight("couriersid"))= trim(courierid) and trim(rsFreight("destinationsid")) = trim(Destinationid) then
        '  Response.Write "freight  is = " & rsFreight("freights") & "<br>"
          mfreight=rsFreight("freights")
           
          end if
           
          rsFreight.MoveNext
          wend
          
          
          
          end if
         
         
  
         if mfreight = "" then  
         mfreight=0
         end if
         
         if mfreight <> "" then
         
         
         
         
           
         aFreight = trim(mfreight) * trim(TFS_Weight)
         
    
           nTotal=t_tot+afreight  'Grand Total
         
         if trim(aFreight) <> "" then
                   

					session("afreight")=aFreight 
					session("ntotal")=nTotal
					session("t_tot")=t_tot
					      ' Response.Write "<br>afreight is = " & session("afreight") & "<br>"
                          ' Response.Write "<br>ntotal is = " & session("ntotal") & "<br>"
                          ' Response.Write "<br>t_tot is = " & session("t_tot") & "<br>"
           
		 end if	
         end if 
          '  response.write "mrefight is = " & mfreight
            session("mfreight")=mfreight
          %>
                    </p>
                    <p> <br>
                      (Freight x Weight=<%=mfreight%>x<%=wt1%>)</p>
                  </TD>
                  <TD width=127 bgcolor="#FFFFFF" class="text" align="right">Rs. 
                    <%=mid(formatcurrency(aFreight),2,12)%></TD>
                  <TD width=113 bgcolor="#FFFFFF">&nbsp;</TD>
                  <TD width=113 bgcolor="#FFFFFF">&nbsp;</TD>
                </TR>
                <TR> 
                  <TD width=89 bgcolor="#FFFFFF">&nbsp;</TD>
                  <TD width=349 bgcolor="#FFFFFF" align="right" class="text" colspan="2"><b><font color="#0099FF"> 
                    Total Amount</font></b></TD>
                  <TD width=127 bgcolor="#FFFFFF" align="right" class="text"><b><font color="#0099FF"> 
                    Rs. <%=mid(formatcurrency(nTotal),2,12)%></font></b></TD>
                  <TD width=113 bgcolor="#FFFFFF" align="center">&nbsp; </TD>
                  <TD width=113 bgcolor="#FFFFFF" align="center">&nbsp; </TD>
                </TR>
                <%
                'Response.Write "_________" & rsgst("gstcharge")
                
                if trim(rsgst("gstcharge"))>0 then
                %>
                <TR> 
                  <TD width=89 bgcolor="#FFFFFF">&nbsp;</TD>
                  <TD width=349 bgcolor="#FFFFFF" align="right" class="text" colspan="2">Total 
                    Taxable items with Sales Tax(GST) 
                    <input type="text" name="textfield" size="5" maxlength="5" readonly value ="<%=mgst%>" STYLE="text-align:right">
                    %</TD>
                  <TD width=127 bgcolor="#FFFFFF" align="right" class="text"> 
                    <%=mid(formatcurrency(totaltax),2,12)%></TD>
                  <TD width=113 bgcolor="#FFFFFF" align="center">&nbsp; </TD>
                  <TD width=113 bgcolor="#FFFFFF" align="center">&nbsp;</TD>
                </TR>
                <%end if%>
                <TR> 
                  <TD width=89 bgcolor="#FFFFFF">&nbsp;</TD>
                  <TD width=349 bgcolor="#FFFFFF" align="right" class="text" colspan="2"><b><font color="#FF0000">Total 
                    Invoice Value</font></b></TD>
                                <%
                  xGst=ntotal*(cdbl(mGST)/100)
                 '' grandTotal=xgst+chtax+nTotal

				'	if xgst0 <> "" then
					session("freight")=aFreight
					session("mgst")=mGST
		  		
		  		
		  		
		  		
				xgst0=tax_tot+t_tot+chtax+afreight
                session("grandTotal")=xgst0
   
                  session("xgst")=totaltax
                  %>
                  <TD width=127 bgcolor="#FFFFFF" align="right" class="text"><font color="#FF0000"><b><%=mid(formatcurrency(xgst0),2,12)%></b></font> 
                  </TD>
                  <TD width=113 bgcolor="#FFFFFF" align="center">&nbsp;</TD>
                  <TD width=113 bgcolor="#FFFFFF" align="center">&nbsp;</TD>
                </TR>
                 <%
                  dim htax
                  htax= rshtax("htaxcharge")
                  session("htax")=htax
                  'Chtax=round(trim(t_tot+totaltax) / trim(100-trim(rsextras("htax"))) * 100-trim(t_tot+totaltax),2)
				  Chtax=round(trim(xgst0) / trim(100-trim(rshtax("htaxcharge"))) * 100-trim(xgst0),2)
				  session("Chtax")=chtax
				  Chtax=mid(formatcurrency(Chtax),2,12)
				 'response.write xgst0
				  %>
				
				
				<%
                'Response.Write "_________" & rshtax("htaxcharge")
                
                if trim(rshtax("htaxcharge"))>0 then
                %>
			   
			    <TR> 
                  <TD width=89 bgcolor="#FFFFFF">&nbsp;</TD>
                  <TD width=349 bgcolor="#FFFFFF" align="right" class="text" colspan="2"><b>With 
                    holding Tax 
                    <input type="text" name="htaxT" size="5" maxlength="5" readonly value="<%=rshtax("htaxcharge")%>" style="text-align:right">
                    %</b></TD>
                  <TD width=127 bgcolor="#FFFFFF" align="right" class="text"><font color="#FF0000"><b>
                   
                    <%=mid(formatcurrency(Chtax),2,12)%> </b></font></TD>
                  <TD width=113 bgcolor="#FFFFFF" align="center">&nbsp;</TD>
                  <TD width=113 bgcolor="#FFFFFF" align="center">&nbsp;</TD>
                </TR>
              <%end if%>
			    <TR> 
                  <TD width=89 bgcolor="#FFFFFF">&nbsp;</TD>
                  <TD width=349 bgcolor="#FFFFFF" align="right" class="text" colspan="2">&nbsp;</TD>
                  <TD width=127 bgcolor="#FFFFFF" align="right" class="text"><font color="#FF0000"><b> 
                    </b></font></TD>
                  <TD width=113 bgcolor="#FFFFFF" align="center">&nbsp; </TD>
                  <TD width=113 bgcolor="#FFFFFF" align="center"> 
                    <input type="submit" name="checkout" value="Checkout">
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
                    <DIV align=center> </DIV>
                  </TD>
                  <TD width=204> 
                    <DIV align=center> </DIV>
                  </TD>
                  <TD width=81> 
                    <DIV align=center> </DIV>
                  </TD>
                  <TD height=18 width=179> 
                    <DIV align=center><A 
                  href="acart.asp" 
                  onclick="return Checkout()"> </A></DIV>
                  </TD>
                </TR>
                </TBODY> 
              </TABLE>
              <%
				
                session("wt1")=wt1
                  
                  
                '  Response.Write "<br> wt1 is = " & session("wt1") & "<br>"
     

				
				if trim(request("checkout"))<> ""  then 
				Response.Redirect "aprocess_cartCheckout.asp?cartida=" & session("cartida") & "&custID=" & session("CustomerID") & "&DealID=" & session("dealerID")
				%>
              <script language="javascript">
               // alert ("i am checkout")
                </script>
              <%end if%>
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
  </TBODY>
</TABLE>
</html>

<%end if%>