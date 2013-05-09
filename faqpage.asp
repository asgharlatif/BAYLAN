<%@ Language=VBScript %>
<link href="acss/newstyle.css" rel="stylesheet" type="text/css" />
<div class="FaqPage" >

<!--#INCLUDE FILE="header.asp"-->
<br>
<%
'on error resume next
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
<html><head><title>Baylan Technology Inc. Faq/Help</title> 

 <table width="100%" border="0">
  <tr>
  <td><div align="center"> 
  <%  
   Dim objAd
   'Set objAd = Server.CreateObject("MSWC.AdRotator") 'Create the component
   'Response.Write  objAd.GetAdvertisement ("/banners/adrot.txt") 
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
                  href="mailto:support@bbl.com.pk" class="textl">support@bbl.com.pk</a> 
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

	
<br />
 <!--#INCLUDE FILE="footer.asp"-->
  


  </div>

</html>