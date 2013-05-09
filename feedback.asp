<%@ Language=VBScript %>

<link href="assets/css/bootstrap.css" rel="stylesheet" type="text/css" />

<link href="acss/newstyle.css" rel="stylesheet" type="text/css" />
<div class="FeedBackPage">

<!--#INCLUDE FILE="header.asp"-->
<br>
<%
 if Request.Cookies("isDealer") <> "" then
 xmsg="You are Logged in as " & Request.Cookies("typename") & " : [" & Request.Cookies("dealer_id") & "]"
 end if
%>
<!--#INCLUDE FILE="dbconnect.asp"-->
<%
aaj=date
stype=request("stype")
commentON=request("commentON")
other=request("other")
suggestion=request("suggestion")
name=request("name")
email=request("email")
ph=request("ph")
fax=request("fax")
urgency=request("urgency")


Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
set conn = server.CreateObject ("ADODB.Connection")
set RSgalSQL = server.CreateObject("ADODB.Recordset")
conn.Open con 
SQLtxt="select * from suggestion" 
rsgalSQL.CursorLocation = adUseClient
rsgalSQL.Open SQLtxt,conn,adOpenForwardOnly,adLockReadOnly,adcmdtext
set rsgalSQL.ActiveConnection = nothing


if request("name") <> "" and request("email") <> "" then
	'Conn.Open con
	galSQL="insert into suggestion values('"&name&"','"&email&"','"&ph&"','"&fax&"','"&commentON&"','"&suggestion&"','"&stype&"','"&urgency&"','"&aaj&"','"&other&"')"
	conn.Execute(galSQL)
mailfrom=Request.Form("email")




'''for sending email
mb=mb & "<body><html>"
mb=mb & "<head>"
mb=mb & "<title>send email</title>"
mb=mb & "<div align=left><font size=3 face=Arial, Helvetica, sans-serif color=#000000><b>&nbsp;SUGGESSTION DETAILS</b></font><br>"
mb=mb & "</div>"
mb=mb & "<table width=100% border=1 cellspacing=0 cellpadding=2 align=center bgcolor=#CCCCCC bordercolor=#FFFFFF>"
mb=mb & "<tr>"
mb=mb & "<td width=123>" 
mb=mb & "<p><b><font face=Verdana, Arial size=2>Comments Type:</font></b></p>"
mb=mb & "</td>"
mb=mb & "<td width=107><font face=Verdana, Arial size=2 color=#0000CC>"&stype&" </font></td>"
mb=mb & "<td width=36><b><font face=Verdana, Arial size=2>Name:</font></b></td>"
mb=mb & "<td width=236><font face=Verdana, Arial size=2 color=#0033FF>"&Name&" </font></td>"
mb=mb & "</tr>"
mb=mb & "<tr>" 
mb=mb & "<td width=123><b><font face=Verdana, Arial size=2>Comments At:</font></b></td>"
mb=mb & "<td width=107><font face=Verdana, Arial size=2 color=#0000CC>"&commentON&"</font></td>"
mb=mb & "<td width=36><b><font face=Verdana, Arial size=2>Email:</font></b></td>"
mb=mb & "<td width=236><font face=Verdana, Arial size=2 color=#0033FF>"&email&"</font></td>"
mb=mb & "</tr>"
mb=mb & "<tr>"
mb=mb & "<td width=123><b><font face=Verdana, Arial size=2>Others Comments At:</font></b></td>"
mb=mb & "<td width=107><font face=Verdana, Arial size=2 color=#0000CC>"&other&"</font></td>"
mb=mb & "<td width=36><b><font face=Verdana, Arial size=2>Urgency</font></b></td>"
mb=mb & "<td width=236><font face=Verdana, Arial size=2 color=#0033FF>"&urgency&"</font></td>"
mb=mb & "</tr>"
mb=mb & "<tr>"
mb=mb & "<td colspan=2 rowspan=2><b><font face=Verdana, Arial size=2></font></b><font face=Verdana, Arial size=2></font><b><font face=Verdana, Arial size=2></font></b><font face=Verdana, Arial size=2></font></td>"
mb=mb & "<td width=36><b><font face=Verdana, Arial size=2>Ph:</font></b></td>"
mb=mb & "<td width=236><font face=Verdana, Arial size=2 color=#0033FF>"&ph&"</font></td>"
mb=mb & "</tr>"
mb=mb & "<tr>"
mb=mb & "<td width=36><b><font face=Verdana, Arial size=2>Fax:</font></b></td>"
mb=mb & "<td width=236><font face=Verdana, Arial size=2 color=#0033FF>"&fax&"</font></td>"
mb=mb & "</tr>"
mb=mb & "<tr>"
mb=mb & "<td width=123><b><font face=Verdana, Arial size=2>Comments:</font></b></td>"
mb=mb & "<td colspan=3><font face=Verdana, Arial size=2 color=#0000FF>"&suggestion&"</font></td>"
mb=mb & "</tr>"
mb=mb & "<tr>"
mb=mb & "</body></html>"

Dim Mail, Body
Set Mail  = Server.CreateObject("CDO.Message")
set rsEmail= server.CreateObject("ADODB.Recordset")
Setxt="select * from aextras" 
rsemail.CursorLocation = adUseClient
rsemail.Open Setxt,conn,adOpenForwardOnly,adLockReadOnly,adcmdtext
set rsemail.ActiveConnection = nothing
'mail.Host =rsemail("ip")
mail.From = mailfrom
'mail.FromName = "ARSHI COMPUTER (Subscribe)"
Mail.To rsemail("feedback")
Mail.Subject = "Feedback Form" 
'Mail.IsHTML = True
Mail.HtmlBody = "<HTML><BODY>" & mb & "</BODY></HTML>"
Mail.Send
If Err <> 0 Then 
    Response.Write "An error occurred: " & Err.Description 
End If

Set Mail = nothing
end if






'messagesubject="Feedback Form" 


		
'Dim Mail, Body
'Set Mail = Server.CreateObject("Persits.MailSender")
'Mail.Host =rsemail("ip")
'Mail.AddAddress =rsemail("feedback") 'email
'Mail.FromName = "ARSHI COMPUTER (Subscribe)"
'Mail.From = mailfrom
'Mail.Subject = "Congratulations "
'Mail.Subject = messagesubject
'Mail.IsHTML = True.IsHTML = True
'Mail.Body = "<HTML><BODY>" & mb & "</BODY></HTML>"
'Mail.Send
'If Err <> 0 Then 
'    Response.Write "An error occurred: " & Err.Description 
'End If
'Set Mail = nothing

'''end sending email	
	
	
'Response.Write "<script language=javascript>alert('Your Comments have been sent to Support Department')</script>"

%>
<SCRIPT language=JavaScript>
<!--

function PlaceM(thisv){
if (thisv.checked==true){

document.Form1.commentON.value=document.Form1.other.value
}
}


function Others(){
if (document.form1.other.value != " "){
document.form1.commentON.value='Others'
}
}                     


function CheckBlan(){
var correctEmail=0;
var CNameOK = 1;
var emaillength = 0;

if (document.form1.name.value==""){
alert ('Please enter your name.')
document.form1.name.focus()
CNameOK=0
}

if (CNameOK==1) {
emaillength=document.form1.email.value.length;
	if (emaillength==0)
	alert ('Please provide the email address for feedback.')
	document.form1.email.focus()

if (emaillength != 0 ){
    for ( var i=0; i<emaillength;i++){
    if (document.form1.email.value.charAt(i) == "@")
	correctEmail=1;
    }
       if (correctEmail==0)
	   alert("Incorrect Email Address.")
	   else{
	   document.form1.submit()
	   }}
}
}                




//-->
</SCRIPT>



<html><head><title>Baylan Technology Inc. Feedback Page</title> 
 <table width="100%" border="0">
  <tr>
  <td><div align="center"> 
  <%  
   Dim objAd
   'Set objAd = Server.CreateObject("MSWC.AdRotator") 'Create the component
   'Response.Write  objAd.GetAdvertisement ("/banners/adrot.txt") 
   'Set objAd = Nothing 'Destroy the object
   %>
</div>
 </td></tr></table>
  <tr></tr>
<br>

<div>
Feed Back
</div>

<form action="feedback.asp" method="post"   id=form1 name=form1>
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="139">
  <tr> 
    <td valign="top" height="11"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><b><font color="#FF6600" face="Verdana, Arial" size="2">What 
      Kind of Comments would you send ?</font></b></font></td>
  </tr>
  <tr> 
    <td height="16" valign="top"> 
      <table width="68%" border="0" cellspacing="0" cellpadding="0" height="26">
        <tr> 
          <td width="24" height="17"> 
            <input type="radio" name="stype" value="suggestion" checked>
          </td>
          <td width="94" height="17"><font face="Verdana, Arial, Helvetica, sans-serif" size="1">Suggestions</font> 
          </td>
          <td width="14" height="17"> 
            <input type="radio" name="stype" value="complaint">
          </td>
          <td width="108" height="17"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">Complaints</font></td>
          <td width="24" height="17"> 
            <input type="radio" name="stype" value="problem">
          </td>
          <td width="81" height="17"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">Problems</font></td>
          <td width="24" height="17"> 
            <input type="radio" name="stype" value="Praise">
          </td>
          <td width="67" height="17"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">Praise</font></td>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
    <td height="12"> 
        <table width="100%" border="0" cellspacing="0" cellpadding="0" height="120">
          <tr> 
            <td colspan="6" height="24"><font size="2"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#FF6600">What 
              about us do you want to comment on ?</font></b></font> 
              <div align="right"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><b></b></font></div>
            </td>
          </tr>
          <tr> 
            <td width="17%" height="18" valign="top"> 
              <select name="commentON">
                <option value="Web Site">Web site</option>
                <option value="Products">Products</option>
                <option value="Company">Company</option>
                <option value="Support Service">Support Service</option>
				<option value="Lost Password ">Lost Password</option>
                <option value="Others">Others</option>
 			  </select>
              <b><font face="Verdana, Arial, Helvetica, sans-serif" size="1">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font></b></td>
            <td width="16%" height="18" valign="top"> 
              <div align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1">&nbsp;&nbsp;Other:</font></b> 
              </div>
            </td>
          			<td colspan="4" height="18" valign="top"> &nbsp; 
              <input  class=mono maxlength="50" name="other" size="15" onBlur="Others()">
              <font size="1" face="Verdana, Arial, Helvetica, sans-serif"><b></b></font></td>
          </tr>
          <tr> 
            <td colspan="3" height="15" valign="top"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><b><font size="2" color="#FF6600">Ent</font></b></font><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><b><font size="2" color="#FF6600">er 
              your comments in the space below</font></b></font></td>
            <td colspan="3" height="15" valign="top"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><b><font color="#FF6600" face="Verdana, Arial" size="2">Tell 
              us how to get in touch with you ? </font></b></font></td>
          </tr>
          <tr> 
            <td colspan="3" height="32" rowspan="4" valign="top"> <b> 
              <textarea cols=30 name=suggestion rows="3"></textarea>
              <br>
              <font size="2"><b><font size="2"> 
              <input type="radio" name="urgency" value="urgent">
              <font face="Verdana" color="#FF6600"><b> Please contact me as soon 
              as possible.</b></font></font> </b></font> <br>
              </b><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><b> 
              <font face="Verdana, Arial, Helvetica, sans-serif" size="1"><b> 
              </b></font></b></font></b> </td>
            <td width="1%" height="14">&nbsp;</td>
            <td height="14" width="6%" valign="middle"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1">Name</font></b> 
            </td>
            <td height="14" width="41%" valign="top"> 
              <div align="left">&nbsp; 
                <input name="name" size="25" >
                *</div>
            </td>
          </tr>
          <tr> 
            <td width="1%" height="12">&nbsp;</td>
            <td width="6%" height="12" valign="middle"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1">Email</font></b> 
            </td>
            <td width="41%" height="12" valign="top"> 
              <div align="left">&nbsp; 
                <input name="email" size="25" >
                *</div>
            </td>
          </tr>
          <tr> 
            <td width="1%" height="11"><b></b><b><font size="2"> </font><font face="Verdana, Arial, Helvetica, sans-serif" size="1"> 
              </font></b></td>
            <td width="6%" height="11" valign="middle"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1">Tel 
              </font></b></td>
            <td width="41%" height="11" valign="top"> 
              <div align="left">&nbsp; 
                <input name="ph" size="25" >
              </div>
            </td>
          </tr>
          <tr> 
            <td width="1%" height="21">&nbsp;</td>
            <td width="6%" height="21" valign="middle"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1">Fax 
              </font></b></td>
            <td width="41%" height="21" valign="top"> 
              <div align="left"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1"> 
                &nbsp; &nbsp; 
                <input name="fax" size="25" >
                <b> </b> </font></b> </div>
            </td>
          </tr>
        </table>


      <div align="right" style="padding-right:185px; padding-top:50px" ><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><b> 


        <input type="button" name="Submit" value="Submit" onclick="CheckBlan()" class="btn btn-primary buttons">
        </b></font> &nbsp;&nbsp; 
        <input type="reset" name="Submit2" value="Clear Form" class="btn btn-primary buttons">


        </b></font></b></div>
    </td>
  </tr>
</table>
</form>
<br>
  <div align="center"></div>
  <tr valign="middle"> 
    
 	<div align="center"> 
    <table width="100%" border="0" cellspacing="0" cellpadding="0" height="1" background="aimages/dot_grey1.gif">
    <tr> 
   
    </tr>
    </table>


    <!--#INCLUDE FILE="footer.asp"-->
  
</body>
</html>

</div>