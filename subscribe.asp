<html>
<head>
<script language="JavaScript">

function CheckBlan(){
var correctEmail=0;
var CNameOK = 1;
var emaillength = 0;
var x1 =  (screen.width-300)/2;
var y1 =  (screen.height-300 )/2;

if (CNameOK==1) {
emaillength=document.homeP.Email.value.length;
	if (emaillength==0)
	alert ('Please provide the email address for subscription.')
	document.homeP.Email.focus()

if (emaillength != 0 ){
    for ( var i=0; i<emaillength;i++){
    if (document.homeP.Email.value.charAt(i) == "@")
	correctEmail=1;
    }
       if (correctEmail==0)
	   alert("Incorrect Email Address")
	   else{
	   document.cookie = "emailT=" + escape(document.homeP.Email.value);
	   document.cookie = "NameT=" + escape(document.homeP.NameT.value);
	   document.cookie = "subt=" + escape(subt);
	   childWin= open('ap_visitors.asp','childWin','screenX='+ x1 + ',screenY=' + y1 + ',left='+ x1 + ',top='+ y1 + ',width=420,height=100,resizable=no,toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no');  
	   }}
}
}                
var subt = 'sc'
function checkboxo(pickit){
subt=pickit.value
//alert (subt)
}


                     //--></script>
<link rel="stylesheet" href="acss/standard.css" type="text/css">
<title>Baylan Technology Inc. Subscribe / UnSubscribe Window:</title><meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"><style type="text/css">
<!--
body {
	background-color: #B5C7EF;
}
-->
</style></head>
<body leftmargin="0" topmargin="0"  >
<form name="homeP" method="post" onSubmit="CheckBlan(); return false">
  <table background="" border=0 cellpadding=0 cellspacing=0 
style="WIDTH: 505px; HEIGHT: 152px" width="505" bordercolor="#6699cc">
    <tr bgcolor="#000000"> 
      <td height="56" colspan="3"> 
      <h1><font size="6" face="Verdana, Arial, Helvetica, sans-serif" color="#cc6633">Subscribe / <font face="Verdana, Arial">UnSubsribe: </font></font></h1></td>
    </tr>
    <tr> 
      <td  bgcolor="#FFFFFF">&nbsp;</td>
      <td  bgcolor="#FFFFFF">&nbsp;</td>
      <td colspan="2" bgcolor="#FFFFFF">&nbsp;</td>
    </tr>
    <tr> 
      <td  bgcolor="#FFFFFF"></td>
      <td  bgcolor="#FFFFFF"> 
        <div align="right"><font color="#ff0000" size="2"></font></div>
      </td>
      <td colspan="2" align="middle"  bgcolor="#FFFFFF"> 
        <div align="left"><font color="#ff0000" size="2"></font></div>
      </td>
    </tr>
    <tr> 
      <td  bgcolor="#FFFFFF"></td>
      <td  bgcolor="#FFFFFF">&nbsp;</td>
      <td colspan="2" bgcolor="#FFFFFF">&nbsp;</td>
    </tr>
    <tr>
      <td ></td>
      <td >&nbsp;</td>
      <td colspan="2"><font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#CC3333">&nbsp;        </font></td>
    </tr>
    <tr> 
      <td ></td>
      <td ><b><font color="#333399"><font color="#0000cc"><span class="text"><font color="#cc6633">        Email:</font></span></font></font></b><font color="#cc6633"><span class="text">&nbsp; 
        </span> </font></td>
      <td colspan="2"> 
          <input type="text" name="Email" size="20">
      </td>
    </tr>
    <tr> 
      <td ></td>
      <td ><b><font color="#333399"><font color="#0000cc"><span class="text"><font color="#cc6633">Name:</font></span></font></font></b><font color="#cc6633"><span class="text"></span></font></td>
      <td colspan="2"> 
        <input type="text" name="NameT" size="20">
      </td>
    </tr>
    <tr> 
      <td ></td>
      <td ></td>
      <td colspan="2"> 
        <div align="left">        </div>
      </td>
    </tr>
    <tr> 
      <td  align="right"></td>
      <td  align="right"></td>
      <td align="centre" class="text4"> <input type="radio" name="emailstatus" value="sc" checked onClick="checkboxo(this)">
                Subscribe</td>
    </tr>
    <tr> 
      <td  align="right">&nbsp;</td>
      <td  align="right"></td>
      <td  align="centre" class="text4" height="19"> 
    
                <input type="radio" name="emailstatus" value="nc" onClick="checkboxo(this)">
                Unsubscribe</td>
    </tr>
    <tr> 
      <td  align="right" height="19"> 
        <div align="center"></div>
      </td>
      <td  align="middle" height="19"> 
        <div align="center"></div>
      </td>
      <td  align="left" height="19"> 
        <div align="left"></div>
      </td>
    </tr>
    <tr> 
      <td  align="right" height="10">&nbsp;</td>
      <td  align="middle" height="10">&nbsp;</td>
      <td  align="left" height="10"><input  name="submitB" type="submit" value="Submit" onClick=window.close()>
        <input id=reset1 name=reset122 type=button value=Close onClick=window.close()></td>
    </tr>
  </table>
  
		

  
  
  
  
</FORM>
</body></HTML>
