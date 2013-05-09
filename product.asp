<%@ Language=VBScript %>

<link href="acss/newstyle.css" rel="stylesheet" type="text/css" />
<div class="ProductPage">

 <!--#INCLUDE FILE="header.asp"-->

<%
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
'on error resume next

mitemno=request("pid")
set conn = server.CreateObject ("ADODB.Connection")
set rsMain = server.CreateObject("ADODB.Recordset")
set rsSub = server.CreateObject("ADODB.Recordset")
set rsData = server.CreateObject("ADODB.Recordset")
set rsDriver = server.CreateObject("ADODB.Recordset")

%>
<!--#INCLUDE FILE="dbconnect.asp"-->
<%
conn.Open con 
SQLtxt3="select * from vAllData where productid='"&mItemno&"'" 
	rsData.CursorLocation = adUseClient
	rsData.Open SQLtxt3,conn,adOpenForwardOnly,adLockReadOnly,adcmdtext
	set rsData.ActiveConnection = nothing%>
<%if (rsdata.BOF and rsdata.EOF) then
 response.write "Wrong Product:"
 response.end
 End if%>
<%ldesc=rsdata("longdescription")%>
<html>
<head>
<title>Baylan Technology Inc. Product Information Page</title>
<%
Function GetHTMLFile(strFileName)
Dim fso, file, FName
FName = Server.MapPath(ldesc)
Set fso = Server.CreateObject ("Scripting.FileSystemObject")
if Not(fso Is Nothing) Then
If fso.FileExists(FName) Then
Set file=fso.OpenTextFile(FName,1)
GetHTMLFile = Trim(file.readAll())
'response.write "sdfsdkj"
Else
'GetHTMLFile =   "dsds" & FName  & vbCrLf 
End If
Set fso = Nothing
'Else
'GetHTMLFile= "bnfjfj" & err.description 
End If
End Function%>
<script language="JavaScript" type="text/javascript">
function onImgErrorSmall(source)
{
source.src = "images/NOIMAGE.GIF";
// disable onerror to prevent endless loop
source.onerror = "";
return true;
}

function onImgErrorLarge(source)
{
source.src = "images/NOIMAGE1.GIF";
// disable onerror to prevent endless loop
source.onerror = "";
return true;
}
</script>
<link href="product/<%=rsdata("catid")%>style.css" rel="stylesheet" type="text/css">
<link rel="stylesheet" type="text/css" media="print" href="product/printstyle.css" />

<%'<link href="product/003style.css" rel="stylesheet" type="text/css">%>



<link rel="stylesheet" href="admin/arshilogin.css">
<% if rsdata("statuscode") = "08" then %>
<td width="300"  align="center">Wrong Product:</td>
<%
else %>




<body>
<script type="text/javascript">
var windowObjectReference = null; // global variable

function openFFPromotionPopup()
{
  if(windowObjectReference == null || windowObjectReference.closed)
  {
    windowObjectReference = window.open("productprint.asp?pID=<%=rsdata("productID")%>",
   "ProductPage", "resizable=yes,scrollbars=yes,status=yes");
  }
  else
  {
    windowObjectReference.focus();
  };
}
</script>



  
     

<div class="l">
  <div class="t">
    <div class="b">
    	  <div class="r">
        <div class="bl">
          <div class="br">
            <div class="tl">
			<div class="tr">
         <table width="800" border="0">
  <tr>
    <td width="580" height="57" class="fstyle4"> <h4><em>&nbsp;&nbsp;&nbsp;&nbsp;<%=rsData("Groupdesc")%></em></h4>    </td>
    <td><img src=<%=rsdata("pbrand")%> width="140" height="39"  hspace="39" vspace="8"></td>
  </tr>
</table>
<table  border="0" cellpadding="0" cellspacing="0" style="margin-left: 34;">
 <tr> 
 <TD width="365" valign="top" class="fstyle1">
 <table  border="0" cellpadding="0" cellspacing="0">
  <td width="4" height="33">&nbsp;</td>
    <td width="360" height="41" colspan="4" class="fstyle3"><div align="center"><%=rsData("Category")%></div></td>
  <tr>
    <td width="4">&nbsp;</td>
        <td width="120" height="5">&nbsp;</td>
                    <td width="150" height="5">&nbsp;</td>
                    <td width="100" height="5">&nbsp;</td>
                    <td width="1" rowspan="12">&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td width="120" class="fstyle1">&nbsp;</td>
    <td width="150" height="20" class="fstyle2">&nbsp;</td>
    <td width="100" rowspan="5" align="center" class="fsZtyle1"><div align="center"><img src=<%=rsData("picturethumb")%> alt="" width="86" height="75" border=0 align="middle" onerror="onImgErrorSmall(this)"></div></td>
    </tr>
  <tr>
    <td height="4">&nbsp;</td>
    <td width="120" height="4" class="fstyle6"></td>
    <td width="150" height="4" valign="top" class="fstyle6"></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td class="fstyle1">&nbsp;</td>
    <td height="20" class="fstyle2">&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td height="4" valign="top" class="fstyle6"></td>
    <td valign="bottom" class="fstyle6"></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td class="fstyle1">&nbsp;</td>
    <td height="20" class="fstyle2">&nbsp;</td>
  </tr>
  <tr>
    
  <td>&nbsp;</td>
    <td height="4" valign="top" class="fstyle6"></td>
    <td valign="bottom" class="fstyle6"></td>
  <td height="22" align="center" class="fstyle2"><div align="centre"><%=(rsData("PRODUCTID"))%></div></td>
   </tr>
  <tr>
    <td>&nbsp;</td>
    <%if not isnull(rsdata("image"))then %>
                <td class="fstyle1"><div align="LEFT"></div></td>
                <%else%>
                <td class="fstyle1">&nbsp;</td>
                <%end if%>
	      <%
dim datasheet, dfname
dfname = Server.MapPath(rsdata("datasheet"))
Set datasheet = CreateObject("Scripting.FileSystemObject") 
If datasheet.FileExists(dfname) Then
%>
<td height="22" align="center">&nbsp;</td>
 <%'  Response.Write("Available")
ELSE
%><td height="22" align="center"><img src="images/SPACER.GIF" border="0"></td>
<% ' Response.Write("Not Available")
End If 
%> 
   <td height="20" class="fXstyle2">&nbsp;</td>

   
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td height="4" valign="top" class="fXstyle6"></td>
    <td valign="bottom" class="fXstyle6"></td>
     <td height="22" class="fXstyle1">&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td height="22" class="fstyle1"><h4><em><u>Description</u></em></h4></td>
    <td height="22" class="fXstyle1">&nbsp;</td>
    <td height="22" class="fXstyle1">&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td height="40" colspan="4" class="fstyle2"><%=rsdata("description")%></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td width="365" height="240" colspan="4" align="center" class="fstyle1"><img src="<%=(rsdata("picturelarge"))%>" alt="" onerror="onImgErrorLarge(this)"></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td colspan="3" class="fstyle1">&nbsp;&nbsp;<%= FormatDateTime(Date, 1) %>&nbsp;<%= FormatDateTime(Now, 4) %>&nbsp;</td>
	  
  </tr>
    </table></TD>
<%if ldesc <> "" and trim(ldesc) <> "PIMAGES\" then%>
<td width="383" height="200"  valign="top" colspan="3" rowspan="10" class="text4"><%Response.Write GetHTMLFile(request("ldesc"))%></td>
<%end if%>
    <td width="22" rowspan="10">&nbsp;</td>
  </tr>
</TABLE>
</div>
</div>
</div>
</div>
</div>
</div>
</div>
</div>


<% end if%>
 <!--#INCLUDE FILE="footer.asp"-->

 </div>


</body>
</html>
