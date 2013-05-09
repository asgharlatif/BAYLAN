<%@ Language=VBScript %>
 
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
<link href="product/printstyle.css" rel="stylesheet" type="text/css">


<% if rsdata("statuscode") = "08" then %>
<style type="text/css">
<!--
.style3 {
	font-size: 16px;
	font-family: Arial, Helvetica, sans-serif;
}
-->
</style>
<td width="300"  align="center">Wrong Product:</td>
<%
else %>

<body onLoad="window.print()">
<table class="ac00" width="800" border="0">
  <tr>
    <td width="580" height="57"><h2 class="style3">&nbsp;&nbsp;&nbsp;&nbsp;<strong><em><u><%=rsData("Groupdesc")%></u></em></strong></h2></td>
    <td><img src=<%=rsdata("pbrand")%> width="140" height="39"  hspace="39" vspace="8"></td>
  </tr>
  <tr>
    <td width="620px" height="40
	"  colspan="2" >
    
     <table style="width:788px"><tr>
     <td class="LabelClass">&nbsp;</td>
         <td  style="text-align:right">Prod.ID&nbsp;&nbsp;<span class="ValueClass"><b><%=(rsData("itemcode"))%></b></span></td></tr></table>
        
    
    </td>
  </tr>

</table>
<table  border="0" cellpadding="0" cellspacing="0">

 <tr>
     <td width="365" valign="top">
         <table border="0" cellpadding="0" cellspacing="0">
             <tr>
                 <td height="40" colspan="3" class="ac0" style="padding:10px">
                     <%=rsdata("description")%>
                 </td>
             </tr>
             <tr>
                 <td width="365" height="240" colspan="3" align="center" class="ac5">
                     <img src="<%=(rsdata("picturelarge"))%>" alt="" onerror="onImgErrorLarge(this)">
                 </td>
             </tr>
             <tr>
                 <td colspan="3" align="center" class="ac1">&nbsp;<%=(rsData("PRODUCTID"))%>
                 </td>
             </tr>
  <td class="LabelClass"><p>&nbsp;</p>
    <p>&nbsp;</p>
    <p>&nbsp;&nbsp;&nbsp;<%= FormatDateTime(Date, 1) %>&nbsp;<%= FormatDateTime(Now, 4) %>&nbsp;</p></td>
         </table>
     </td>
<%if ldesc <> "" and trim(ldesc) <> "PIMAGES\" then%>
<td width="414" height="200"  valign="top" rowspan="10" class="ac6" style="padding:10px"><%Response.Write GetHTMLFile(request("ldesc"))%></td>
	<%end if%>
  </tr>
</TABLE>

<%'end if%>
<%'else%>
<% end if%>
</body>
</html>
