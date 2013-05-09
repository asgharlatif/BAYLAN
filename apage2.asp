<%@ Language=VBScript %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"><head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />

    
    <script src="JQuery/jquery-1.8.3.min.js" type="text/javascript"></script>
    <script src="yJLib/DomManuplation.js" type="text/javascript"></script>
    
    <script type="text/javascript" src="http://ajax.googleapis.com/ajax/libs/jquery/1.7.1/jquery.min.js"></script>
    <link href="modalPopLite1.3.1/modalPopLite.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="modalPopLite1.3.1/modalPopLite.min.js"></script>
    <script src="modalPopLite1.3.1/HelloWorld.js" type="text/javascript"></script>
    


   <script language="javascript" type="text/javascript" >
       $(function () {
           $(".BtnFeature").on("click", abc);
           $(".BtnAllDrivers").on("click", xyz);
       });

       function xyz() {


           var item = $(this).next(".loader");
          
           var innerhtm = item.html();


           $('#dialog-modal p').html(innerhtm);
           $('#dialog-modal').css("display", "block");
           $('#dialog-modal').modalPopLite({ openButton: '.BtnAllDrivers', closeButton: '#close-btn' });
       }

       function abc() {
           var item = $(this).next(".loader");
           var innerhtm = item.html();
           $('#dialog-modal p').html(innerhtm);
           $('#dialog-modal').css("display", "block");
           $('#dialog-modal').modalPopLite({ openButton: '.BtnFeature', closeButton: '#close-btn' });
           
       }

    </script>

    <script type="text/javascript">

    var windowObjectReference = null; // global variable

    function myfunction( parameter1) {
        alert(parameter1);
    }


    function openFFPromotionPopup(parameter1) {
        if (windowObjectReference == null || windowObjectReference.closed) {
            windowObjectReference = window.open("productprint.asp?" + parameter1, "ProductPage", "resizable=yes,scrollbars=yes,status=yes");
        }
        else {
            windowObjectReference.focus();
        };
    }

  


    </script>


<head>

<%
'ON ERROR RESUME NEXT
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
mcd=request("mcd") ' data from cart2
if mcd <> "" then
myrange=split(mcd,"-")
for i =0 to ubound(myrange)
	p1=myrange(0)  'mcode to view same selected catgory as before o ca
	P2=myrange(1)  'subcategory id
	P3=myrange(2)  'total
	P4=myrange(3)  'no of items
	P5=myrange(4)  'Basket ID
next
session("scode")=p1
scode=p1   'main catid
xdata=p2   'sbcat id
session("tota")=p3
session("nitemsa")=p4
session("cartida")=p5
else
scode=request("mcode")
xData=request("mData")
end if
if request("mData") <> ""  then
mData=request("mData")
session("mData_")=request("mData")
xData=mData
end if
if request("mData") = ""  and Request.QueryString("pgt") = "" then 
mData=session("mData_")
xData=mData
end if
a=request("a")
'------------Parameters for order procedure-----------
pid=request("pid")				'product id
price=request ("price")		'price
'----------------------------------------------------
set conn = server.CreateObject ("ADODB.Connection")
set rsMain = server.CreateObject("ADODB.Recordset")
set rsSub = server.CreateObject("ADODB.Recordset")
set rsData = server.CreateObject("ADODB.Recordset")
set rsDriver = server.CreateObject("ADODB.Recordset")
set RScartid = server.CreateObject("ADODB.Recordset")
set RScartitems = server.CreateObject("ADODB.Recordset")
set rsVQty = server.CreateObject("ADODB.Recordset")
 if Request.Cookies("isDealer") <> "" then
 xmsg="You are Logged in as " & Request.Cookies("typename") & " : [" & Request.Cookies("dealer_id") & "]"
 end if
%>
<!--#INCLUDE FILE="dbconnect.asp"-->

<% 



sub ListFolderContents(path)
       
   
     
     dim fs, folder, file, item, url,present

     present=0

     set fs = CreateObject("Scripting.FileSystemObject")
     set folder = fs.GetFolder(path)

  
     if present=0 then
         for each item in folder.SubFolders
                    ListFolderContents(item.Path)
         next
     end if
   
    for each item in folder.Files
        url = MapURL(item.path)
        
         If Instr(item.Name, productid) Then
          present=1
          'response.Write(1)
          DriverPresent=DriverPresent+"1"
          exit sub

        end if
    next
    
    'response.Write(0)       

   end sub


   function MapURL(path)

     dim rootPath, url

     'Convert a physical file path to a URL for hypertext links.

     rootPath = Server.MapPath("/")
     url = Right(path, Len(path) - Len(rootPath))
     MapURL = Replace(url, "\", "/")

end function %>

<%
conn.Open con 

if trim(Request.Cookies("isDealer")) <> "" or trim(Request.Cookies("isMember")) <> "" then
'SQLtxt_="select * from DCategory where category='"& Request.Cookies("dcategory") &"'" 
set RsPricing=server.CreateObject("adodb.recordset")
RsPricing.CursorLocation = adUseClient
RsPricing.Open SQLtxt_,conn,adOpenForwardOnly,adLockReadOnly,adcmdtext
set RsPricing.ActiveConnection = nothing
price=RsPricing("mration")
price__=RsPricing("mration")
session("price__")=RsPricing("mration")
end if

SQLtxt="select * from VmainCAT where catid in (select catm  from item where brshds='BayLan' group by catm) order by sort_it" 
rsMain.CursorLocation = adUseClient
'rsMain.Open SQLtxt,conn,adOpenForwardOnly,adLockReadOnly,adcmdtext
'set rsMain.ActiveConnection = nothing
set rsmain=conn.Execute(SQLtxt)
'SQLtxt2="select * from VSUBCat where CatID='"&Scode&"' order by category" 
'rsSub.CursorLocation = adUseClient
'rsSub.Open SQLtxt2,conn,adOpenForwardOnly,adLockReadOnly,adcmdtext
'set rsSub.ActiveConnection = nothing
'	if not rssub.eof then
'		msid2=rssub("catcode")
'	end if




baylenProduct="  productid in (select itcd from item where brshds='BayLan' group by itcd) "


if a = "" then
if xData <> "" then

if Request.Cookies("isDealer") <> "" or Request.Cookies("isMember") <> "" then
	SQLtxt3="select productid,item,image,itemcode,picturethumb,picturelarge,datasheet,warranty,brand,description,longdescription,driver,retail,pga,statuscode,support,manual from vAllData where "& baylenProduct &" and Catcode='"&xData&"' and statuscode != '08' order by sorder,brand,description" 	
elseif request.QueryString("itcd") <>"" then
    SQLtxt3="select productid,item,image,itemcode,picturethumb,picturelarge,datasheet,warranty,brand,description,longdescription,driver,retail,pga,statuscode,support,manual from vAllData where "& baylenProduct &" and productid='"&request.QueryString("itcd")&"' and statuscode != '08' order by sorder,brand,description" 	
else
    SQLtxt3="select productid,item,image,itemcode,picturethumb,picturelarge,datasheet,warranty,brand,description,longdescription,driver,retail,pga,statuscode,support,manual from vAllData where "& baylenProduct &" and Catcode='"&xData&"' and statuscode != '08' order by sorder,brand,description" 	

'''''-----1 response.Write SQLtxt3

end if

''response.Write "here at three" & SQLtxt3 & "--------"    

	rsData.CursorLocation = adUseClient
	rsData.Open SQLtxt3,conn,adOpenForwardOnly,adLockReadOnly,adcmdtext
	set rsData.ActiveConnection = nothing
    
               mloop1=rsdata.RecordCount/3
               mloop=int(rsdata.RecordCount/3)
               if mloop1 > mloop then
				mloop3=mloop+1
               end if

				if rsdata.EOF=false then
				if Request.Cookies("isDealer") <> "" then
					mprice=rsdata("pga")
				End if
				if Request.Cookies("isMember") <> "" then
					mprice=rsdata("retail")
				End if
                end if  
else
'	SQLtxt3="select * from AvAllData where Catid='"&scode&"'" 
if Request.Cookies("isDealer") <> "" or Request.Cookies("isMember") <> "" then
	'SQLtxt3="select productid,item,image,itemcode,brand,picturethumb,picturelarge,warranty,description,datasheet,longdescription,driver,[dbo].Editprice(retail * "& price &") as retail,[dbo].Editprice(pga * "& price &") as pga,statuscode,support,manual from vAllData where "& baylenProduct &" and Catid='"&scode&"' and statuscode != '08' order by sorder,brand,description" 
elseif request.QueryString("itcd") <>"" then
    SQLtxt3="select productid,item,image,itemcode,picturethumb,picturelarge,datasheet,warranty,brand,description,longdescription,driver,retail,pga,statuscode,support,manual from vAllData where "& baylenProduct &" and productid='"&request.QueryString("itcd")&"' and statuscode != '08' order by sorder,brand,description" 	
else
SQLtxt3="select productid,item,image,itemcode,brand,picturethumb,picturelarge,datasheet,warranty,driver,description,longdescription,retail,pga,statuscode,support,manual from vAllData where "& baylenProduct &" and Catid='"&scode&"' and statuscode != '08' order by sorder,brand,description" 

'''''-----2response.Write SQLtxt3
end if

	rsData.CursorLocation = adUseClient
	rsData.Open SQLtxt3,conn,adOpenForwardOnly,adLockReadOnly,adcmdtext
	set rsData.ActiveConnection = nothing
               mloop1=rsdata.RecordCount/3
               mloop=int(rsdata.RecordCount/3)
               if mloop1 > mloop then
				mloop3=mloop+1
              end if
end if
else
if Request.Cookies("isDealer") <> "" or Request.Cookies("isMember") <> "" then
	'SQLtxt3="select *  ,[dbo].Editprice(pga * "& price &") as pga  from vAllData where productid='"&pid&"'" 
else
SQLtxt3="select  *  from vAllData where "& baylenProduct &" and productid='"&pid&"' order by sorder,brand,description" 
'''response.Write "here at one"
end if

'''''-----3 response.Write SQLtxt3

	'SQLtxt3="select productid,image,itemcode,picturethumb,description,retail,pga,statuscode from AvAllData where Catid='"&scode&"' and statuscode != 'S008'" 
	rsData.CursorLocation = adUseClient
	rsData.Open SQLtxt3,conn,adOpenForwardOnly,adLockReadOnly,adcmdtext
if rsdata.eof=false then

pid=rsdata("productid")				'product id
name=rsdata("item")
price=request("price")		'price
pic=rsdata("picturelarge")
'movie=rsdata("specs")				'large image
ldesc=rsdata("longdescription")		'long description
stimage=rsdata("image")	'status image like NEW
'mScat=rsdata("mscat")	'subcategory id for qty unit and items in subcategory
brand=rsdata("brand")
warranty=rsdata("warranty")	
'mdriver=request("mdrvr")	
mdsheet=rsdata("datasheet")	

end if
	'SQLtxt4="select name,driverURL from Adrivers where productID='"&pid&"'" 
	'rsDriver.CursorLocation = adUseClient
	'rsDriver.Open SQLtxt4,conn,adOpenForwardOnly,adLockReadOnly,adcmdtext
end if

'''response.Write SQLtxt3

%>
<title>Baylan Technology Inc. Products Page</title> 
<link rel="stylesheet" href="acss/standard.css" type="text/css">
<link rel="stylesheet" type="text/css" href="acss/newstyle.css" />
    <link href="assets/css/bootstrap.css" rel="stylesheet" type="text/css" />

</head>
<body>

    <div id="dialog-modal" title="Basic modal dialog" style="display: none; width: 600px; height: 500px;" class="dialogue">
       <div class="dialogueclosebutton">
            <a id="close-btn" href="#" class="btn btn-inverse">X</a>
        </div>
        <div class="MoveableDic">
            <p>
          
            </p>
        </div>
    </div>

   


		  <script language="javascript"> 
function toggle(showHideDiv, switchTextDiv) {
	var ele = document.getElementById(showHideDiv);
	var text = document.getElementById(switchTextDiv);
	if(ele.style.display == "block") {
    		ele.style.display = "none";
		text.innerHTML = '<img src="plus.gif" border="0">&nbsp;<img src="images/readmore.gif" border="0" align="middle" />';
  	}
	else {
		ele.style.display = "block";
		text.innerHTML = '<img src="minus.gif" border="0">&nbsp;<img src="images/readmore.gif" border="0" align="middle" />';
	}
}
</script>

<%
Function GetHTMLFile(strFileName)
    Dim fso, file, FName
    FName = Server.MapPath(rsdata("longdescription"))
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
End Function


Function GetDRIVERFile(strFileName,ProductId,description,itemcode)

    url = "http://localhost:200/driverdetail.asp?productid="+ProductId +"&description="+ description + "&itemcode="+itemcode
     'url = "http://localhost/Web-Bay/driverdetail.asp?productid="+ProductId 
    'url = "http://baylan.com/driverdetail.asp?productid="+ProductId  
 
    set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP") 
    xmlhttp.open "GET", url, false 
    xmlhttp.send "" 
    GetDRIVERFile= xmlhttp.responseText 
    set xmlhttp = nothing 

End Function

%>
<script src="/scripts/PopBox.js" type="text/javascript"></script>
<script type="text/javascript">
popBoxWaitImage.src = "/images/spinner40.gif" ;
popBoxRevertImage = "/images/magminus.gif";
popBoxPopImage = "/images/magplus.gif";
</script> 
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

   <!-- Begin Wrapper -->
   <div id="wrapper">
   
         <!-- Begin Header -->
         <div id="header_new">
		 <!--#INCLUDE FILE="header.asp"-->
</div>
		 <!-- End Header -->

<table width="100%" border="0">
  <tr>
  <td align="center"> 
  <%  
   Dim objAd
   'Set objAd = Server.CreateObject("MSWC.AdRotator") 'Create the component
   'Response.Write  objAd.GetAdvertisement ("/abanners/adrot.txt") 
   Set objAd = Nothing 'Destroy the object
   %>
</td></tr></table>
 <div id="psubheader">
     <table border="0" cellspacing="0" cellpadding="0" width="980px">
         <tr align="left">
             <td valign="middle" height="27" align="left" width="7%">
                 <a href="default.asp" class="av">&nbsp;Home :</a>
             </td>
             <td valign="middle" height="27" align="left" width="93%">
                 <font color="blue">
                     <%=request("head1")%></font>>> <font color="red">
                         <%=request("head2")%></font>
             </td>
         </tr>
         <tr>
             <td valign="top" colspan="2">
             </td>
         </tr>
     </table>
</div>
	
	  <link rel="stylesheet" type="text/css" href="sdmenu/sdmenu.css" />
	  <script type="text/javascript" src="sdmenu/sdmenu.js">
	</script>
	<script type="text/javascript">
	// <![CDATA[
	var myMenu;
	window.onload = function() {
		myMenu = new SDMenu("my_menu");
		myMenu.oneSmOnly = true
		myMenu.init();
			};
	// ]]>
	</script>
	           <!-- Begin Left Column -->
<div id="leftcolumn">
		  <table width="100%" border="0" cellspacing="0" cellpadding="0" height="20">
	
          <td  valign="top" nowrap width="21%">
			  <%'if rsmain.RecordCount >0 then%>
<div style="float: left" id="my_menu" class="sdmenu">
<%do while (not rsMain.EOF )%>
<div>
<span><%=rsmain("GroupDesc")%></span>
		  <%
          'if rsMain.EOF=false then
          'rsMain.movenext
		  'end if
		  'loop
		  'end if%>
              <%'if rssub.RecordCount > 0 then%>
<%
'SQLtxt2="select * from VSUBCat where CatID='"&Scode&"' order by category"
'catid in (select catm  from item where brshds='BayLan' group by catm)

SQLtxt2="select * from VSUBCat where catcode in (select cat1  from item where brshds='BayLan'  group by cat1) and  CatID='"&rsmain("catid")&"' order by category"  
'&rsmain(0)&

'rsSub.CursorLocation = adUseClient
'rsSub.Open SQLtxt2,conn,adOpenForwardOnly,adLockReadOnly,adcmdtext
'set rsSub.ActiveConnection = nothing
set rssub=conn.Execute(SQLtxt2)%>

               <% do while not rsSub.EOF %>
                <a href="apage2.asp?mData=<%=rsSub("CatCode")%>&mcode=<%=rsSub("CatID")%>&head1=<%=rsmain("groupdesc")%>&head2=<%=rsSub("category")%>"><U><%=rsSub("category")%></U></a>
			  
              <%
          'if rssub.Eof = false then%>
          <%rsSub.movenext
        '  end if
			loop%>
            <%'else%>
          </div>
		      
            <%
 		 rsmain.MoveNext 
		 loop		 
		 %>
      
</div>
		</table>
</div>
<!-- Begin Content Column -->

       <div id="content1">
           
               <% if a="" then  
               if rsdata.RecordCount >0 then
            if xdata <> "" then
               mloop1=rsdata.RecordCount/1
               mloop=int(rsdata.RecordCount/1)
                if mloop1 > mloop then
				mloop3=mloop+1
				else
				mloop3=mloop1
               end if
           else
           mloop1=rsdata.RecordCount/3
           if mloop1 < 1 then
				mloop3=1
			else	           
                mloop3=3
            end if 
           end if  
      
      for i=1 to mloop3
           
             for j=0 to 2
            if rsdata.EOF=false then
                            
            %>
                    <div class="OneProduct" >
                        
                       <table class="OneProductInternalTable">


                    

                       <tr class="PictureRow">
                       <td colspan="1">
                           <img id="<%=trim(rsData("PRODUCTID"))%>" alt="Click Picture Enlarge" onerror="onImgErrorSmall(this)"
                               src="<%=trim(rsData("picturethumb"))%>" pbsrcnl="<%=trim(rsData("picturelarge"))%>"
                               pbcaption="<a href='product.asp?pID=<%=rsdata("productID")%>'>View Product Details</a>"
                               class="PopBoxImageSmall" pbshowrevertbar="true" pbshowreverttext="true" pbshowrevertimage="true"
                               pbreverttext="Click to Shrink" onclick="Pop(this,50,'PopBoxImageLarge');" />
                        
                       </td>
                       <td class="righttalign TopMarginPadding">

                       <table style="width:100%; " >
                       <tr><td style="text-align:right" ><span class="label" > <% response.write right(rsdata("productID"),4)%></span></td></tr>
                       <tr>
                       <tr><td style="height:10px"></td></tr>
                       <td style="text-align:right; vertical-align:bottom;"> <a class="btn btn-mini" href="#">
                                           Prod.ID</a></td>
                       </tr>
                       <tr>
                       <td style="text-align:right;   valign="Top"><b><%=rsData("itemcode")%></b></td>
                       </tr>
                       </table>

                        
                           
                      
                       </td>
                       
                       
                       </tr>

                     
                       <tr>
                       <td colspan="2" class="centeralign DoubleLine GreyBackground">
                            <span class="text-heading rounded ">
                               <%=(rsData("description"))%></span>
                        
                       </td>
                       </tr>
                      

                       <tr>
                           <td colspan="2" class="centeralign" style="padding-left:1px; height:30px; vertical-align:top">
                               <div class="pagination pagination-mini">
                                   <ul>
                                        <li ><a href='#' class="BtnFeature" >Feature</a>
                                            <div style="display: None" class="loader">
                                                
                                                <%
                             
                                                    Dim fso, file, FName
                                                    FName = Server.MapPath(rsdata("longdescription"))
                                                    Set fso = Server.CreateObject ("Scripting.FileSystemObject")
                                                    If fso.FileExists(FName) Then
                                                        Response.Write GetHTMLFile(request("ldesc"))

                                                    end if
                                                
                                                   


                                                %>

                                            </div>
                                        </li>
                                       <%
                                            dim datasheet, dfname
                                            dfname = Server.MapPath(rsdata("datasheet"))
                                            Set datasheet = CreateObject("Scripting.FileSystemObject") 
                                            If datasheet.FileExists(dfname) Then
                                       %>
                                       <li><a href="<%=rsData("datasheet")%>">DataSheet</a></li>
                                       <%
                                       end if 
                                       %>
                                       <%
                                            dim  mname
                                            mname = Server.MapPath(rsdata("manual"))
                                            Set mfmanual = CreateObject("Scripting.FileSystemObject") 
                                            If mfmanual.FileExists(mname) Then
                                       %>
                                      <!--<li><a href = "<%=rsData("manual")%>">Manual</a></li>-->
                                       <%
                                       end if 
                                       %>
                                       <%
                                       DriverPresent="0"
                                       ProductId=rsdata("ProductId")
                                        
                                           'response.Write(Server.MapPath("/Baylan Download Center") )

                                       ListFolderContents(Server.MapPath("/Baylan Download Center") ) 

                                       If DriverPresent<>0 Then
                                       %>

                                       <li><a href="#" class="BtnAllDrivers" >
                                       Download
                                       
                                       </a>
                                            <div style="display: None" class="loader"> 
                                                  
                                                  <%                             
                                                   
                                                    FName = Server.MapPath("driverdetail.asp")
                                                    
                                                    If fso.FileExists(FName) Then
                                                        Response.Write GetDRIVERFile(FName,ProductId,rsData("description"),rsData("itemcode"))
                                                    end if
                                                
                                                %>
                                            </div>
                                       
                                       </li>

                                       <%
                                       end if 
                                       %>
                                        <%
                                            dim  support
                                            support = Server.MapPath(rsdata("support"))
                                            Set fsupport = CreateObject("Scripting.FileSystemObject") 
                                            If fsupport.FileExists(support) Then
                                       %>
                                       <!--<li><a href="<%=(rsdata("support"))%>" >Support</a></li>-->
                                       <%
                                       end if 
                                       %>
                                       <li><a href='#' class="BtnPrint" onclick=openFFPromotionPopup("pID=<%=rsdata("productID")%>")>Print</a></li>                 
                    
                                   </ul>
                               </div>
                           </td>
                       </tr>

                       
                        
                       </table>

                      
                    </div>

               <%
               rsdata.MoveNext
                end if 
            next
            %>

               <br /><br /><br />

           <%
       
       
 next
 end if
end if %>
          
           
       </div>
 
		 
<!-- Begin Footer -->

                      <div id="footer">
 <!--#INCLUDE FILE="footer.asp"-->
 
 

                          <br />
                          <br />
 
 

                     </div>
		 <!-- End Footer -->
		 
   </div>
   <!-- End Wrapper -->
</body>
</html>