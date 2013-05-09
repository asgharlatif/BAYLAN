<%@ Language=VBScript %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Picture Catelogue.</title>
    <link rel="stylesheet" type="text/css" href="anylinkmenu.css" />

<link rel="stylesheet" type="text/css" href="assets/css/bootstrap.css" />
<link href="acss/newstyle.css" rel="stylesheet" type="text/css" />
<link href="background/printpagestyle.css" rel="Stylesheet" type="text/css" />

 <script language="javascript" type="text/javascript" >
     $(function () {
         $(".BtnFeature").on("click", abc);

     });

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

        function myfunction(parameter1) {
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
</head>
<body onLoad="window.print()">

    
<%
set conn = server.CreateObject ("ADODB.Connection")
set rsMain = server.CreateObject("ADODB.Recordset")
set rsSub = server.CreateObject("ADODB.Recordset")
set rsData = server.CreateObject("ADODB.Recordset")
set rsDriver = server.CreateObject("ADODB.Recordset")

%>
<!--#INCLUDE FILE="dbconnect.asp"-->
<%
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3

CatRange=request.QueryString("CatRange")

'on error resume next

conn.Open con 
'catcode in (select cat1  from item where brshds='BayLan'  group by cat1)
SQLtxt2="select * from VSUBCat where catcode in (select cat1  from item where brshds='BayLan'  group by cat1) and CatID in ("& CatRange &") order by category"  
set rssub=conn.Execute(SQLtxt2)

do while not rsSub.EOF
   

baylenProduct="  productid in (select itcd from item where brshds='BayLan' group by itcd) "
SQLtxt3="select productid,item,image,itemcode,brand,picturethumb,picturelarge,datasheet,warranty,driver,description,longdescription,retail,pga,statuscode,support,manual from vAllData where "& baylenProduct &" and Catid in ("& CatRange &") and catcode='"& rssub("catcode") &"' and statuscode != '08' order by sorder,brand,description" 


	rsData.CursorLocation = adUseClient
	rsData.Open SQLtxt3,conn,adOpenForwardOnly,adLockReadOnly,adcmdtext
	set rsData.ActiveConnection = nothing
               mloop1=rsdata.RecordCount/3
               mloop=int(rsdata.RecordCount/3)
               if mloop1 > mloop then
				mloop3=mloop+1
              end if

%>

<table class="PrintPage">
<tr>
<td class="PrintPageTd1">

</td>


<td class="td2">




<table>

<tr>
    <td >
        <!--
        <img src="images/Baylan-Banner.jpg" style="width:750px" />
        -->
        <img alt="" src="Background/BCPT1.jpg" style="width:750px" />


    </td>
</tr>

<tr>
<td>

<h5><%=request.querystring("textn") %> >> <%=rsSub("category") %></h5>
</td>

</tr>

<tr >

<td>
    <div id="content3">
           
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


                    

                       <tr>
                       <td colspan="1">
                           <img id="<%=trim(rsData("PRODUCTID"))%>" alt="Click Picture Enlarge" onerror="onImgErrorSmall(this)"
                               src="<%=trim(rsData("picturethumb"))%>" pbsrcnl="<%=trim(rsData("picturelarge"))%>"
                               pbcaption="<a href='product.asp?pID=<%=rsdata("productID")%>'>View Product Details</a>"
                               class="PopBoxImageSmall" pbshowrevertbar="true" pbshowreverttext="true" pbshowrevertimage="true"
                               pbreverttext="Click to Shrink" onclick="Pop(this,50,'PopBoxImageLarge');" />
                        
                       </td>
                       <td class="righttalign TopMarginPadding">

                       <table style="width:100%" >
                       <tr><td style="text-align:right" ><span class="label" > <% response.write right(rsdata("productID"),4)%></span></td></tr>
                       <tr>
                       <tr><td style="height:10px"></td></tr>
                       <td style="text-align:right; vertical-align:bottom;"> <a class="btn btn-mini" href="#">
                                           Prod.ID</a></td>
                       </tr>
                       <tr>
                       <td style="text-align:right; vertical-align:top;"><b><%=rsData("itemcode")%></b></td>
                       </tr>
                       </table>

                        
                           
                      
                       </td>
                       
                       
                       </tr>

                     
                       <tr>
                       <td colspan="2" class="centeralign DoubleLinePrint GreyBackground"  >
                            <span class="text-heading rounded "  >
                               <%=(rsData("description"))%></span>
                        
                       </td>
                       </tr>
                      

                        

                       
                        
                       </table>

                      
                    </div>


               <%
               rsdata.MoveNext
                end if 
            next
           
       
       
 next
 end if
end if %>
          
           
       </div>

</td>

</tr>
<tr>
<td>
    <table style="width:100%" class="SetBackgroung" >
    <tr>
        <td>http://wwww.baylan.com</td>
        <td style="text-align:center">tw-sales@baylan.com</td>
        <td style="text-align:right"><%= FormatDateTime(Date, 1) %>&nbsp;<%= FormatDateTime(Now, 4) %></td>
    </tr>
    </table>
    

</td>
</tr>

</table>

<p class="breakhere" />


</td>

<td class="PrintPageTd2"></td>

</tr>
</table>

<%

rsData.close
rsSub.movenext
loop
 %>


</body>
</html>
