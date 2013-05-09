<%@ Language=VBScript %>
<link href="acss/newstyle.css" rel="stylesheet" />

<link rel="stylesheet" type="text/css" >
<head></head>
<style type="text/css">



.EmailTemplateDriver
{
    background-color: #f3f3f3;
    width: 540px;
    text-align: center;
    vertical-align: middle;
    color: #000080;
}
.EmailTemplateDriver #dname
{
    font-size: 20px;

}
.centeralignRedlabeling
{
    font-family: 'Trebuchet MS' , 'Lucida Sans Unicode' , 'Lucida Grande' , 'Lucida Sans' , Arial, sans-serif;
    font-size: 16px;
    display: block;
    width: 100%;
    text-align: center;
    vertical-align: middle;
    color: #FAA732;
}
</style>

<link href="assets/css/apnicss.css" rel="stylesheet" />
<link href="assets/css/bootstrap.css" rel="stylesheet" />





<script src="JsFiles/underscore.js"></script>
<script src="JsFiles/backbone.js"></script>
<script type="text/javascript" src="JsFiles/handlebars-1.0.rc.1.min.js"></script>
<script type="text/javascript" src="BayTemplate.js"></script>


<script id="Email-Template" type="text/html">
        
        <input type="text" id="IdEmailId" class="search-query span2" placeholder="Enter email to send" style="width:200px" value="">
</script>




<div class="FileListening">
    <table style="width:100%;" ><tr><td><h4>Baylan Download Center</h4></td><td style="text-align:right; padding-right:5px"><div id="Send-Page-Button">
        <a id="Send-btn" href="#" class="btn btn-info icon-bar" >Share it.</a></div></td></tr></table>
        <div id="email-content-placing" style="display:none" ></div>
<br />
<table class='filelisteningtable'><tr><td class='filelisteningtabletd1'>File Name</td><td class='filelisteningtabletd2'>Last Updated</td><td class='filelisteningtabletd3'>File Size</td></tr></table>

<% ListFolderContents(Server.MapPath("/Baylan Download Center")) %>
</div>

<% 
    
    
    sub ListFolderContents(path)

     dim fs, folder, file, item, url
     dim FileCount 
     set fs = CreateObject("Scripting.FileSystemObject")
     set folder = fs.GetFolder(path)

     productid=Request.QueryString("productid")

    'Display the target folder and info.
    FileCount=0

    for each item in folder.Files
         url = MapURL(item.path)
        
         If Instr(item.Name, productid) Then
            FileCount=FileCount+1
         end if
    next

     if (FileCount>0) then
        Response.Write("<h4>"& folder.Name &" ("& FileCount &")</h4>")
     end if

     'Display a list of sub folders.

     for each item in folder.SubFolders
                ListFolderContents(item.Path)
     next

     'Display a list of files.
    Response.Write("<ul>")
    for each item in folder.Files
        url = MapURL(item.path)
        
         If Instr(item.Name, productid) Then

        Response.Write("<li><table class='filelisteningtable'><tr><td class='filelisteningtabletd1'><a href=""" & url & """>" & item.Name & "</a></td><td class='filelisteningtabletd2'>"& item.DateCreated &"</td><td class='filelisteningtabletd3'>"&   round(item.Size/1024,0) & " KB "  &"</td><tr></table></li>")

        end if
    next
        Response.Write("</ul>")
   end sub


   function MapURL(path)

     dim rootPath, url

     'Convert a physical file path to a URL for hypertext links.

     rootPath = Server.MapPath("/")
     url = Right(path, Len(path) - Len(rootPath))
     MapURL = Replace(url, "\", "/")

end function


%>

<input value='<%=Request.QueryString("productid") %>' id="txtproductid" style="display:none" />
<input value='<%=Request.QueryString("description") %>' id="txtdescription" style="display:none" />
<input value='<%=Request.QueryString("itemcode") %>' id="txtitemcode"  style="display:none" />

<script src="JsFiles/SendEmail.Js"></script>



