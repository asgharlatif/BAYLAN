<%@ Language=VBScript %>
<link href="acss/newstyle.css" rel="stylesheet" />
<link href="assets/css/apnicss.css" rel="stylesheet" />
<link href="assets/css/bootstrap.css" rel="stylesheet" />
<script src="JQuery/jquery-1.8.3.min.js"></script>
<script src="JsFiles/underscore.js"></script>
<script src="JsFiles/backbone.js"></script>

<script id="Email-Template" type="text/html">
        
        <spin id="spin1" >Email:-</spin>         
        <input id="txt-email"  />
        
</script>

<script language="javascript" type="text/javascript" >
    $(function () {

        // $(alert("This is source control test"));

    });
</script>

<div class="FileListening">
    &nbsp;<h2>Baylan Download Center</h2><div id="Send-Page_Button"></div>

<br />

<div id="email-content-placing">
    
</div>

<br />


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

        Response.Write("<li><table class='filelisteningtable'><tr><td class='filelisteningtabletd1'>File Name</td><td class='filelisteningtabletd2'>Last Updated</td><td class='filelisteningtabletd3'>File Size</td></tr><tr><td class='filelisteningtabletd1'><a href=""" & url & """>" & item.Name & "</a></td><td class='filelisteningtabletd2'>"& item.DateCreated &"</td><td class='filelisteningtabletd3'>"&   round(item.Size/1024,0) & " KB "  &"</td><tr></table></li>")

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

end function %>



<script src="JsFiles/SendEmail.Js"></script>
