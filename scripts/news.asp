<%@ Language=VBScript %>
<%
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3

set conn = server.CreateObject ("ADODB.Connection")
set rsNews = server.CreateObject("ADODB.Recordset")

function FixQuote(strSource)
	FixQuote = Replace(strSource, "'", "\")
end function

Dim strFile 
strFile = Server.MapPath("buzz.js")
'strFile ="../script/buzz_content.js"
	Dim objFSO 
	Dim objFile 
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	Set objFile = objFSO.OpenTextFile(strFile, 2, True)

%>
<!--#INCLUDE FILE="../dbconnect.asp"-->
<%
conn.Open con 

SQLtext="select * from anews order by newsID" 
rsNews.CursorLocation = adUseClient
rsNews.Open SQLtext,conn,adOpenForwardOnly,adLockReadOnly,adcmdtext
set rsNews.ActiveConnection = nothing
a=0
do while not rsnews.EOF
mnews=rsnews("news")

objFile.Write "var i"&a 
objFile.Write " = '<table border=0 cellspacing=0 cellpadding=0 width=160 ><tr><td valign=top><img src=aimages/doublearrow.gif width=10 height=6  border=0 hspace=5 vspace=3></td>"
objFile.Write "<td><a class=newshead href=# onmouseover=javascript:ticker.stop() onmouseout=javascript:ticker.start() target=_blank>"
objFile.Write mnews
objFile.Write "</a></td></tr></table>';"
objFile.writeline
rsnews.MoveNext
a=a+1
loop
	
	
	' Close the file and dispose of our objects
	objFile.Close
	Set objFile = Nothing
	Set objFSO = Nothing
%>
<html><body></body>
<br><br><br>
<div align="center">
<img src="../aimages/spacer.gif" width=500 height=1><br><br><br>
<p><font face="arial" size=2> <b>NEWS UPDATED !!</b></font></p><br>
<img src="../aimages/spacer.gif" width=500 height=1></div>
 </HTML>