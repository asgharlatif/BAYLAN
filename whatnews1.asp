<html><head>
<title>News</title>
<style type="text/css">
BODY 		{ margin-top: 0px; margin-right: 0px; margin-bottom: 0px; margin-left: 0px; }
 /* FONT COLORS */
TABLE		{ COLOR: #000000; FONT: 11px arial, sans-serif; font-weight: normal }
.title		{ COLOR: #0033FF; FONT: 12px arial, sans-serif; font-weight: bold; }
#NewsDiv	{ position: absolute; left: 0; top: 0; width: 100% }
/* PAGE LINK COLORS */
a:link		{ color: #0033FF; text-decoration: underline; }
a:visited	{ color: #6633FF; text-decoration: underline; }
a:active	{ color: #0033FF; text-decoration: underline; }
a:hover		{ color: #6699FF; text-decoration: none; }
.style1 {color: #FF0000}
</style>
</head>
<BODY BGCOLOR="#F9F799" TEXT="#000000" onMouseover="scrollspeed=0" onMouseout="scrollspeed=current" OnLoad="NewsScrollStart();">
<div id="NewsDiv">
<table cellpadding="0" cellspacing="0" border="0" width="100%"><tr><td>
  <p>
    <!-- SCROLLER CONTENT STARTS HERE -->
    
    
         <span class="title">Price Revised<br>
		</span> TSC Barcode LabelPrinter<BR>
		</SPAN> USB Thermal/Direct<br>
        <br>PC Terminal</p>
		
    
    
    
          <!-- SCROLLER CONTENT ENDS HERE -->
    
    
    
    
    
    
    </p></td></tr></table>
</div>




<!-- YOU DO NOT NEED TO EDIT BELOW THIS LINE -->




<script language="JavaScript" type="text/javascript">
<!-- HIDE CODE

var scrollspeed		= "1"		// SET SCROLLER SPEED 1 = SLOWEST
var speedjump		= "30"		// ADJUST SCROLL JUMPING = RANGE 20 TO 40
var startdelay 		= "2" 		// START SCROLLING DELAY IN SECONDS
var nextdelay		= "0" 		// SECOND SCROLL DELAY IN SECONDS 0 = QUICKEST
var topspace		= "2px"		// TOP SPACING FIRST TIME SCROLLING
var frameheight		= "200px"	// IF YOU RESIZE THE WINDOW EDIT THIS HEIGHT TO MATCH



current = (scrollspeed)


function HeightData(){
AreaHeight=dataobj.offsetHeight
if (AreaHeight==0){
setTimeout("HeightData()",( startdelay * 1000 ))
}
else {
ScrollNewsDiv()
}}

function NewsScrollStart(){
dataobj=document.all? document.all.NewsDiv : document.getElementById("NewsDiv")
dataobj.style.top=topspace
setTimeout("HeightData()",( startdelay * 1000 ))
}

function ScrollNewsDiv(){
dataobj.style.top=parseInt(dataobj.style.top)-(scrollspeed)
if (parseInt(dataobj.style.top)<AreaHeight*(-1)) {
dataobj.style.top=frameheight
setTimeout("ScrollNewsDiv()",( nextdelay * 1000 ))
}
else {
setTimeout("ScrollNewsDiv()",speedjump)
}}



// END HIDE CODE -->
</script>


</body>
</html>
