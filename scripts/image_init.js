var myTimer = null;
var myTimer2 = null;
var myTimer3 = null;
var preactiveH;
var preactiveB1;
var preactiveB2;
var showI;
var preshowI;
var newsItemA;
var seed;
var preSeed;

function selectFile(sURL) {
	if (sURL != "") location = sURL;
}

var param = "";
function toggleVisibility() {
  var inc, endInc=arguments.length;
  for (inc=0; inc<endInc; inc+=2) {
    var id = arguments[inc];
    if (arguments[inc+1] == 'hidden')param="hidden";
    else if(arguments[inc+1]=='visible')param="visible";
    if (document.getElementById) {
      eval("document.getElementById(id).style.visibility = \"" + param + "\"");
    } else {
      if (document.layers) document.layers[id].visibility = param;
      else if (document.all) eval("document.all." + id + ".style.visibility = \"" + param + "\"");
    }
  }
}

function randomNum(max) {
  var rNum=NaN
  while (isNaN(rNum)) {
    rNum=Math.floor(Math.random()*(max))
  }
  return rNum;
}

function changeImages() {
  for (var i=0; i<changeImages.arguments.length; i+=2) document[changeImages.arguments[i]].src = changeImages.arguments[i+1];
}

function newImage(arg) {
  if (document.images) {
    rslt = new Image();
    rslt.src = arg;
    return rslt;
  }
}

function preloadImages() {
  if (document.images) {
    var imgFiles = preloadImages.arguments;
    for (var i=0; i<imgFiles.length; i+=2){
      eval(arguments[i]+' = new Image()');
      eval(arguments[i]+'.src = "'+arguments[i+1]+'"');
    }
  }
}

preloadImages('arrow2',imgurl+'images/doublearrow.gif');

// headlines
function headlineMover(param) {
  clearTimeout(myTimer);
  activeH = param;
  if (preactiveH) changeImages(preactiveH,imgurl+'images/spacer.gif');
  changeImages(activeH,imgurl+'images/doublearrow_white.gif');  
  preactiveH = activeH;
}

function headlineMout() {
  changeImages(activeH,imgurl+'images/spacer.gif');
  myTimer = setTimeout("headlineMover('spacer1');",200);
}

// bottom nav 1
function bottomNav1Mover(param) {
  clearTimeout(myTimer2);
  activeB1 = param;
  if (preactiveB1) changeImages(preactiveB1,imgurl+'images/spacer.gif');
  changeImages(activeB1,imgurl+'images/doublearrow.gif');  
  preactiveB1 = activeB1;
}

function bottomNav1Mout() {
  changeImages(activeB1,imgurl+'images/spacer2.gif');
  myTimer2 = setTimeout("bottomNav1Mover('b1spacer1');",200);
}

// bottom nav 2
function bottomNav2Mover(param) {
  clearTimeout(myTimer3);
  activeB2 = param;
  if (preactiveB2) changeImages(preactiveB2,imgurl+'images/spacer.gif');
  changeImages(activeB2,imgurl+'images/doublearrow.gif');  
  preactiveB2 = activeB2;
}

function bottomNav2Mout() {
  changeImages(activeB2,imgurl+'images/spacer.gif');
  myTimer3 = setTimeout("bottomNav2Mover('b2spacer1');",200);
}

// grid
function gridOn(param) {
	showI = param;
	if (preshowI) toggleVisibility(preshowI,'hidden',showI,'visible');
	else toggleVisibility(showI,'visible');
	preshowI = showI;
}

function gridOff() {
	if (preshowI) toggleVisibility(preshowI,'hidden');
}

function startNews() {
    toggleVisibility('NewsTicker0Item'+seed,'visible');
    preSeed = seed;
    if (seed == newsItemA) seed = 0;
    else seed = seed + 1;
    setTimeout('rotateNews()',5000);
}

function rotateNews() {
  toggleVisibility('NewsTicker0Item'+preSeed,'hidden','NewsTicker0Item'+seed,'visible');
  startNews();
}
