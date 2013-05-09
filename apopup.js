function popdisclaimer() 
{
if (navigator.appName == "Netscape")
{
if (navigator.appVersion.charAt(0) < 4) 
 {
        window.open('dLogin.htm','propop','screenX='+ x1 + ',screenY=' + y1 + ',width=350,height=240,resizable=yes,toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no');
} 
else
{
var x1 = this.screenX + (this.outerWidth  - 385)/2;
var y1 = this.screenY + (this.outerHeight - 260)/2;
						window.open('dLogin.htm','propop','screenX='+ x1 + ',screenY=' + y1 + ',width=385,height=260,resizable=no,toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no');
}
}
if (navigator.appName == "Microsoft Internet Explorer")					
{
var x1 =  (screen.width  - 385)/2;
var y1 =  (screen.height - 260)/2;
window.open('dLogin.htm','propop','screenX='+ x1 + ',screenY=' + y1 + ',left='+ x1 + ',top='+ y1 + ',width=400,height=250,resizable=yes,toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no');
}	

}