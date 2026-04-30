<script TYPE="text/javascript">
<!--//
OffsetX=-100; // MODIFY THESE VALUES TO
OffsetY=18;   // CHANGE THE POSITION.
var skn=(document.all);

var ns4=document.layers;
var ns6=document.getElementById&&!document.all;
var ie4=document.all;
//var ffx=navigator.userAgent.toLowerCase().indexOf('firefox') > -1;

if (ns4) skn=document.pdqbox;
else if (ns6) skn=document.getElementById("pdqbox").style;
else if (ie4) skn=document.all.pdqbox.style;

if (ns4) document.captureEvents(Event.MOUSECLICK);
else { skn.visibility="visible"; skn.display="none"; }

function popup(msg,e){
	var content=msg;
	if(ns4){skn.document.write(content);skn.document.close();skn.visibility="visible"}
	if(ns6){document.getElementById("pdqbox").innerHTML=content;skn.display=''}
	if(ie4){document.all("pdqbox").innerHTML=content;skn.display=''}

	var posx, posy = 0;
	if (!e) var e = window.event;
	if (e.pageX || e.pageY) { posx = e.pageX; posy = e.pageY; }
	else 
	if (e.clientX || e.clientY) { posx = e.clientX + document.body.scrollLeft; posy = e.clientY + document.body.scrollTop; }

	skn.left = posx + OffsetX;
	skn.top  = posy + OffsetY;
}

function remove_popup(){
	if(ns4){ skn.visibility="hidden"; }
	else if (ns6||ie4) { skn.display="none"; }
}
//-->
</script>
