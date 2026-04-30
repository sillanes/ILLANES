<script TYPE="text/javascript">
<!--//
OffsetX=-100; // MODIFY THESE VALUES TO
OffsetY=18;   // CHANGE THE POSITION.
var skn=(document.all);
var yyy=-1000;

var ns4=document.layers;
var ns6=document.getElementById&&!document.all;
var ie4=document.all;
//var ffx=navigator.userAgent.toLowerCase().indexOf('firefox') > -1;

if (ns4) skn=document.pdqbox;
else if (ns6) skn=document.getElementById("pdqbox").style;
else if (ie4) skn=document.all.pdqbox.style;

if (ns4) document.captureEvents(Event.MOUSEMOVE);
else { skn.visibility="visible"; skn.display="none"; }

document.onmousemove=get_mouse;

function popup(msg){
	var content=msg;
	yyy=OffsetY;
	if(ns4){skn.document.write(content);skn.document.close();skn.visibility="visible"}
	if(ns6){document.getElementById("pdqbox").innerHTML=content;skn.display=''}
	if(ie4){document.all("pdqbox").innerHTML=content;skn.display=''}
}

function get_mouse(e){
	var x=(ns4||ns6)?e.pageX:event.x+document.body.scrollLeft;
	var y=(ns4||ns6)?e.pageY:event.y+document.body.scrollTop;
	skn.left=x+OffsetX;
	skn.top=y+yyy;
}

function remove_popup(){
	yyy=-1000;
	if(ns4){ skn.visibility="hidden"; }
	else if (ns6||ie4) { skn.display="none"; }
}
//-->
</script>
