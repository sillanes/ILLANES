var strColumns_Current = "*,260";
function hidetoc(){
	strColumns_Current = top.fstMain.cols
	top.fstMain.cols = "*,1";
	top.fraMiddleBar.document.all("showtoc").style.display = "block";
}
function showtoc(){
	top.fstMain.cols = strColumns_Current;
	top.fraMiddleBar.document.all("showtoc").style.display = "none";
}

function mouseover(item){
	switch (item){
		case "hidetoc" :
		window.status = "Hide TOC";
		document.all.imgHideToc.src = "images/hidetoc2.gif"
		break;

		case "showtoc" :
		window.status = "Show TOC";
		document.all.imgShowToc.src = "images/showtoc2.gif"
		break;
	}
}
function mouseout(item){
	switch (item){
		case "hidetoc" :
		window.status = "";
		document.all.imgHideToc.src = "images/hidetoc1.gif"
		break;

		case "showtoc" :
		window.status = "";
		document.all.imgShowToc.src = "images/showtoc1.gif"
		break;
	}
}

function mOvr(src,clrOver,cursor){
//	if (!src.contains(event.fromElement)){
		src.style.cursor = cursor;
		src.bgColor = clrOver;
//	}
}
function mOut(src,clrIn){
//	if (!src.contains(event.toElement)){
		src.style.cursor = 'default';
		src.bgColor = clrIn;
//	}
} 
function mClk(src){
//	if(event.srcElement.tagName=='TR'){
//		src.children.tags('A')[0].click();
		window.location = src;
//	}
}

function showContent_a(thismenu1,thismenu2){
	thismenu1.style.display="none";
	thismenu2.style.display="block";
}

function setSelectedIndex(obj,objVal){
	for(var i=0; i<obj.options.length; i++){
		if (obj.options[i].value==objVal){
			obj.options[i].selected=true;
			break;
		}
	}
}

function setSelectedIndexMul(obj,objVal){
	for(var i=0; i<obj.options.length; i++)
		if (objVal.indexOf(", "+obj.options[i].value+",") >= 0)
			obj.options[i].selected=true;
}

function trex(obj){ obj.value=""; }

function mywin(prm,hg,wd){
	window.open(prm,"zero","height="+hg+",width="+wd+",status=yes,toolbar=no,directories=no,menubar=no,location=no,resizable=no,scrollbars=yes,left=55,top=25");
}

function mywin2(prm,hg,wd){
	window.open(prm,"zero","height="+hg+",width="+wd+",status=no,toolbar=no,directories=no,menubar=no,location=no,resizable=no,scrollbars=no,left=250,top=100");
}

function MM_openBrWindow(theURL,winName,features) {
	pop1 = window.open(theURL,winName,features);
	pop1.focus();
}

function FillCombo(obj1, obj2, Arr1, Arr2){
	cnt = 0;
	obj2.options[cnt] = new Option(obj2.options[0].text,obj2.options[0].value);
	str = obj1.options[obj1.options.selectedIndex].value;
	for(ctr=0;ctr<Arr1.length;ctr++){
		if(Arr1[ctr] == str){
			cnt = cnt + 1;
			pos = Arr2[ctr].indexOf('|')
			obj2.options[cnt] = new Option(Arr2[ctr].substr(0,pos),Arr2[ctr].substr(pos+1));
		}
	}
	obj2.length=cnt+1;
}

function FormatCurrency(obj){
	spc = new RegExp(" ","g");
	objVal = obj.value.replace(spc,"");
	valLen = objVal.length;
	if(valLen > 0 && objVal.indexOf('.') < 0){
		if(valLen == 2)	obj.value =  '0.' + objVal;
		else
		if(valLen == 1)	obj.value =  '0.0' + objVal;
		else
			obj.value = objVal.substr(0,valLen-2) + '.' + objVal.substr(valLen-2);
	}
}

function grabText(obj1, obj2, prm){
	xVal = "";
	if(obj1.options[obj1.options.selectedIndex].value != "")
		xVal = obj1.options[obj1.options.selectedIndex].text;
	if(xVal != "" && prm == 2) { xVal = xVal.substr(xVal.indexOf(' - ')+3); }
	if(xVal != "") obj2.value = xVal;
}

function cancelMe(nextpage){
	if (window.opener) { window.close(); }
	else if (nextpage.length > 0) { window.location = nextpage; }
	else { window.location = '../index.asp'; }
}

function DateAdd(DateStr, AddNumber, DatePart){
	var curDate = new Date(DateStr);
	curDate.setHours(1,0,0,0);

	if(DatePart == "d")	
		var newDate = new Date(curDate.getTime()+(AddNumber*24*60*60*1000));
	else
		var newDate = curDate;

	newDate.setHours(0,0,0,0);
	return	newDate;
}

function MM_findObj(n, d) {
	var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
	d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
	if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
	for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
	if(!x && d.getElementById) x=d.getElementById(n);
	return x;
}

function MM_swapImgRestore(){
  var i,x,a=document.MM_sr; 
  for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_preloadImages(){
var d=document; 
	if(d.images){ 
		if(!d.MM_p) d.MM_p=new Array();
		var i,j=d.MM_p.length,a=MM_preloadImages.arguments; 
		for(i=0; i<a.length; i++)
    	if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}
	}
}

function MM_swapImage(){
	var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; 
	for(i=0;i<(a.length-2);i+=3)
		if ((x=MM_findObj(a[i]))!=null){
			document.MM_sr[j++]=x; 
			if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];
		}
}

var selectedRow = null;
var selectedAll = false;
function mOvrRow(row, highlightColor, cursor){
	if(selectedRow == null || row.id != selectedRow.id){
		if(!selectedAll) row.bgColor = highlightColor;
		row.style.cursor = cursor;
	}
}
function mOutRow(row, rowColor, cursor){
	if(selectedRow == null || row.id != selectedRow.id){
		if(!selectedAll) row.bgColor = rowColor;
		row.style.cursor = cursor;
	}
}
function mClkRow(row, selectColor, rowColor, cursor, clkEvent){
	if(selectedRow != null) selectedRow.bgColor = rowColor;
	if(selectedRow == null || row.id != selectedRow.id){
		if(clkEvent.length > 0) eval(clkEvent);
		selectedAll = false;
		selectedRow = row;
		row.bgColor = selectColor;
		row.style.cursor = cursor;
	}
}

function chk_StartDate(obj){
	var inputDate = new Date();
	inputDate.setTime(Date.parse(obj.value));

	var baseDate = new Date();
	baseDate.setTime(Date.parse("08/01/2006"));

	if(inputDate.valueOf() < baseDate.valueOf()){
		obj.value = "08/01/2006";
		//return !!! don't have any data before August 8th 2003 !!!
	}

	return "";
}

function IsNotChecked(obj){
	if(obj){
		objLen = obj.length;
		if(objLen){
			for(var i=0; i<objLen; i++)
				if(obj[i].checked) return false;
		}
		else
		if(obj.checked) return false;
	}
	return true;
}

function IsNotNumeric(str){
	if (str.length == 0) return false;
	for (k=0;k<str.length;k++){
		if (str.charCodeAt(k)<48 || str.charCodeAt(k)>57)	return true;
	}
	return false;
}

function onAniValidation(ani) {
  const checkSpaces = new RegExp("\\s+");
  const checkType = new RegExp("^[0-9]+$");
    
  if(!ani) {
    return { isValid: false, message: "Please enter a CallerID!"};
  }
    
  if(!checkType.test(ani)) {
    return { isValid: false, message: "Please enter a valid number"};
  }
    
  if(checkSpaces.test(ani)) {
    return { isValid: false, message: "CallerID contains spaces please check"};
  }
    
  if(ani.charAt(0) == 0 && ani.substr(0, 3) !== "011") {
    return { isValid: false, message: "Number not allowed"};
  }
    
  if(ani.length < 8){
    return { isValid: false, message: "CallerID must be min 8 digits!"};
  }
    
  return { isValid: true, message: "ANI is valid"};
}

function isValidAni(ani) {
  const checkSpaces = new RegExp("\\s+");
  const checkType = new RegExp("^[0-9]+$");
    
  if(!ani) return false;

  if(!checkType.test(ani)) return false;
    
  if(checkSpaces.test(ani)) return false;
    
  if(ani.charAt(0) == 0 && ani.substr(0, 3) !== "011") return false;
    
  if(ani.length < 8) return false;
    
  return true;
}

  /*
    inp: input text elemtent
    arr: [{id, label}, {id, label}]
    callback: is called after selecting an item
  */
  function autocomplete(inp, arr, callback) {
    /*the autocomplete function takes two arguments,
    the text field element and an array of possible autocompleted values:*/
    var currentFocus;
    /*execute a function when someone writes in the text field:*/
    inp.addEventListener("input", function(e) {
      var a, b, i, val = this.value;
      /*close any already open lists of autocompleted values*/
      closeAllLists();
      if (!val) { return false;}
      currentFocus = -1;
      /*create a DIV element that will contain the items (values):*/
      a = document.createElement("DIV");
      a.setAttribute("id", this.id + "-" + "autocomplete-list");
      a.setAttribute("class", "autocomplete-items");
      /*append the DIV element as a child of the autocomplete container:*/
      this.parentNode.appendChild(a);
        /*for each item in the array...*/
      for (i = 0; i < arr.length; i++) {
        /*check if the item starts with the same letters as the text field value:*/
        if (arr[i].label.substr(0, val.length).toUpperCase() == val.toUpperCase()) {
          /*create a DIV element for each matching element:*/
          b = document.createElement("DIV");
          /*make the matching letters bold:*/
          b.innerHTML = "<strong>" + arr[i].label.substr(0, val.length) + "</strong>";
          b.innerHTML += arr[i].label.substr(val.length);
          /*insert a input field that will hold the current array item's value:*/
          b.innerHTML += "<input type='hidden' value='" + arr[i].id + "|" + arr[i].label + "'>";
          /*execute a function when someone clicks on the item value (DIV element):*/
          b.addEventListener("click", function(e) {
            /*insert the value for the autocomplete text field:*/
            const value = this.getElementsByTagName("input")[0].value;
            const [id, label] = value.split("|");           
            inp.value = label;
            callback({id, label});
            /*close the list of autocompleted values,
            (or any other open lists of autocompleted values:*/
            closeAllLists();
          });
          a.appendChild(b);
        }
      }
    });
    /*execute a function presses a key on the keyboard:*/
    inp.addEventListener("keydown", function(e) {
      var x = document.getElementById(this.id + "-" + "autocomplete-list");
      if (x) x = x.getElementsByTagName("div");
      if (e.keyCode == 40) {
        /*If the arrow DOWN key is pressed,
        increase the currentFocus variable:*/
        currentFocus++;
        /*and and make the current item more visible:*/
        addActive(x);
      } else if (e.keyCode == 38) { //up
        /*If the arrow UP key is pressed,
        decrease the currentFocus variable:*/
        currentFocus--;
        /*and and make the current item more visible:*/
        addActive(x);
      } else if (e.keyCode == 13) {
        /*If the ENTER key is pressed, prevent the form from being submitted,*/
        e.preventDefault();
        if (currentFocus > -1) {
          /*and simulate a click on the "active" item:*/
          if (x) x[currentFocus].click();
        }
      }
    });
    function addActive(x) {
      /*a function to classify an item as "active":*/
      if (!x) return false;
      /*start by removing the "active" class on all items:*/
      removeActive(x);
      if (currentFocus >= x.length) currentFocus = 0;
      if (currentFocus < 0) currentFocus = (x.length - 1);
      /*add class "autocomplete-active":*/
      x[currentFocus].classList.add("autocomplete-active");
    }
    function removeActive(x) {
      /*a function to remove the "active" class from all autocomplete items:*/
      for (var i = 0; i < x.length; i++) {
        x[i].classList.remove("autocomplete-active");
      }
    }
    function closeAllLists(elmnt) {
      /*close all autocomplete lists in the document,
      except the one passed as an argument:*/
      var x = document.getElementsByClassName("autocomplete-items");
      for (var i = 0; i < x.length; i++) {
        if (elmnt != x[i] && elmnt != inp) {
          x[i].parentNode.removeChild(x[i]);
        }
      }
    }
    /*execute a function when someone clicks in the document:*/
    document.addEventListener("click", function (e) {
      closeAllLists(e.target);
    });
  }
