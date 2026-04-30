function chkForm(Fm){
spc=new RegExp(" ","g");
userMsgFinal = "";
	for(var i=0; i<Fm.elements.length-1; i++){
		obj = Fm.elements[i];

		if (obj.title) objName = obj.title;
		else objName = obj.name;

		if (obj.type == "select-one" || obj.type == "select-multiple"){
			objVal = obj.options[obj.options.selectedIndex].value.replace(spc,"");
			userMsg= "Please select ";
		}
		else if (obj.type == "radio" || obj.type == "checkbox"){
			objVal = (obj.checked) ? (obj.value.replace(spc,"")) : "";
			userMsg= "Please check ";
		}
		else{
			objVal = obj.value.replace(spc,"");
			userMsg= "Please enter ";
		}

		// chktype format : field type|field required|min length
		// myArr[0]=(field type	: str, num, date, money, pw, email, tel, zip, custom)
		// myArr[1]=(field required : 0, 1, true, false)
		// myArr[2]=(min length 
		//			 or custom function name to be executed(if field type is custom)
		//			 or min and max values to be checked(if field type is num)
		if (obj.chktype){
			myArr = obj.chktype.split("|");
			ArLen = myArr.length;

			if (ArLen > 1 && eval(myArr[1]) && objVal == "")
				userMsg = userMsg + objName + "...";
			else if (objVal != "" && myArr[0] == "num" && isNaN(objVal))
				userMsg = "Please do not use non-numeric values for " + objName + "...";
			else if (objVal != "" && myArr[0] == "alphanum" && IsNotAlphaNumeric(objVal))
				userMsg = "Please use only AlphaNumeric values for " + objName + "...";
			else if (objVal != "" && myArr[0] == "date")
				userMsg = chkDate(objVal, objName);
			else if (objVal != "" && myArr[0] == "phone")
				userMsg = chkPhone(objVal, objName);
			else if (objVal != "" && myArr[0] == "email" && notValidEmail(objVal))
				userMsg = "Please enter a valid email address for " + objName + "...";
			else if (objVal != "" && myArr[0] == "num" && ArLen==4 && 
					(Number(objVal)<myArr[2] || Number(objVal)>myArr[3]) )
				userMsg = "Please enter a value between "+myArr[2]+" and "+myArr[3]+" for the "+objName+"...";
			else if (myArr[0] == "custom" && ArLen>2)
				userMsg = eval(myArr[2]);
			else if (objVal != "" && ArLen > 2 && objVal.length < parseInt(myArr[2]))
				userMsg = userMsg + objName + " in complete...";
			else
				userMsg = "";

			if (userMsg != "") {
				if (userMsgFinal == "") firstobj = obj;
				userMsgFinal = userMsgFinal + "- " + userMsg + "\n";
			}
		}
	}//for

	if (userMsgFinal != ""){
		alert(userMsgFinal);
		firstobj.focus();
		return false;
	}
	return true;
}

function IsNotNumeric(str){
	if (str.length == 0) return true;
	for (k=0;k<str.length;k++){
		if (str.charCodeAt(k)<48 || str.charCodeAt(k)>57)	return true;
	}
	return false;
}

function IsNotAlphaNumeric(str){
	if (str.length == 0) return true;
	for (k=0;k<str.length;k++){
		chr = str.charCodeAt(k);
		if (! ((chr>=48 && chr<=57) || (chr>=65 && chr<=90) || (chr>=97 && chr<=122) || (chr==95)) ) return true;
	}
	return false;
}

function chkDate(dateStr, objName){
	if(dateStr.length == 0)	return "Please enter a valid date for " + objName + "...";

	if (dateStr.indexOf("/",dateStr.indexOf("/")+1) > 1)		seperator = "/";
	else if (dateStr.indexOf(".",dateStr.indexOf(".")+1) > 1)	seperator = ".";
	else return "Please use one of the '/' or '.' seperators for " + objName + "...";
		
	dateArr = dateStr.split(seperator);
	dateArr[0] = Number(dateArr[0]);
	dateArr[1] = Number(dateArr[1]);
	dateArr[2] = Number(dateArr[2]);

	if (isNaN(dateArr[0]) || isNaN(dateArr[1]) || isNaN(dateArr[2]))
		return "Please do not use non-numeric characters for " + objName + "...";
	if (dateArr[0]<1 || dateArr[0]>12 || dateArr[1]<1)
		return "Please enter a valid date for " + objName + "...";
	if (dateArr[2] < 1000)
		return "Please enter 4 digits for the year part of " + objName + "...";
	if (dateArr[2] < 1900 || dateArr[2] > 2050)
		return "Please enter the year part of " + objName + " between 1900 and 2050...";

	Ar=[0,31,28,31,30,31,30,31,31,30,31,30,31];
	if (dateArr[2] % 4 == 0) Ar[2]=29;
	if (dateArr[1] > Ar[dateArr[0]])
		return "Please enter a valid date for " + objName + "...";

	return "";
}

function notValidEmail(str){
	strLen = str.length;
	if (strLen < 5) 			return true;
	if (str.indexOf(" ") >= 0)	return true;
	if (str.indexOf("..")>= 0)	return true;
	if (str.indexOf(".") < 1) 	return true;
	if (str.indexOf("@") < 1) 	return true;
	if (str.lastIndexOf(".")>strLen-3 || str.lastIndexOf("@")>strLen-3) return true;
	for (i=0;i<strLen;i++){
		chr=str.charCodeAt(i);
		if (chr<=39 || chr>=127 || chr==44 || chr==59)	return true;
	}
	return false;
}

function chkPhone(str, objName){
	str = str.toLowerCase();
	if (str.indexOf("x") != str.lastIndexOf("x"))
		return ("Please enter " + objName + " in correct format...");
	for (k=0;k<str.length;k++){
		chr=str.charCodeAt(k);
		if (!((chr>=48 && chr<=57) || chr==46 || chr==40 || chr==41 || chr==45 || chr==32 || chr==120))
			return ("Please enter " + objName + " in correct format...");
	}
	return "";
}
