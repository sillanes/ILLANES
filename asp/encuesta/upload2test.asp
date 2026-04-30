<!-- #include file="uploadhelper.asp" -->
<!DOCTYPE html>
<%
ID = Request.QueryString("ID")&""
FID = Request.QueryString("FID")&""
<!--#include file="./Upload.asp" --><%

Dim Uploader, File
Set Uploader = GetASPUploader
%>



<html lang="es">
  <head>
	<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
    <meta charset="utf-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title></title>
	<script type="text/javascript"language="j78avascript"></script>	
  </head>
  <body>
	<form name="FR" method="post" action="whole_rate_import.asp" enctype="multipart/form-data">
	<input type="hidden" name="doWhat" value="">
	<input type="hidden" name="Reseller" value="<%=Reseller%>">
	<input type="hidden" name="Company" value="<%=Company%>">
	<input type="hidden" name="OrjFileName" value="">
	<tr>
		<td height="20" align="right" class="td13"><b>Upload Name :</b>&nbsp;</td>
		<td class="td20"><input name="UploadName" type="text" class="" size="27" maxlength="20" value=""></td>
	</tr>
	<tr>
		<td height="20" align="right" class="td13"><b>Notes :</b>&nbsp;</td>
		<td class=""><textarea cols="27" rows="2" name="Notes" class=""></textarea></td>
	</tr>
	<tr>
		<td height="20" align="right" class="td13"><b>Rate File :</b>&nbsp;</td>
		<td class="td20"><font size="1"><a href="whole_rate_file.csv"><b>[Sample Rate File]</b></a></font><br><input type="file" class="btn1" name="RateFile" size="26"></td>
	</tr>

	<tr valign="bottom">
		<td colspan="2" align="right"><input type="button" class="btn1" value="Download Current Rates" onclick="document.FEXCEL3.submit()" <%If CLng("0"&Company)=0 Or CLng("0"&Reseller)=0 Then%>disabled<%End If%> style="width:175px;">
		<input type="button" class="btn1" value="Upload Rates " onclick="chk_Form(document.FR,2)" <%If CLng("0"&Company)=0 Or CLng("0"&Reseller)=0 Then%>disabled<%End If%> ></td>
	</tr>
	</table>
	</form>
<script language="javascript">
<!-- Begin
function ShowProgressBar(Fm) {
	var ID = (new Date()).getTime() % 1000000000;
	Fm.action = Fm.action + "?ID=" + ID + "&FID=<%=FID%>";
	var URL = "Progress.asp?ID=" + ID;
	var Ver = navigator.appVersion;
	var I = Ver.indexOf("MSIE");
	if (I > -1 && Ver.charAt(I + 5) > 4) // IE 5 or later
		window.showModelessDialog(URL, null, "dialogWidth=320px; dialogHeight:170px; help:no; status:no; resizable:yes; scroll:no");
	else
		window.open(URL, "_blank", "left=240,top=240,width=320,height=150,location=no,menubar=no,resizable=yes,scrollbars=no,status=no,toolbar=no,directories=no");
}

function GetFileExtension(Filename) {
	var I = Filename.lastIndexOf(".");
	return (I > -1) ? Filename.substring(I + 1, Filename.length).toLowerCase() : "";
}

function GetFileName(Filename) {
	var I = Filename.lastIndexOf("\\");
	return (I > -1) ? Filename.substring(I + 1, Filename.length).toLowerCase() : "";
}

function chk_Form(Fm,prm){
spc=new RegExp(" ","g");
	msg = ""
	if(prm == 2){
		if(Fm.Reseller.value == "")
			msg = "Please select a Reseller!";
		else
		if(Fm.Company.value == "")
			msg = "Please select a Company!";
		else
		if(Fm.RateFile.value == "")
			msg = "Please select a Rate File!";
		else
		if(GetFileExtension(Fm.RateFile.value) != "csv")
			msg = "Invalid file type for Rate File!";
	}

	if(msg != "") alert(msg);
	else{
		Fm.OrjFileName.value = GetFileName(Fm.RateFile.value);
		Fm.doWhat.value = prm;
		ShowProgressBar(Fm);
		Fm.submit();
	}
}

function showHistory(theURL,winName,features) {
	window.open(theURL,winName,features);
}

function onM(row, prm) {
	if(prm == 1 && row.bgColor != "#ff0000")
		winSrc.mOvr(row,"#00FFFF","hand");
	else
	if(prm == 2 || prm == 4){
		winSrc.mOut(row,"#cccccc");
	}
	else
	if(prm == 3){
		winSrc.mOvr(row,"#ff0000","default");
	}
}

function onClk(row, url, prm) {
	if(prm == 4 && row.bgColor != "#ff0000"){
		if (confirm("Are you sure you want to apply these rates!")) { self.location = url; }
		//showHistory(url,"history","scrollbars=yes,width=450,height=350");
	}
	if(prm == 5){
		if (confirm("Are you sure you want to delete this upload!")) { self.location = url; }
	}
}

if(document.FF.Reseller) winSrc.setSelectedIndex(document.FF.Reseller,'<%=Reseller%>');
if(document.FF.Company) winSrc.setSelectedIndex(document.FF.Company,'<%=Company%>');
// End -->
</script>
	
	</body>
</html>
	
	
