<%@ Language=VBScript %>
<%'Option Explicit%>
<!--#include file="../includes/gen_func.asp" -->
<%CheckSession "RM_TOP_DESTINATIONS_ASR","",""%>
<%
UserCompanyID = Session("UserCompanyID")
Dim dbCon, dbRS, cM, sSQL, comboList
%><!--#include file="../includes/db_command_const.asp" --><%
Server.ScriptTimeout = 300
Set dbCon = Server.CreateObject("ADODB.Connection")
%><!--#include file="../includes/db_con_open_report.asp" --><%

If Request.QueryString.Count>0 Then
	Set objRequest = Request.QueryString
Else
	Set objRequest = Request.Form
End If

doWhat = CLng("0" & objRequest("doWhat"))
procedureName = objRequest("procedureName")
databaseName = objRequest("databaseName")



StartDate = Trim(objRequest("StartDate")&"")
EndDate = Trim(objRequest("EndDate")&"")

pgTitle = "Date Range"
Set objRequest = Nothing

Set cM = Server.CreateObject("ADODB.Command")
cM.ActiveConnection = dbCon
cM.CommandText = "dbo.usp_ArchiveDateRanges_sel"
cM.CommandType = adCmdStoredProcedure
cM.Parameters.Append cM.CreateParameter("RETURN_VALUE",adParamInt,adParamReturnValue,0)
cM.Parameters.Append cM.CreateParameter("DatabaseName",adParamVarchar,1,28,databaseName)
cM.Parameters.Append cM.CreateParameter("ProcedureName",adParamVarchar,1,28,procedureName)
cM.Parameters.Append cM.CreateParameter("LiveLimit",adParamInt,3,,0)

cM.Execute
ReturnValue = CLng(cM("RETURN_VALUE"))
liveLimit= CLng("0"&cM("LiveLimit"))

Set cM = Nothing

Set dbRS = Nothing
%>
<HTML>
<HEAD>
<TITLE><%=pgTitle%></TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" type="text/css" href="../includes/style.css">
<link rel="stylesheet" type="text/css" href="../includes/css/new-style.css">
<script language="JavaScript" src="../includes/common.js"></script>

<style>
.daterangepicker .calendar-table th, .daterangepicker .calendar-table td {
    white-space: nowrap;
    text-align: center;
    vertical-align: middle;
    min-width: 32px;
    width: 32px;
    height: 12 !important;
    line-height: 24px;
    font-size: 11px !important;
    border-radius: 4px;
    border: 1px solid transparent;
    white-space: nowrap;
    cursor: pointer;
}
</style>


<script type="text/javascript" src="https://cdn.jsdelivr.net/jquery/latest/jquery.min.js"></script>
<script type="text/javascript" src="https://cdn.jsdelivr.net/momentjs/latest/moment.min.js"></script>
<script type="text/javascript" src="https://cdn.jsdelivr.net/npm/daterangepicker/daterangepicker.min.js"></script>
<link rel="stylesheet" type="text/css" href="https://cdn.jsdelivr.net/npm/daterangepicker/daterangepicker.css" />


</HEAD>
<BODY leftmargin="0" topmargin="0">

<center class="pagetitle"><%=pgTitle%></center>

<form name="FF" method="post" action="calendar_popup.asp">
<input type="hidden" name="doWhat" value="<%=doWhat%>">
<input type="hidden" name="StartDate" value="<%=StartDate%>">
<input type="hidden" name="EndDate" value="<%=EndDate%>">

<table class="tableForm" align="center" border="0" cellpadding="0" cellspacing="0">
<tr>
	<td style="width: 250px;">
		<input type="Text" name="daterange" id="daterange">
	</td>
</tr>
<tr  ><td colspan="1"></td></tr>
</form>
</table>
<br><center><a href="javascript:self.close();" class="popnav" title="Close Window"><b>[ Close ]</b></a></center><br>

<script language="javascript">
<!-- Begin


let picker = $('#daterange').daterangepicker(
{
	"applyButtonClasses": "btn1",
	"cancelClass": "btn1",
	"opens": "center"
}
);
$('#daterange').on('apply.daterangepicker', function(ev, picker) {
  console.log(picker.startDate.format('YYYY-MM-DD'));
  console.log(picker.endDate.format('YYYY-MM-DD'));


  gotoWin(picker.startDate.format('MM-DD-YYYY'), picker.endDate.format('MM-DD-YYYY'));

});

document.FF.daterange.focus();

function gotoWin(prm1,prm2){
	//self.opener.document.getElementById("myDiv").style.display = "none";
	//self.opener.document.getElementById("myDiv00").style.display = "block";

	self.opener.document.FF.StartDateEndDateRange.value = prm1 + "," + prm2;
	self.opener.document.FF.doWhat.value = "0";
	self.opener.document.FF.submit();
	self.opener.focus();
	self.close();
}

function chk_Frm(Fm,pgv){
	Fm.pg.value=pgv;
	Fm.submit();
}

window.setTimeout("self.close();",180000);
// End -->
</script>
</BODY>
</HTML>
<%dbCon.Close
Set dbCon = nothing%>