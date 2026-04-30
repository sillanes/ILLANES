<%
'ARCHIVING TICKET 20191028
Set cM = Server.CreateObject("ADODB.Command")
cM.ActiveConnection = dbCon
cM.CommandText = "dbo.usp_ArchiveDateRanges_sel"
cM.CommandType = adCmdStoredProcedure
cM.Parameters.Append cM.CreateParameter("RETURN_VALUE",adParamInt,adParamReturnValue,0)
cM.Parameters.Append cM.CreateParameter("DatabaseName",adParamVarchar,1,28,ReportProcedureDatabase)
cM.Parameters.Append cM.CreateParameter("ProcedureName",adParamVarchar,1,100,ReportProcedureName)
cM.Parameters.Append cM.CreateParameter("LiveLimit",adParamInt,3,,0)

cM.Execute
ReturnValue = CLng(cM("RETURN_VALUE"))
liveLimit= CLng("0"&cM("LiveLimit"))
Dim dblVbEpoch
dblVbEpoch = CDbl(DateAdd("s", liveLimit, #1970/1/1#))

liveLimitDate = CDate(dblVbEpoch)
%>