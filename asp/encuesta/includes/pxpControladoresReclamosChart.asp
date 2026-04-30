<%
'Response.write(PeriodoID)
sSQL = "EXEC  [pxp].[usp_Controladores_Periodo_sel] " & PeriodoID  & "," &  EmpleadoID  & ", 1"
'Response.write(sSQL)
Set dbRS=Server.CreateObject("ADODB.Recordset")
dbRS.open sSQL, dbCon
IsParentInactive = False
IF ERR.NUMBER <> 0 THEN
	RESPONSE.CLEAR
	response.redirect "./error.asp"
End If
 
FALTANTE_DE_MERCADERIA = dbRS("FALTANTE_DE_MERCADERIA")
ERROR_UNIDAD_DE_MEDIDA = dbRS("ERROR_UNIDAD_DE_MEDIDA")
MERCADERIA_EQUIVOCADA = dbRS("MERCADERIA_EQUIVOCADA")
MERCADERIA_DANADA = dbRS("MERCADERIA_DANADA")
DEVOLUCION_DE_MERCADERIA = dbRS("DEVOLUCION_DE_MERCADERIA")
OTRO = dbRS("OTRO") 
TotalRC = dbRS("TOTALRC") 

Set objRequest = Nothing

%>