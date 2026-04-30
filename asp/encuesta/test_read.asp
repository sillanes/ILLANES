<%
Response.Write "<pre>"
Response.Write "METHOD: " & Request.ServerVariables("REQUEST_METHOD") & vbCrLf
Response.Write "CONTENT_TYPE: " & Request.ServerVariables("CONTENT_TYPE") & vbCrLf
Response.Write "TOTAL_BYTES: " & Request.TotalBytes & vbCrLf
Response.Write "</pre>"
%>
