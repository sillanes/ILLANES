<%
' Ejemplo de llamada a la API de OpenAI desde Classic ASP
' Requiere que el servidor permita TLS/HTTPS y MSXML2.ServerXMLHTTP

Dim apiKey, requestBody, http, url, responseText
apiKey = "YOUR_API_KEY"
url = "https://api.openai.com/v1/chat/completions"

requestBody = "{""model"": ""gpt-4o-mini"", ""messages"": [{""role"": ""user"", ""content"": ""Hola, quiero integrar ChatGPT Cloud en mi proyecto ASP.""}]}"

Set http = Server.CreateObject("MSXML2.ServerXMLHTTP.6.0")
http.open "POST", url, False
http.setRequestHeader "Content-Type", "application/json"
http.setRequestHeader "Authorization", "Bearer " & apiKey
http.send requestBody

responseText = http.responseText
Response.ContentType = "application/json"
Response.Write responseText

Set http = Nothing
%>
