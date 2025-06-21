<!--#include file="../../v2/RabbitJSON.v2.asp" -->
<html>
	<head>
		<title>Sample 11 - Dictionary Keys ile Loop</title>
		<meta charset="utf-8">
	</head>
	<body>
<%
' Configuration objesi
Set json = CreateRabbitJSON()
Dim configJSON
configJSON = "{""database"":{""host"":""localhost"",""port"":3306,""name"":""mydb""},"
configJSON = configJSON & """cache"":{""enabled"":true,""ttl"":3600,""provider"":""redis""},"
configJSON = configJSON & """email"":{""smtp"":""smtp.gmail.com"",""port"":587}}"

json.Parse(configJSON)

' Root level keys'leri al
Dim allKeys, i
allKeys = json.GetKeys()

Response.Write "<h4>⚙️ Konfigürasyon Ayarları</h4>"

For i = 0 To UBound(allKeys)
    Dim currentKey
    currentKey = allKeys(i)
    
    Response.Write "<h5>📂 " & UCase(currentKey) & "</h5>"
    Response.Write "<ul>"
    
    ' Her kategorinin alt ayarlarını listele
    Select Case currentKey
        Case "database"
            Response.Write "<li>Host: " & json.GetValue("database.host", "") & "</li>"
            Response.Write "<li>Port: " & json.GetValue("database.port", 0) & "</li>"
            Response.Write "<li>Database: " & json.GetValue("database.name", "") & "</li>"
        Case "cache"
            Response.Write "<li>Enabled: " & json.GetValue("cache.enabled", False) & "</li>"
            Response.Write "<li>TTL: " & json.GetValue("cache.ttl", 0) & " seconds</li>"
            Response.Write "<li>Provider: " & json.GetValue("cache.provider", "") & "</li>"
        Case "email"
            Response.Write "<li>SMTP: " & json.GetValue("email.smtp", "") & "</li>"
            Response.Write "<li>Port: " & json.GetValue("email.port", 0) & "</li>"
    End Select
    
    Response.Write "</ul>"
Next

Set json = Nothing
%>
	</body>
</html> 