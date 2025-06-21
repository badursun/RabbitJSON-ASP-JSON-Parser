<!--#include file="../../v2/RabbitJSON.v2.asp" -->
<html>
	<head>
		<title>Sample 15 - Configuration Management</title>
		<meta charset="utf-8">
	</head>
	<body>
<%
' Environment-based config
Set json = CreateRabbitJSON()
json.Parse("{}")

' Environment detect (örnek)
Dim currentEnv
currentEnv = "development"  ' production'da farklı logic

' Base settings
json.SetValue "app.name", "RabbitJSON Application"
json.SetValue "app.version", "2.1.1"

' Environment-specific
If currentEnv = "development" Then
    json.SetValue "database.host", "localhost"
    json.SetValue "database.name", "myapp_dev"
    json.SetValue "debug.enabled", True
    json.SetValue "cache.enabled", False
ElseIf currentEnv = "production" Then
    json.SetValue "database.host", "prod-server.com"
    json.SetValue "database.name", "myapp_prod"
    json.SetValue "debug.enabled", False
    json.SetValue "cache.enabled", True
End If

' Security settings
json.SetValue "security.sessionTimeout", 30
json.SetValue "security.maxLoginAttempts", 5

' Feature flags
json.SetValue "features.newDashboard", True
json.SetValue "features.betaFeatures", (currentEnv <> "production")

' Application-level'da sakla
Application("AppConfig") = json.StringifyCompact(json.Data)

Response.Write "<h4>⚙️ " & UCase(currentEnv) & " Configuration</h4>"
Response.Write "<pre>" & json.Stringify(json.Data, 2) & "</pre>"

' Kullanım örneği
Response.Write "<h5>📖 Config Kullanımı:</h5>"
Response.Write "Database Host: " & json.GetValue("database.host", "") & "<br>"
Response.Write "Debug Mode: " & json.GetValue("debug.enabled", False) & "<br>"
If json.GetValue("cache.enabled", False) = True Then
    Response.Write "Cache: Enabled"
Else
    Response.Write "Cache: Disabled"
End If

Set json = Nothing
%>
	</body>
</html> 