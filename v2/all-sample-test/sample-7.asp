<!--#include file="../../v2/RabbitJSON.v2.asp" -->
<html>
	<head>
		<title>Sample 7 - Nested Object Erişimi</title>
		<meta charset="utf-8">
	</head>
	<body>
<%
' Derinlemesine nested JSON
Set json = CreateRabbitJSON()
Dim nestedJSON
nestedJSON = "{""user"":{""profile"":{""personal"":{""name"":""Ali"","
nestedJSON = nestedJSON & """age"":25},""contact"":{""email"":""ali@test.com"","
nestedJSON = nestedJSON & """phone"":""+90555123456""}},""settings"":{""theme"":""dark"","
nestedJSON = nestedJSON & """language"":""tr""}}}"

json.Parse(nestedJSON)

' Derin path erişimi
userName = json.GetValue("user.profile.personal.name", "")
userAge = json.GetValue("user.profile.personal.age", 0)
userEmail = json.GetValue("user.profile.contact.email", "")
userTheme = json.GetValue("user.settings.theme", "light")
userLang = json.GetValue("user.settings.language", "en")

Response.Write "<h4>👤 Kullanıcı Profili</h4>"
Response.Write "İsim: " & userName & "<br>"
Response.Write "Yaş: " & userAge & "<br>"
Response.Write "Email: " & userEmail & "<br>"
Response.Write "Tema: " & userTheme & "<br>"
Response.Write "Dil: " & userLang

Set json = Nothing
%>
	</body>
</html> 