<!--#include file="../v2/RabbitJSON.v2.asp" -->
<html>
	<head>
		<title>Sample 1</title>
		<meta charset="utf-8">
	</head>
	<body>
<%
' JSON instance oluştur
Set json = New RabbitJSON

' Karmaşık JSON yapısı / veya yükleyin
json.Parse("{ ""user"" : {""profile"" : {""personal"" : {""name"":""Ali"",""age"":25} } } }")

' JSON'a geri çevir
result = json.Stringify(json.Data, 2)
Response.Write "<h4>Verinin İlk Hali</h4>"
Response.Write "<pre>"& result & "</pre>"

' Derin path erişimi
name = json.GetValue("user.profile.personal.name", "Bilinmiyor")
age = json.GetValue("user.profile.personal.age", "Bilinmiyor")

' Veri ekleme
json.SetValue "user.profile.settings.theme", "dark"
json.SetValue "user.profile.settings.lang", "tr"

' JSON'a geri çevir
result = json.Stringify(json.Data, 2)
Response.Write "<h4>Verinin Son Hali</h4>"
Response.Write "<pre>"& result & "</pre>"

Set json = Nothing
%>
	</body>
</html>