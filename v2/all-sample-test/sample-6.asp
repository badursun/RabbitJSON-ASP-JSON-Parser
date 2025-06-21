<!--#include file="../../v2/RabbitJSON.v2.asp" -->
<html>
	<head>
		<title>Sample 6 - Basit Veri Erişimi</title>
		<meta charset="utf-8">
	</head>
	<body>
<%
' Sample data parse et
Set json = CreateRabbitJSON()
json.Parse "{""name"":""Mehmet"",""age"":30,""active"":true,""salary"":null}"

' Farklı tip değerler al
userName = json.GetValue("name", "Bilinmiyor")           ' String
userAge = json.GetValue("age", 0)                        ' Number
isActive = json.GetValue("active", False)                ' Boolean
userSalary = json.GetValue("salary", "Belirtilmemiş")    ' Null değer
userPhone = json.GetValue("phone", "Yok")                ' Olmayan key

Response.Write "👤 İsim: " & userName & "<br>"
Response.Write "🎂 Yaş: " & userAge & "<br>"
Response.Write "✅ Aktif: " & isActive & "<br>"
Response.Write "💰 Maaş: " & userSalary & "<br>"
Response.Write "📱 Telefon: " & userPhone

Set json = Nothing
%>
	</body>
</html> 