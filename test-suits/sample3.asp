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

' HTTP'den JSON yükle
json.Parse "https://jsonplaceholder.typicode.com/users/1"

' API verilerini oku
userName = json.GetValue("name", "Bilinmiyor")
userEmail = json.GetValue("email", "Bilinmiyor")
userCompany = json.GetValue("company.name", "Bilinmiyor")

Response.Write "Kullanıcı: " & userName
Response.Write "<br>"
Response.Write "Email: " & userEmail
Response.Write "<br>"
Response.Write "Şirket: " & userCompany

Set json = Nothing
%>
	</body>
</html>