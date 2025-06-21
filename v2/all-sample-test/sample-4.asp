<!--#include file="../../v2/RabbitJSON.v2.asp" -->
<html>
	<head>
		<title>Sample 4 - HTTP API'den Veri Çekme</title>
		<meta charset="utf-8">
	</head>
	<body>
<%
' API'den direkt JSON yükle
Set json = CreateRabbitJSON()
Set result = json.Parse("https://jsonplaceholder.typicode.com/users/1")

If Not (result Is Nothing) Then
    userName = json.GetValue("name", "Bilinmiyor")
    userEmail = json.GetValue("email", "Bilinmiyor")
    userCompany = json.GetValue("company.name", "Bilinmiyor")
    
    Response.Write "<h4>API'den Gelen Veri:</h4>"
    Response.Write "👤 Kullanıcı: " & userName & "<br>"
    Response.Write "📧 Email: " & userEmail & "<br>"
    Response.Write "🏢 Şirket: " & userCompany
Else
    Response.Write "API Error: " & json.LastError
End If

Set json = Nothing
%>
	</body>
</html> 