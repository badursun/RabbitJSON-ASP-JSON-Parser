<!--#include file="../v2/RabbitJSON.v2.asp" -->
<html>
	<head>
		<title>Sample 1</title>
		<meta charset="utf-8">
	</head>
	<body>
<%
' Sample JSON 
Dim JSON_SAMPLE
	JSON_SAMPLE = "{""kullanici"" : { ""ad"" : ""Ahmet"" , ""yas"" : 25} }"

' JSON instance oluştur
Set JSON = CreateRabbitJSON()

	' JSON parse et
	JSON.Parse(JSON_SAMPLE)

' Veri oku
kullaniciAdi = JSON.GetValue("kullanici.ad", "Bilinmiyor")
kullaniciYas = JSON.GetValue("kullanici.yas", "*")

Response.Write "Merhaba " & kullaniciAdi & "! Yaşınız: " & kullaniciYas

Set JSON = Nothing
%>
	</body>
</html>