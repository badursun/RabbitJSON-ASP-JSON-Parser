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

' JSON parse et
json.Parse "{""ad"":""Mehmet"",""yaş"":30,""aktif"":true}"

' Veri oku
Response.Write json.GetValue("ad", "Bilinmiyor")  ' Çıktı: Mehmet
Response.Write "<br>"
Response.Write json.GetValue("yaş", "Bilinmiyor") ' Çıktı: 30

Set json = Nothing
%>
	</body>
</html>