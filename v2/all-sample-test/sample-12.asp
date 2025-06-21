<!--#include file="../../v2/RabbitJSON.v2.asp" -->
<html>
	<head>
		<title>Sample 12 - Basit Veri Yazma (SetValue)</title>
		<meta charset="utf-8">
	</head>
	<body>
<%
' Boş JSON'dan başla
Set json = CreateRabbitJSON()
json.Parse("{}")

' Basit değerler set et
json.SetValue "name", "Mehmet"
json.SetValue "age", 30
json.SetValue "active", True
json.SetValue "salary", 5000.50

' Sonucu göster
result = json.Stringify(json.Data, 2)
Response.Write "<h4>✅ Oluşturulan JSON:</h4>"
Response.Write "<pre>" & result & "</pre>"

' Değerleri oku
Response.Write "<h4>📖 Okunan Değerler:</h4>"
Response.Write "İsim: " & json.GetValue("name", "") & "<br>"
Response.Write "Yaş: " & json.GetValue("age", 0) & "<br>"
Response.Write "Aktif: " & json.GetValue("active", False) & "<br>"
Response.Write "Maaş: " & FormatCurrency(json.GetValue("salary", 0))

Set json = Nothing
%>
	</body>
</html> 