<!--#include file="../../v2/RabbitJSON.v2.asp" -->
<html>
	<head>
		<title>Sample 8 - Array Elemanlarına Erişim</title>
		<meta charset="utf-8">
	</head>
	<body>
<%
' Array içeren JSON
Set json = CreateRabbitJSON()
Dim arrayJSON
arrayJSON = "{""products"":[{""id"":1,""name"":""Laptop"",""price"":999.99},"
arrayJSON = arrayJSON & "{""id"":2,""name"":""Mouse"",""price"":29.99},"
arrayJSON = arrayJSON & "{""id"":3,""name"":""Keyboard"",""price"":79.99}],"
arrayJSON = arrayJSON & """categories"":[""Electronics"",""Computers"",""Accessories""]}"

json.Parse(arrayJSON)

' Array elemanlarına erişim (0-based index)
firstProduct = json.GetValue("products.0.name", "")
firstPrice = json.GetValue("products.0.price", 0)
secondProduct = json.GetValue("products.1.name", "")
thirdProduct = json.GetValue("products.2.name", "")

' String array erişimi
firstCategory = json.GetValue("categories.0", "")
secondCategory = json.GetValue("categories.1", "")

Response.Write "<h4>🛍️ Ürünler</h4>"
Response.Write "1. " & firstProduct & " - $" & firstPrice & "<br>"
Response.Write "2. " & secondProduct & "<br>"
Response.Write "3. " & thirdProduct & "<br>"
Response.Write "<h4>📂 Kategoriler</h4>"
Response.Write "• " & firstCategory & "<br>"
Response.Write "• " & secondCategory

Set json = Nothing
%>
	</body>
</html> 