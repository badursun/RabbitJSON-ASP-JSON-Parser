<!--#include file="../../v2/RabbitJSON.v2.asp" -->
<html>
	<head>
		<title>Sample 13 - E-Commerce Sepet Yönetimi</title>
		<meta charset="utf-8">
	</head>
	<body>
<%
' Sepet JSON'u oluştur
Set json = CreateRabbitJSON()
json.Parse("{}")

' Sepet bilgileri
json.SetValue "cart.sessionId", Session.SessionID
json.SetValue "cart.customerId", 12345
json.SetValue "cart.currency", "TRY"

' Ürünleri sepete ekle
json.SetValue "cart.items.0.productId", 1001
json.SetValue "cart.items.0.name", "MacBook Pro 14"
json.SetValue "cart.items.0.price", 25999.99
json.SetValue "cart.items.0.quantity", 1

json.SetValue "cart.items.1.productId", 1002
json.SetValue "cart.items.1.name", "Magic Mouse"
json.SetValue "cart.items.1.price", 699.99
json.SetValue "cart.items.1.quantity", 2

' Sepet toplamları hesapla
Dim totalAmount, i
totalAmount = 0
For i = 0 To 1
    quantity = json.GetValue("cart.items." & i & ".quantity", 0)
    price = json.GetValue("cart.items." & i & ".price", 0)
    totalAmount = totalAmount + (quantity * price)
Next

json.SetValue "cart.totals.subtotal", totalAmount
json.SetValue "cart.totals.tax", totalAmount * 0.18
json.SetValue "cart.totals.total", totalAmount * 1.18

' Sepet göster
Response.Write "<h4>🛒 Alışveriş Sepeti</h4>"
For i = 0 To 1
    productName = json.GetValue("cart.items." & i & ".name", "")
    quantity = json.GetValue("cart.items." & i & ".quantity", 0)
    price = json.GetValue("cart.items." & i & ".price", 0)
    Response.Write "• " & productName & " (x" & quantity & ") - " & FormatCurrency(price * quantity) & "<br>"
Next
Response.Write "<strong>Toplam: " & FormatCurrency(json.GetValue("cart.totals.total", 0)) & "</strong>"

Set json = Nothing
%>
	</body>
</html> 