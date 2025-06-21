<!--#include file="../../v2/RabbitJSON.v2.asp" -->
<html>
	<head>
		<title>Sample 9 - Array Üzerinde Loop</title>
		<meta charset="utf-8">
	</head>
	<body>
<%
' Ürün listesi JSON'u
Set json = CreateRabbitJSON()
Dim productJSON
productJSON = "{""products"":[{""name"":""Laptop"",""price"":999},"
productJSON = productJSON & "{""name"":""Mouse"",""price"":29},"
productJSON = productJSON & "{""name"":""Keyboard"",""price"":79}]}"

json.Parse(productJSON)

' Array uzunluğunu bul (manuel olarak)
Dim productCount, i
productCount = 0
Do While json.HasValue("products." & productCount)
    productCount = productCount + 1
Loop

Response.Write "<h4>🛒 Ürün Listesi (" & productCount & " adet)</h4>"
Response.Write "<ul>"

' Array üzerinde loop
For i = 0 To productCount - 1
    productName = json.GetValue("products." & i & ".name", "")
    productPrice = json.GetValue("products." & i & ".price", 0)
    
    If productName <> "" Then
        Response.Write "<li>" & productName & " - $" & productPrice & "</li>"
    End If
Next

Response.Write "</ul>"

Set json = Nothing
%>
	</body>
</html> 