<!--#include file="../../v2/RabbitJSON.v2.asp" -->
<html>
	<head>
		<title>Sample 3 - Error Handling ile Parse</title>
		<meta charset="utf-8">
	</head>
	<body>
<%
Set json = CreateRabbitJSON()

' Geçersiz JSON test
Set result = json.Parse("{invalid json}")

If result Is Nothing Then
    Response.Write "<div style='color:red'>"
    Response.Write "❌ Parse Hatası: " & json.LastError
    Response.Write "</div>"
    
    ' Hata temizle
    json.ClearError()
    
    ' Tekrar dene - geçerli JSON ile
    Set result = json.Parse("{""status"":""ok""}")
    If Not (result Is Nothing) Then
        Response.Write "<div style='color:green'>"
        Response.Write "✅ İkinci deneme başarılı!"
        Response.Write "</div>"
    End If
End If

Set json = Nothing
%>
	</body>
</html> 