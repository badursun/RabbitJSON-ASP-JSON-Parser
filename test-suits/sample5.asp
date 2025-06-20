<!--#include file="../v2/RabbitJSON.v2.asp" -->
<html>
	<head>
		<title>Sample 1</title>
		<meta charset="utf-8">
	</head>
	<body>
<%
' Gelişmiş configuration ile RabbitJSON oluştur
Set json = New RabbitJSON

' Konfigürasyon ayarları
json.Config("maxDepth") = 50
json.Config("strictMode") = True

' Hata kontrolü ile JSON parse
json.Parse("{ ""kullanici"" : {""ad"" : ""Mehmet"",""email"" : ""test@example.com""} }")

If json.HasError() Then
    Response.Write "Parse Hatası: " & json.LastError
    json.ClearError()
Else
    ' Başarılı - Default değerlerle güvenli erişim
    kullaniciAdi = json.GetValue("kullanici.ad", "Bilinmiyor")
    kullaniciEmail = json.GetValue("kullanici.email", "Yok")
    kullaniciTel = json.GetValue("kullanici.telefon", "Belirtilmemiş")
    
    Response.Write "Ad: " & kullaniciAdi & "<br>"
    Response.Write "Email: " & kullaniciEmail & "<br>"
    Response.Write "Tel: " & kullaniciTel & "<br>"
    
    ' Nesne bilgileri
    Response.Write "Versiyon: " & json.Version & "<br>"
    Response.Write "Eleman Sayısı: " & json.Count & "<br>"
    
    ' Tüm anahtarları listele
    allKeys = json.GetKeys()
    Response.Write "Root Anahtarlar: "
    For i = 0 To UBound(allKeys)
        Response.Write allKeys(i) & " "
    Next
End If

Set json = Nothing
%>
	</body>
</html>