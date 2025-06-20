# 🐇 RabbitJSON - En Gelişmiş Classic ASP JSON Parser

> 🚀 **Daha hızlı, daha akıllı, daha temiz.** RabbitJSON, Classic ASP için modern, yüksek performanslı JSON parser'dır — aspJSON.com ve rcdmk/aspJSON gibi eski ve hatalı alternatifleri tamamen geride bırakır.

---

## 🌟 Neden RabbitJSON?

Classic ASP'de JSON verisiyle çalışmak artık çile değil! Mevcut çözümlerin çoğu ya güncellenmiyor ya da modern JSON yapılarıyla bozuluyor. RabbitJSON, tüm bu sorunları çözüyor:

✅ **Path-based** veri erişimi (`user.profile.settings.theme`)  
✅ **HTTP/HTTPS** üzerinden doğrudan JSON yükleme  
✅ **Unicode**, emoji ve bilimsel sayı tam desteği  
✅ **Configuration system** - maxDepth, strictMode ayarları  
✅ **Default value support** - güvenli veri erişimi  
✅ **Gelişmiş hata yönetimi** - HasError(), LastError, ClearError()  
✅ **Stringify** & kompakt stringify özellikleri  
✅ **Kapsamlı test yapısı** ve gerçek zamanlı performans ölçümü  
✅ **Sıfır dış bağımlılık** - sadece Classic ASP  
✅ **Sadece 41 KB** boyut (sıkıştırılmamış)

---

## 📊 Benchmark Karşılaştırması

Aşağıdaki tablo, gerçek zamanlı testler ve performans ölçümlerine dayanmaktadır:

| Test Kategorisi            | RabbitJSON   | [rcdmk/aspJSON](https://github.com/rcdmk/aspJSON) | [gerritvankuipers/aspjson](https://github.com/gerritvankuipers/aspjson)  |
|----------------------------|-----------------|------------------|--------------|
| **Yerel JSON Parse**       | ✅ ~150ms        | ✅ ~200ms         | ✅ ~180ms     |
| **HTTP JSON Yükleme**      | ✅ ~500ms        | ❌ Desteklenmiyor | ❌ Desteklenmiyor |
| **Path Tabanlı Erişim**    | ✅ `a.b[0].c`    | ❌ Manuel kod     | ❌ Manuel kod |
| **Unicode Desteği**        | ✅ Tam destek    | ❌ Sınırlı        | ❌ Kısmi      |
| **Emoji Desteği**          | ✅ Mükemmel      | ❌ Desteklenmiyor | ❌ Hatalı     |
| **Stringify**              | ✅ Gelişmiş      | ✅ Basit          | ✅ Basit      |
| **Kompakt Stringify**      | ✅ Optimize      | ❌ Yok            | ❌ Yok        |
| **Hata Yönetimi**          | ✅ Kapsamlı      | ❌ Minimal        | ❌ Yetersiz   |
| **Bellek Yönetimi**        | ✅ Otomatik      | ❌ Manuel         | ❓ Belirsiz   |
| **Dosya Boyutu**           | ✅ 41 KB         | ~25 KB           | ~9.52 KB       |

---

## 🧪 Test Sistemi

RabbitJSON, şeffaf ve kapsamlı test sistemiyle gelir. İki ana test türümüz var:

### 📈 1. Kapsamlı Fonksiyon Testleri
**104 farklı test senaryosu** ile tüm fonksiyonları test eder:

- **Test Kodunu Gör:** [comprehensive-test.asp.txt](https://rabbitjson.com/v2/test/comprehensive-test.asp.txt)
- **Son Test Sonucu:** [comprehensive-test.html](https://rabbitjson.com/v2/test/comprehensive-test.html)

### ⚡ 2. Performans & Yük Testleri
**Gerçek zamanlı performans** ölçümü ve büyük veri testleri:

- **Test Kodunu Gör:** [load-test.asp.txt](https://rabbitjson.com/v2/test/load-test.asp.txt)
- **Son Test Sonucu:** [load-test.html](https://rabbitjson.com/v2/test/load-test.html)

### 📄 Test Verisi
- **Örnek JSON:** [sample-data.json](https://rabbitjson.com/tests/sample-data.json) (7.6 KB, gerçek e-ticaret verisi)

---

## ✨ Öne Çıkan Özellikler

### 🔹 JSON Path Erişimi
```asp
Set json = New RabbitJSON
json.Parse "{""users"":[{""name"":""Ahmet"",""age"":25},{""name"":""Ayşe"",""age"":30}]}"
Response.Write json.GetValue("users.0.name")  ' Çıktı: Ahmet
Response.Write json.GetValue("users.1.name")  ' Çıktı: Ayşe
```

### 🔹 HTTP/HTTPS JSON Yükleme
```asp
Set json = New RabbitJSON
json.Parse "https://jsonplaceholder.typicode.com/users/1"
Response.Write json.GetValue("name")  ' API'den direkt veri
```

### 🔹 Unicode & Emoji Tam Desteği
```asp
' Bu karakterler sorunsuz çalışır: ş, ğ, ü, ı, ö, ç, €, 😊, 🚀, 🐇
json.Parse "{""mesaj"":""Merhaba 🐇 RabbitJSON! 😊""}"
```

### 🔹 Gelişmiş Stringify
```asp
' Okunabilir format
jsonString = json.Stringify(data, 2)

' Kompakt format (bandwidth tasarrufu)
compactJson = json.StringifyCompact(data)
```

### 🔹 Gelişmiş Hata Yönetimi & Configuration
```asp
Set json = New RabbitJSON
' Konfigürasyon ayarları
json.Config("maxDepth") = 50
json.Config("strictMode") = True

json.Parse invalidJsonString
If json.HasError() Then
    Response.Write "Parse Hatası: " & json.LastError
    json.ClearError()
End If
```

### 🔹 Default Değer Desteği
```asp
Set json = New RabbitJSON
json.Parse "{""user"":{""name"":""Mehmet""}}"
' Güvenli erişim - yoksa default değer döner
userName = json.GetValue("user.name", "Bilinmiyor")
userAge = json.GetValue("user.age", 0)
userPhone = json.GetValue("user.phone", "Belirtilmemiş")
```

---

## 🛠️ Kurulum

### 1. Dosyayı İndirin
RabbitJSON'yi projenize dahil edin:

```asp
<!--#include file="RabbitJSON.v2.asp" -->
```

### 2. Kullanmaya Başlayın
```asp
<%
Set json = New RabbitJSON
json.Parse "{""ad"":""Mehmet"",""yaş"":30}"
Response.Write json.GetValue("ad", "")  ' Çıktı: Mehmet
%>
```

---

## 📚 API Dokümantasyonu

### Temel Metodlar
- `Parse(jsonString)` - JSON string'ini parse eder
- `GetValue(path)` - Path ile veri alır (`user.profile.name`)
- `SetValue(path, value)` - Path ile veri yazar
- `Stringify(object, indent)` - Object'i JSON'a çevirir
- `StringifyCompact(object)` - Kompakt JSON formatı

### Utility Metodlar
- `HasValue(path)` - Path'in varlığını kontrol eder
- `RemoveValue(path)` - Path'deki veriyi siler
- `GetKeys()` - Root level anahtarları listeler
- `Clear()` - Tüm veriyi temizler

### Factory Fonksiyonlar
- `CreateRabbitJSON()` - Yeni RabbitJSON instance
- `QuickParse(jsonString)` - Hızlı parse (static method)
- `QuickStringify(object, indent)` - Hızlı stringify (static method)

### Gelişmiş Özellikler
- `Config(key, value)` - Yapılandırma ayarları (`maxDepth`, `strictMode`)
- `HasError()` - Hata durumu kontrolü
- `LastError` - Son hata mesajı property'si
- `ClearError()` - Hata durumunu temizle
- `GetValue(path, defaultValue)` - Default değer desteği ile güvenli erişim
- `GetValueSimple(path)` - Basit değer erişimi
- `Version` - Sürüm bilgisi property'si
- `Count` - Eleman sayısı property'si

---

## 💡 Kullanım Senaryoları

### 🛒 E-ticaret Veri İşleme
```asp
Set json = New RabbitJSON
json.Parse "https://api.example.com/products"
' Array elementlerine erişim
For i = 0 To json.GetValue("products").Count - 1
    Response.Write json.GetValue("products." & i & ".name", "-") & "<br>"
    Response.Write json.GetValue("products." & i & ".price", "0") & " TL<br>"
Next
```

### 🔄 REST API Entegrasyonu
```asp
Set json = New RabbitJSON
json.Parse Request.Form("jsonData")
userName = json.GetValue("user.profile.displayName")
userEmail = json.GetValue("user.contact.email")
```

### ⚙️ Konfigürasyon Yönetimi
```asp
Set config = New RabbitJSON
config.Parse Server.MapPath("config.json")
dbServer = config.GetValue("database.server", "localhost")
dbPort = config.GetValue("database.port", 1433)
```

### 🛡️ Güvenli Production Kullanımı
```asp
Set json = New RabbitJSON
' Production için güvenli ayarlar
json.Config("maxDepth") = 20
json.Config("strictMode") = True

json.Parse userInputJson
If json.HasError() Then
    Response.Write "Geçersiz JSON formatı"
    json.ClearError()
Else
    ' Güvenli veri işleme
    userName = json.GetValue("user.name", "Anonim")
    Response.Write "Merhaba " & userName
End If
```

---

## 🔧 Teknik Özellikler

- **Platform:** Classic ASP / VBScript 5.x+
- **Bağımlılık:** Sıfır (sadece built-in COM objects)
- **Dosya Boyutu:** 40.463 bayt (41 KB)
- **Bellek Kullanımı:** Otomatik cleanup ile optimize
- **Performans:** String işlemleri için optimize edilmiş
- **Uyumluluk:** IIS 6.0+ ve Windows Server 2003+

---

## 🆚 Neden Diğerleri Değil?

### aspJSON.com Sorunları:
❌ Unicode karakterlerde bozulma  
❌ Büyük JSON dosyalarında bellek sızıntısı  
❌ HTTP yükleme desteği yok  
❌ Path-based erişim yok  

### rcdmk/aspJSON Sorunları:
❌ Artık güncellenmiyor (son güncelleme 2019)  
❌ Modern JSON yapılarında hatalar  
❌ Emoji desteği yok  
❌ Performans sorunları  

### RabbitJSON Avantajları:
✅ **2025'te geliştiriliyor** - güncel ve modern  
✅ **Kapsamlı test coverage** - 104 test senaryosu  
✅ **Gerçek zamanlı benchmark** - şeffaf performans  
✅ **Production-ready** - canlı projelerde kullanılıyor  

---

## 📝 Sürüm Notları

### v2.1.1 (Güncel)
- ✅ HTTP/HTTPS URL'den JSON yükleme
- ✅ Configuration system (maxDepth, strictMode)
- ✅ Default value support - GetValue(path, defaultValue)
- ✅ Gelişmiş hata yönetimi (HasError, LastError, ClearError)
- ✅ QuickParse ve QuickStringify static methodları
- ✅ Version ve Count property'leri
- ✅ Geliştirilmiş key parsing algoritması
- ✅ VBScript uyumluluk düzeltmeleri
- ✅ 104 test senaryosu ile %100 başarı oranı

### v2.1.0
- ✅ Path-based veri erişimi
- ✅ Unicode ve emoji tam desteği
- ✅ Stringify ve StringifyCompact
- ✅ Otomatik bellek yönetimi

---

## 📄 Lisans

**MIT License** © 2025 Anthony Burak Dursun

```
Bu yazılım MIT lisansı altında dağıtılmaktadır.
Ticari ve kişisel projelerde özgürce kullanabilirsiniz.
```

---

## 🤝 Katkıda Bulun

RabbitJSON'yi geliştirmek için:

1. Projeyi fork edin
2. Feature branch oluşturun (`git checkout -b feature/amazing-feature`)
3. Değişikliklerinizi commit edin (`git commit -m 'Add amazing feature'`)
4. Branch'inizi push edin (`git push origin feature/amazing-feature`)
5. Pull Request açın

---

## 🔗 Faydalı Linkler

- 📖 **API Dokümantasyonu:** [RabbitJSON-Documentation.html](https://rabbitjson.com/v2/documentation.html)
- 🐛 **Bug Raporu:** [GitHub Issues](https://github.com/your-repo/RabbitJSON/issues)
- 💬 **Destek:** [Discussions](https://github.com/your-repo/RabbitJSON/discussions)

---

**🐇 RabbitJSON - Classic ASP'nin JSON geleceği!**