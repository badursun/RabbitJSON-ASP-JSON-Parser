# 🐇 RabbitJSON - The Most Advanced Classic ASP JSON Parser

> 🚀 **Faster, smarter, cleaner.** RabbitJSON is a modern, high-performance JSON parser for Classic ASP — fully surpassing outdated and buggy alternatives like aspJSON.com and rcdmk/aspJSON.

---

## 🌟 Why RabbitJSON?

Working with JSON data in Classic ASP is no longer a nightmare! Most existing solutions are either outdated or break with modern JSON structures. RabbitJSON solves all these issues:

✅ **Path-based** data access (`user.profile.settings.theme`)
✅ Load JSON directly from **HTTP/HTTPS**
✅ Full support for **Unicode**, emojis, and scientific notation
✅ **Configuration system** - set `maxDepth`, `strictMode`
✅ **Default value support** - safe data access
✅ **Advanced error handling** - `HasError()`, `LastError`, `ClearError()`
✅ **Stringify** & compact stringify support
✅ **Comprehensive test coverage** and real-time benchmarking
✅ **Zero dependencies** - pure Classic ASP
✅ Only **41 KB** (uncompressed)

---

## 📊 Benchmark Comparison

Below is a real-world performance comparison:

| Test Category         | RabbitJSON      | [rcdmk/aspJSON](https://github.com/rcdmk/aspJSON) | [gerritvankuipers/aspjson](https://github.com/gerritvankuipers/aspjson) |
| --------------------- | --------------- | ------------------------------------------------- | ----------------------------------------------------------------------- |
| **Local JSON Parse**  | ✅ \~150ms       | ✅ \~200ms                                         | ✅ \~180ms                                                               |
| **HTTP JSON Loading** | ✅ \~500ms       | ❌ Not supported                                   | ❌ Not supported                                                         |
| **Path-based Access** | ✅ `a.b[0].c`    | ❌ Manual code                                     | ❌ Manual code                                                           |
| **Unicode Support**   | ✅ Full          | ❌ Limited                                         | ❌ Partial                                                               |
| **Emoji Support**     | ✅ Excellent     | ❌ None                                            | ❌ Buggy                                                                 |
| **Stringify**         | ✅ Advanced      | ✅ Basic                                           | ✅ Basic                                                                 |
| **Compact Stringify** | ✅ Optimized     | ❌ None                                            | ❌ None                                                                  |
| **Error Handling**    | ✅ Comprehensive | ❌ Minimal                                         | ❌ Poor                                                                  |
| **Memory Management** | ✅ Automatic     | ❌ Manual                                          | ❓ Unknown                                                               |
| **File Size**         | ✅ 41 KB         | \~25 KB                                           | \~9.52 KB                                                               |

---

## 🧪 Test System

RabbitJSON comes with a transparent and robust testing system. Two main test types:

### 📈 1. Comprehensive Function Tests

**104 unique test scenarios** validate all core features:

* **View Test Code:** [comprehensive-test.asp.txt](https://rabbitjson.com/v2/test/comprehensive-test.asp.txt)
* **Live Test Result:** [comprehensive-test.html](https://rabbitjson.com/v2/test/comprehensive-test.html)

### ⚡ 2. Performance & Load Tests

**Real-time performance** benchmarking with large data:

* **View Test Code:** [load-test.asp.txt](https://rabbitjson.com/v2/test/load-test.asp.txt)
* **Live Test Result:** [load-test.html](https://rabbitjson.com/v2/test/load-test.html)

### 📄 Test Data

* **Sample JSON:** [sample-data.json](https://rabbitjson.com/tests/sample-data.json) (7.6 KB real e-commerce data)

---

## ✨ Key Features

### 🔹 JSON Path Access

```asp
Set json = New RabbitJSON
json.Parse "{""users"": [{""name"": ""Ahmet"", ""age"": 25}, {""name"": ""Ayşe"", ""age"": 30}]}"
Response.Write json.GetValue("users.0.name")  ' Output: Ahmet
Response.Write json.GetValue("users.1.name")  ' Output: Ayşe
```

### 🔹 HTTP/HTTPS JSON Loading

```asp
Set json = New RabbitJSON
json.Parse "https://jsonplaceholder.typicode.com/users/1"
Response.Write json.GetValue("name")  ' Fetched directly from API
```

### 🔹 Full Unicode & Emoji Support

```asp
' These characters work flawlessly: ş, ğ, ü, ı, ö, ç, €, 😊, 🚀, 🐇
json.Parse "{\"message\":\"Hello 🐇 RabbitJSON! 😊\"}"
```

### 🔹 Advanced Stringify

```asp
' Readable format
jsonString = json.Stringify(data, 2)

' Compact format (bandwidth saving)
compactJson = json.StringifyCompact(data)
```

### 🔹 Error Handling & Configuration

```asp
Set json = New RabbitJSON
json.Config("maxDepth") = 50
json.Config("strictMode") = True

json.Parse invalidJsonString
If json.HasError() Then
    Response.Write "Parse Error: " & json.LastError
    json.ClearError()
End If
```

### 🔹 Default Value Support

```asp
Set json = New RabbitJSON
json.Parse "{\"user\":{\"name\":\"Mehmet\"}}"
userName = json.GetValue("user.name", "Unknown")
userAge = json.GetValue("user.age", 0)
userPhone = json.GetValue("user.phone", "Not Provided")
```

---

## 🛠️ Installation

### 1. Download File

Include RabbitJSON in your project:

```asp
<!--#include file="RabbitJSON.v2.asp" -->
```

### 2. Start Using

```asp
<%
Set json = New RabbitJSON
json.Parse "{\"name\":\"Mehmet\",\"age\":30}"
Response.Write json.GetValue("name", "")  ' Output: Mehmet
%>
```

---

## 📚 API Documentation

> 📖 **Full API Reference:** [api.md](v2/api.md) — All methods, parameters, and examples included.

### Core Methods

* `Parse(jsonString)` - Parses a JSON string
* `GetValue(path)` - Retrieves value using a path (e.g., `user.profile.name`)
* `SetValue(path, value)` - Writes value to a path
* `Stringify(object, indent)` - Converts object to JSON
* `StringifyCompact(object)` - Compact JSON output

### Utility Methods

* `HasValue(path)` - Checks path existence
* `RemoveValue(path)` - Deletes value at path
* `GetKeys()` - Lists root keys
* `Clear()` - Clears all data

### Factory Functions

* `CreateRabbitJSON()` - Returns a new RabbitJSON instance
* `QuickParse(jsonString)` - Static quick parse method
* `QuickStringify(object, indent)` - Static stringify method

### Advanced Features

* `Config(key, value)` - Configuration (`maxDepth`, `strictMode`)
* `HasError()` - Checks for error state
* `LastError` - Last error message property
* `ClearError()` - Clears error state
* `GetValue(path, defaultValue)` - Safe value access
* `GetValueSimple(path)` - Simple value access
* `Version` - Version info property
* `Count` - Total element count property

---

## 💡 Usage Scenarios

### 🛒 E-commerce Data Parsing

```asp
Set json = New RabbitJSON
json.Parse "https://api.example.com/products"
For i = 0 To json.GetValue("products").Count - 1
    Response.Write json.GetValue("products." & i & ".name", "-") & "<br>"
    Response.Write json.GetValue("products." & i & ".price", "0") & " TL<br>"
Next
```

### ↺ REST API Integration

```asp
Set json = New RabbitJSON
json.Parse Request.Form("jsonData")
userName = json.GetValue("user.profile.displayName")
userEmail = json.GetValue("user.contact.email")
```

### ⚙️ Configuration Management

```asp
Set config = New RabbitJSON
config.Parse Server.MapPath("config.json")
dbServer = config.GetValue("database.server", "localhost")
dbPort = config.GetValue("database.port", 1433)
```

### 🛡️ Secure Production Use

```asp
Set json = New RabbitJSON
json.Config("maxDepth") = 20
json.Config("strictMode") = True

json.Parse userInputJson
If json.HasError() Then
    Response.Write "Invalid JSON format"
    json.ClearError()
Else
    userName = json.GetValue("user.name", "Anonymous")
    Response.Write "Hello " & userName
End If
```

---

## 🔧 Technical Specs

* **Platform:** Classic ASP / VBScript 5.x+
* **Dependencies:** None (uses built-in COM objects only)
* **File Size:** 40.463 bytes (41 KB)
* **Memory Use:** Optimized with auto-cleanup
* **Performance:** Optimized for string operations
* **Compatibility:** IIS 6.0+ and Windows Server 2003+

---

## 🔞 Why Not Others?

### aspJSON.com Issues:

❌ Corrupts Unicode characters
❌ Memory leaks on large files
❌ No HTTP loading support
❌ No path-based access

### rcdmk/aspJSON Issues:

❌ No longer maintained (last update in 2019)
❌ Fails with modern JSON formats
❌ No emoji support
❌ Performance issues

### RabbitJSON Advantages:

✅ **Actively developed in 2025**
✅ **Full test coverage** - 104 test scenarios
✅ **Transparent benchmarking** - real performance metrics
✅ **Production-ready** - used in live environments

---

## 📜 Release Notes

### v2.1.1 (Current)

* ✅ Load JSON from HTTP/HTTPS URLs
* ✅ Config system (`maxDepth`, `strictMode`)
* ✅ Default value support
* ✅ Advanced error handling
* ✅ QuickParse and QuickStringify static methods
* ✅ Version and Count properties
* ✅ Improved key parsing algorithm
* ✅ VBScript compatibility fixes
* ✅ 104 test scenarios with 100% success

### v2.1.0

* ✅ Path-based data access
* ✅ Full Unicode and emoji support
* ✅ Stringify and StringifyCompact
* ✅ Automatic memory management

---

## 🚀 Roadmap & Future Plans

### 📋 v2.2.0 - Performance & Monitoring (Q1 2025)

#### ⚡ Performance Improvements

* String parsing optimization
* Advanced memory management
* Lazy loading support
* Internal caching system
* Async processing for large datasets

#### 📊 Performance Monitoring & Debug

* Built-in profiler
* `GetPerformanceStats()` metrics
* Debug mode with verbose logs

#### 🌐 HTTP Request Features

* Custom headers: `Parse(url, headers)`
* Auth: Basic, Bearer Token, API Key
* Timeout control
* Retry mechanism
* Request delay
* SSL/TLS options
* Proxy support

### 📋 Planned v2.x.x (2026\~)

#### ⚙️ New API Features

* JSON Schema validation
* Diff & Merge support

#### 🔒 Security & Enterprise

* Input sanitization
* Field-level access control
* Audit logging
* Encryption
* Enterprise config options

#### 🎯 Major Innovations

* Native COM component written in C++

---

### 🤝 Contributions & Feedback

* **Feature requests:** via [GitHub Issues](https://github.com/badursun/RabbitJSON-ASP-JSON-Parser/issues)
* **Beta testing:** join the beta program
* **Discussions:** Share feedback via [GitHub Discussions](https://github.com/badursun/RabbitJSON-ASP-JSON-Parser/discussions)

---

## 📄 License

**MIT License** © 2025 Anthony Burak Dursun

```
This software is licensed under the MIT License.
You are free to use it in commercial and personal projects.
```

---

## 🤝 Contribute

To contribute to RabbitJSON:

1. Fork the project
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit changes (`git commit -m 'Add amazing feature'`)
4. Push to your branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

---

## 🔗 Useful Links

* 📖 **API Docs:** [RabbitJSON-Documentation.html](https://rabbitjson.com/v2/documentation.html)
* 📖 **API Reference:** [api.md](v2/api.md)
* 📌 **All Samples:** [all-sample.html](https://rabbitjson.com/v2/all-sample.html)
* 🧪 **Test Runner:** [all-sample-test/](https://rabbitjson.com/v2/all-sample-test/)
* �\udbug **Report Bugs:** [GitHub Issues](https://github.com/badursun/RabbitJSON-ASP-JSON-Parser/issues)
* 💬 **Community Support:** [GitHub Discussions](https://github.com/badursun/RabbitJSON-ASP-JSON-Parser/discussions)

---

**🐇 RabbitJSON - The JSON future of Classic ASP!**
