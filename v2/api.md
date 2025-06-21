# 🐰 RabbitJSON v2.1.1 - Complete API Reference

## 📋 Table of Contents
- [Properties](#properties)
- [Core Methods](#core-methods)
- [Data Manipulation Methods](#data-manipulation-methods)
- [Serialization Methods](#serialization-methods)
- [Static/Utility Methods](#staticutility-methods)
- [Factory Functions](#factory-functions)
- [Usage Examples](#usage-examples)

---

## 🔧 Properties

### Read-Only Properties

| Property | Type | Parameters | Description |
|----------|------|------------|-------------|
| `Version` | String | None | Returns library version (e.g., "2.1.1") |
| `LastError` | String | None | Returns last error message |
| `Data` | Dictionary Object | None | Returns main data storage object |
| `Count` | Integer | None | Returns number of items in root data object |

### Read-Write Properties

| Property | Type | Parameters | Description |
|----------|------|------------|-------------|
| `Config(key)` | Variant | `key` (String) | Gets configuration value for specified key |
| `Config(key) = value` | - | `key` (String), `value` (Variant) | Sets configuration value |

#### Configuration Keys
- `strictMode` (Boolean) - Default: False
- `parseMode` (String) - Default: "standard"
- `maxDepth` (Integer) - Default: 100
- `allowComments` (Boolean) - Default: False

### Write-Only Properties

| Property | Type | Parameters | Description |
|----------|------|------------|-------------|
| `Set Data = obj` | - | `obj` (Dictionary Object) | Sets main data storage object |

---

## 🔧 Core Methods

### Parse Method

| Method | Return Type | Parameters | Required | Description |
|--------|-------------|------------|----------|-------------|
| `Parse(jsonString)` | Dictionary Object or Nothing | `jsonString` (String) | ✅ Yes | Parses JSON string or loads from HTTP URL |

**Parameters:**
- `jsonString` (String): JSON string to parse OR HTTP/HTTPS URL to load JSON from

**Returns:**
- Dictionary Object on success
- Nothing on error (check LastError)

**Example:**
```vbscript
Set json = New RabbitJSON
Set result = json.Parse("{""name"":""John""}")
' OR
Set result = json.Parse("https://api.example.com/data.json")
```

---

## 📊 Data Manipulation Methods

### GetValue Method

| Method | Return Type | Parameters | Required | Description |
|--------|-------------|------------|----------|-------------|
| `GetValue(path, defaultValue)` | Variant | `path` (String), `defaultValue` (Variant) | ✅ Both | Gets value by dot notation path with default fallback |

**Parameters:**
- `path` (String): Dot notation path (e.g., "user.profile.name", "products.0.price")
- `defaultValue` (Variant): Value to return if path not found

**Returns:**
- Value at specified path
- defaultValue if path not found or error occurs

**Example:**
```vbscript
userName = json.GetValue("user.name", "Unknown")
firstProduct = json.GetValue("products.0.name", "")
```

### GetValueSimple Method

| Method | Return Type | Parameters | Required | Description |
|--------|-------------|------------|----------|-------------|
| `GetValueSimple(path)` | Variant | `path` (String) | ✅ Yes | Gets value by path, returns Empty if not found |

**Parameters:**
- `path` (String): Dot notation path

**Returns:**
- Value at specified path
- Empty if path not found

### SetValue Method

| Method | Return Type | Parameters | Required | Description |
|--------|-------------|------------|----------|-------------|
| `SetValue(path, value)` | None (Sub) | `path` (String), `value` (Variant) | ✅ Both | Sets value by dot notation path, creates nested objects as needed |

**Parameters:**
- `path` (String): Dot notation path where to set value
- `value` (Variant): Value to set (any type: String, Number, Boolean, Object, etc.)

**Example:**
```vbscript
json.SetValue "user.name", "John"
json.SetValue "user.settings.theme", "dark"
json.SetValue "products.0.price", 999.99
```

### RemoveValue Method

| Method | Return Type | Parameters | Required | Description |
|--------|-------------|------------|----------|-------------|
| `RemoveValue(path)` | Boolean | `path` (String) | ✅ Yes | Removes value by dot notation path |

**Parameters:**
- `path` (String): Dot notation path to remove

**Returns:**
- True if successfully removed
- False if path not found or error

### HasValue Method

| Method | Return Type | Parameters | Required | Description |
|--------|-------------|------------|----------|-------------|
| `HasValue(path)` | Boolean | `path` (String) | ✅ Yes | Checks if path exists in data |

**Parameters:**
- `path` (String): Dot notation path to check

**Returns:**
- True if path exists
- False if path not found

### GetKeys Method

| Method | Return Type | Parameters | Required | Description |
|--------|-------------|------------|----------|-------------|
| `GetKeys()` | Array | None | None | Returns array of all keys at root level |

**Returns:**
- Array of String keys
- Empty array if no keys

---

## 📝 Serialization Methods

### Stringify Method

| Method | Return Type | Parameters | Required | Description |
|--------|-------------|------------|----------|-------------|
| `Stringify(obj, indent)` | String | `obj` (Variant), `indent` (Integer) | ✅ obj, ❌ indent | Converts object to formatted JSON string |

**Parameters:**
- `obj` (Variant): Object to convert to JSON
- `indent` (Integer): Number of spaces for indentation (optional, default: 0)

**Returns:**
- Formatted JSON string
- Empty string on error

**Example:**
```vbscript
jsonString = json.Stringify(json.Data, 2)  ' Pretty-printed
jsonString = json.Stringify(json.Data)     ' Compact
```

### StringifyCompact Method

| Method | Return Type | Parameters | Required | Description |
|--------|-------------|------------|----------|-------------|
| `StringifyCompact(obj)` | String | `obj` (Variant) | ✅ Yes | Converts object to compact JSON string (no formatting) |

**Parameters:**
- `obj` (Variant): Object to convert to JSON

**Returns:**
- Compact JSON string

---

## 🛠️ Utility Methods

### Clear Method

| Method | Return Type | Parameters | Required | Description |
|--------|-------------|------------|----------|-------------|
| `Clear()` | None (Sub) | None | None | Clears all data and resets error state |

### ClearError Method

| Method | Return Type | Parameters | Required | Description |
|--------|-------------|------------|----------|-------------|
| `ClearError()` | None (Sub) | None | None | Resets error state only |

### HasError Method

| Method | Return Type | Parameters | Required | Description |
|--------|-------------|------------|----------|-------------|
| `HasError()` | Boolean | None | None | Checks if there was an error |

**Returns:**
- True if there's an error
- False if no error

---

## ⚡ Static/Utility Methods

### QuickParse Method

| Method | Return Type | Parameters | Required | Description |
|--------|-------------|------------|----------|-------------|
| `QuickParse(jsonString)` | Dictionary Object or Nothing | `jsonString` (String) | ✅ Yes | Static method: Quick parse without instance setup |

**Parameters:**
- `jsonString` (String): JSON string to parse

**Returns:**
- Dictionary Object on success
- Nothing on error

### QuickStringify Method

| Method | Return Type | Parameters | Required | Description |
|--------|-------------|------------|----------|-------------|
| `QuickStringify(obj, indent)` | String | `obj` (Variant), `indent` (Integer) | ✅ obj, ❌ indent | Static method: Quick stringify with optional indent |

**Parameters:**
- `obj` (Variant): Object to convert
- `indent` (Integer): Indentation spaces (optional)

**Returns:**
- JSON string

### QuickStringifyCompact Method

| Method | Return Type | Parameters | Required | Description |
|--------|-------------|------------|----------|-------------|
| `QuickStringifyCompact(obj)` | String | `obj` (Variant) | ✅ Yes | Static method: Quick compact stringify |

**Parameters:**
- `obj` (Variant): Object to convert

**Returns:**
- Compact JSON string

---

## 🏭 Factory Functions

### CreateRabbitJSON Function

| Function | Return Type | Parameters | Required | Description |
|----------|-------------|------------|----------|-------------|
| `CreateRabbitJSON()` | RabbitJSON Object | None | None | Factory function for easy instantiation |

**Returns:**
- New RabbitJSON instance

**Example:**
```vbscript
Set json = CreateRabbitJSON()
```

---

## 📚 Usage Examples

### Basic Usage
```vbscript
' Create instance
Set json = New RabbitJSON

' Parse JSON
Set result = json.Parse("{""name"":""John"",""age"":30}")

' Get values
userName = json.GetValue("name", "Unknown")
userAge = json.GetValue("age", 0)

' Set values
json.SetValue "email", "john@example.com"

' Convert back to JSON
jsonString = json.Stringify(json.Data, 2)
```

### Array Access
```vbscript
' Parse array JSON
json.Parse "[{""name"":""Product1""},{""name"":""Product2""}]"

' Access array elements
firstProduct = json.GetValue("0.name", "")
secondProduct = json.GetValue("1.name", "")
```

### HTTP Loading
```vbscript
' Load from URL
Set result = json.Parse("https://api.example.com/users.json")
If Not (result Is Nothing) Then
    userName = json.GetValue("0.name", "")
End If
```

### Error Handling
```vbscript
Set result = json.Parse(invalidJSON)
If result Is Nothing Then
    Response.Write "Error: " & json.LastError
    json.ClearError()
End If
```

### Configuration
```vbscript
' Set configuration
json.Config("maxDepth") = 50
json.Config("strictMode") = True

' Get configuration
maxDepth = json.Config("maxDepth")
```

---

## 📊 Summary Table

| Category | Method Count | Description |
|----------|--------------|-------------|
| **Properties** | 5 | Version, LastError, Data, Count, Config |
| **Core Methods** | 1 | Parse |
| **Data Manipulation** | 5 | GetValue, GetValueSimple, SetValue, RemoveValue, HasValue, GetKeys |
| **Serialization** | 2 | Stringify, StringifyCompact |
| **Utility Methods** | 3 | Clear, ClearError, HasError |
| **Static Methods** | 3 | QuickParse, QuickStringify, QuickStringifyCompact |
| **Factory Functions** | 1 | CreateRabbitJSON |
| **Total Public APIs** | **20** | Complete public interface |

---

*RabbitJSON v2.1.1 - Classic ASP JSON Parser & Builder*  
*© 2025 - MIT License*
