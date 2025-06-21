<!--#include file="../../v2/RabbitJSON.v2.asp" --><%
' API Response oluştur
Set json = CreateRabbitJSON()
json.Parse("{}")

' Response metadata
json.SetValue "status", "success"
json.SetValue "message", "Data retrieved successfully"
json.SetValue "timestamp", Now()
json.SetValue "version", "2.1.1"

' User data ekle
json.SetValue "data.users.0.id", 1001
json.SetValue "data.users.0.username", "john_doe"
json.SetValue "data.users.0.email", "john@example.com"
json.SetValue "data.users.0.role", "admin"
json.SetValue "data.users.0.lastLogin", "2025-01-20"

json.SetValue "data.users.1.id", 1002
json.SetValue "data.users.1.username", "jane_smith"
json.SetValue "data.users.1.email", "jane@example.com"
json.SetValue "data.users.1.role", "user"
json.SetValue "data.users.1.lastLogin", "2025-01-19"

' Pagination bilgisi
json.SetValue "meta.totalCount", 245
json.SetValue "meta.currentPage", 1
json.SetValue "meta.perPage", 10
json.SetValue "meta.totalPages", 25

' JSON response'u oluştur
Response.ContentType = "application/json"
' Response.Write json.StringifyCompact(json.Data)

' Debug için HTML output (production'da remove et)
Response.Write json.Stringify(json.Data, 2)

Set json = Nothing
%>