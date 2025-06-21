<!--#include file="../../v2/RabbitJSON.v2.asp" -->
<html>
	<head>
		<title>Sample 10 - Nested Object Loop</title>
		<meta charset="utf-8">
	</head>
	<body>
<%
' Kullanıcı listesi
Set json = CreateRabbitJSON()
Dim userJSON
userJSON = "{""users"":[{""name"":""John"",""role"":""Admin"",""active"":true},"
userJSON = userJSON & "{""name"":""Jane"",""role"":""Editor"",""active"":false},"
userJSON = userJSON & "{""name"":""Bob"",""role"":""User"",""active"":true}]}"

json.Parse(userJSON)

' Kullanıcı sayısını bul
Dim userCount, i
userCount = 0
Do While json.HasValue("users." & userCount)
    userCount = userCount + 1
Loop

Response.Write "<h4>👥 Kullanıcı Listesi</h4>"
Response.Write "<table border='1' style='border-collapse:collapse'>"
Response.Write "<tr><th>İsim</th><th>Rol</th><th>Durum</th></tr>"

For i = 0 To userCount - 1
    userName = json.GetValue("users." & i & ".name", "")
    userRole = json.GetValue("users." & i & ".role", "")
    isActive = json.GetValue("users." & i & ".active", False)
    
    Dim statusText, statusColor
    If isActive Then
        statusText = "Aktif"
        statusColor = "green"
    Else
        statusText = "Pasif"
        statusColor = "red"
    End If
    
    Response.Write "<tr>"
    Response.Write "<td>" & userName & "</td>"
    Response.Write "<td>" & userRole & "</td>"
    Response.Write "<td style='color:" & statusColor & "'>" & statusText & "</td>"
    Response.Write "</tr>"
Next

Response.Write "</table>"

Set json = Nothing
%>
	</body>
</html> 