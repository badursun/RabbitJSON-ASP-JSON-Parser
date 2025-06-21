<!--#include file="../../v2/RabbitJSON.v2.asp" -->
<html>
	<head>
		<title>Sample 2 - Kompleks JSON Parse</title>
		<meta charset="utf-8">
	</head>
	<body>
<%
' Kompleks JSON yapısı
Dim complexJSON
complexJSON = "{""company"":{""name"":""RabbitCMS"","
complexJSON = complexJSON & """employees"":[{""name"":""John"",""role"":""Dev""},"
complexJSON = complexJSON & "{""name"":""Jane"",""role"":""Designer""}]}}"

Set json = CreateRabbitJSON()
Set result = json.Parse(complexJSON)

If Not (result Is Nothing) Then
    ' Şirket adı
    companyName = json.GetValue("company.name", "")
    
    ' İlk çalışan
    firstEmployee = json.GetValue("company.employees.0.name", "")
    firstRole = json.GetValue("company.employees.0.role", "")
    
    Response.Write "Şirket: " & companyName & "<br>"
    Response.Write "İlk Çalışan: " & firstEmployee & " (" & firstRole & ")"
End If

Set json = Nothing
%>
	</body>
</html> 