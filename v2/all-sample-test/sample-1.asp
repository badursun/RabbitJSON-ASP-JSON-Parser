<!--#include file="../../v2/RabbitJSON.v2.asp" -->
<html>
	<head>
		<title>Sample 1 - Basit JSON String Parse</title>
		<meta charset="utf-8">
	</head>
	<body>
<%
' JSON string'ini parse et
Set json = CreateRabbitJSON()
Set result = json.Parse("{""name"":""John"",""age"":30}")

If Not (result Is Nothing) Then
    Response.Write "Parse başarılı!"
    Response.Write "<br>İsim: " & json.GetValue("name", "")
    Response.Write "<br>Yaş: " & json.GetValue("age", 0)
Else
    Response.Write "Parse hatası: " & json.LastError
End If

Set json = Nothing
%>
	</body>
</html> 