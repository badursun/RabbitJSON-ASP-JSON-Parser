<!--#include file="../../v2/RabbitJSON.v2.asp" -->
<html>
	<head>
		<title>Sample 17 - Data Validation & Type Checking</title>
		<meta charset="utf-8">
	</head>
	<body>
<%
' Sample user data validate et
Set json = CreateRabbitJSON()
json.Parse "{""user"":{""name"":""John"",""age"":30,""email"":""john@test.com"",""active"":true}}"

' Validation rules
Response.Write "<h4>✅ Data Validation Results</h4>"

' Required fields check
requiredFields = Array("user.name", "user.email", "user.age")
allValid = True

For Each field In requiredFields
    If json.HasValue(field) Then
        value = json.GetValue(field, "")
        If value <> "" And Not IsNull(value) Then
            Response.Write "✅ " & field & ": " & value & "<br>"
        Else
            Response.Write "❌ " & field & ": Empty or null<br>"
            allValid = False
        End If
    Else
        Response.Write "❌ " & field & ": Missing<br>"
        allValid = False
    End If
Next

' Type validation
age = json.GetValue("user.age", 0)
If IsNumeric(age) And age > 0 And age < 120 Then
    Response.Write "✅ Age validation: Valid (" & age & ")<br>"
Else
    Response.Write "❌ Age validation: Invalid<br>"
    allValid = False
End If

' Email validation (basit)
email = json.GetValue("user.email", "")
If InStr(email, "@") > 0 And InStr(email, ".") > 0 Then
    Response.Write "✅ Email validation: Valid (" & email & ")<br>"
Else
    Response.Write "❌ Email validation: Invalid<br>"
    allValid = False
End If

Response.Write "<br><strong>Overall Result: "
If allValid Then
    Response.Write "<span style='color:green'>✅ All validations passed</span>"
Else
    Response.Write "<span style='color:red'>❌ Some validations failed</span>"
End If
Response.Write "</strong>"

Set json = Nothing
%>
	</body>
</html> 