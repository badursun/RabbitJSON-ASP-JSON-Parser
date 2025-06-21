<!--#include file="../../v2/RabbitJSON.v2.asp" -->
<html>
	<head>
		<title>Sample 16 - Error Handling & Debugging</title>
		<meta charset="utf-8">
	</head>
	<body>
<%
Set json = CreateRabbitJSON()

' Error test scenarios
Dim testCases(3)
testCases(0) = "{""valid"": ""json""}"        ' Valid
testCases(1) = "{invalid json}"               ' Invalid syntax
testCases(2) = ""                             ' Empty
testCases(3) = "https://invalid-url.com"     ' Invalid URL

Response.Write "<h4>🧪 Error Handling Tests</h4>"

For i = 0 To 3
    Response.Write "<h5>Test " & (i+1) & ":</h5>"
    
    Set result = json.Parse(testCases(i))
    
    If result Is Nothing Then
        Response.Write "<div style='color:red; border:1px solid red; padding:10px'>"
        Response.Write "❌ Error: " & json.LastError
        Response.Write "</div>"
        json.ClearError()
    Else
        Response.Write "<div style='color:green; border:1px solid green; padding:10px'>"
        Response.Write "✅ Success! Items: " & json.Count
        Response.Write "</div>"
    End If
    Response.Write "<br>"
Next

' HasError test
json.Parse("{invalid}")
If json.HasError() Then
    Response.Write "<p>Error detected: " & json.LastError & "</p>"
    json.ClearError()
    Response.Write "<p>Error cleared successfully!</p>"
End If

Set json = Nothing
%>
	</body>
</html> 