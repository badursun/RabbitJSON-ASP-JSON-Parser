<!--#include file="../../v2/RabbitJSON.v2.asp" -->
<html>
	<head>
		<title>Sample 5 - GitHub API Entegrasyonu</title>
		<meta charset="utf-8">
	</head>
	<body>
<%
' GitHub API'den user bilgisi çek
Set json = CreateRabbitJSON()
Set result = json.Parse("https://api.github.com/users/octocat")

If Not (result Is Nothing) Then
    login = json.GetValue("login", "")
    name = json.GetValue("name", "")
    followers = json.GetValue("followers", 0)
    publicRepos = json.GetValue("public_repos", 0)
    
    Response.Write "<div style='border:1px solid #ccc; padding:15px; border-radius:8px'>"
    Response.Write "<h4>🐙 GitHub Profile</h4>"
    Response.Write "Username: <strong>" & login & "</strong><br>"
    Response.Write "Name: " & name & "<br>"
    Response.Write "Followers: " & FormatNumber(followers, 0) & "<br>"
    Response.Write "Public Repos: " & publicRepos
    Response.Write "</div>"
End If

Set json = Nothing
%>
	</body>
</html> 