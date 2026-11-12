<%@ Language="VBScript" CodePage="65001" %>
<%
' =====================================================================
' BROKEN: calling a modern HTTPS/JSON API the old way.
'  - "MSXML2.ServerXMLHTTP" = old version 3.0; on older Windows builds it
'    cannot negotiate TLS 1.2 -> msxml3.dll error '80072f7d'
'    "An error occurred in the secure channel support"
'  - no timeouts: a slow API hangs the page (and an ASP thread)
'  - no User-Agent: many APIs (GitHub included) answer 403
'  - status code never checked, JSON "parsed" with InStr/Mid
' =====================================================================
Response.Charset = "utf-8"
Dim http, body, p, repoName

Set http = Server.CreateObject("MSXML2.ServerXMLHTTP")
http.Open "GET", "https://api.github.com/repos/TanmayMishra22/Legacy-ASP-Rescue", False
http.Send
body = http.responseText

p = InStr(body, """full_name"":")
repoName = Mid(body, p + 13, 40)      ' guesses the length; breaks on any change
%>
<!doctype html>
<html><head><meta charset="utf-8"><title>API call (BROKEN)</title></head>
<body>
<h1>Repository</h1>
<p>Name: <%= repoName %></p>
</body></html>
