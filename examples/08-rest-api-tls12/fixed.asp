<%@ Language="VBScript" CodePage="65001" %>
<% Option Explicit %>
<%
' =====================================================================
' FIXED: robust HTTPS/JSON call from Classic ASP.
'  - MSXML2.ServerXMLHTTP.6.0 (server-safe, uses WinHTTP)
'  - explicit timeouts, User-Agent and Accept headers
'  - transport errors caught, HTTP status checked
'  - value read with a proper pattern, JSON escapes decoded
' TLS 1.2 itself is a SERVER setting - see the README / explainer page.
' For complex JSON use a proven parser library instead of patterns.
' =====================================================================
Response.Charset = "utf-8"

Const API_URL = "https://api.github.com/repos/TanmayMishra22/Legacy-ASP-Rescue"

' Returns the string value of a top-level "key": "value" pair, or Null.
Function JsonString(json, key)
    Dim re, m
    JsonString = Null
    Set re = New RegExp
    re.Pattern = """" & key & """\s*:\s*""((?:[^""\\]|\\.)*)"""
    re.Global = False
    Set m = re.Execute(json)
    If m.Count > 0 Then JsonString = JsonUnescape(m(0).SubMatches(0))
End Function

Function JsonUnescape(s)
    Dim i, c, out, nxt
    out = "" : i = 1
    Do While i <= Len(s)
        c = Mid(s, i, 1)
        If c = "\" And i < Len(s) Then
            nxt = Mid(s, i + 1, 1)
            Select Case nxt
                Case "n": out = out & vbLf
                Case "r": out = out & vbCr
                Case "t": out = out & vbTab
                Case "b": out = out & Chr(8)
                Case "f": out = out & Chr(12)
                Case "u"
                    If i + 5 <= Len(s) Then
                        out = out & ChrW(CLng("&H" & Mid(s, i + 2, 4)))
                        i = i + 4
                    End If
                Case Else: out = out & nxt       ' \"  \\  \/
            End Select
            i = i + 2
        Else
            out = out & c
            i = i + 1
        End If
    Loop
    JsonUnescape = out
End Function

Dim http, status, body, errText, repoName, stars

Set http = Server.CreateObject("MSXML2.ServerXMLHTTP.6.0")
' resolve, connect, send, receive timeouts in milliseconds
http.setTimeouts 5000, 5000, 10000, 15000

On Error Resume Next
http.Open "GET", API_URL, False
http.setRequestHeader "User-Agent", "LegacyAspRescue-Demo/1.0"
http.setRequestHeader "Accept", "application/json"
http.Send
If Err.Number <> 0 Then
    errText = "Could not reach the API (0x" & Hex(Err.Number) & "): " & Err.Description & _
              " - if this mentions the secure channel, enable TLS 1.2 for WinHTTP on this server."
    Err.Clear
Else
    status = http.Status
    body = http.responseText
End If
On Error GoTo 0
Set http = Nothing

If Len(errText) = 0 Then
    If status <> 200 Then
        errText = "The API answered HTTP " & status & ". First part of the reply: " & Left(body, 300)
    Else
        repoName = JsonString(body, "full_name")
        If IsNull(repoName) Then errText = "The reply did not contain full_name."
    End If
End If
%>
<!doctype html>
<html><head><meta charset="utf-8"><title>API call (FIXED)</title></head>
<body>
<h1>Repository</h1>
<% If Len(errText) > 0 Then %>
<p style="color:#b00020"><strong>Problem:</strong> <%= Server.HTMLEncode(errText) %></p>
<% Else %>
<p>Name: <%= Server.HTMLEncode(repoName) %></p>
<% End If %>
</body></html>
