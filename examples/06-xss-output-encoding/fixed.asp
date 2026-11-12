<%@ Language="VBScript" CodePage="65001" %>
<% Option Explicit %>
<%
' =====================================================================
' FIXED: every piece of visitor data is ENCODED for the place it lands.
'   H()  -> HTML text and attribute values (Server.HTMLEncode, Null-safe)
'   JS() -> inside a JavaScript string literal
'   U()  -> inside a URL query string (Server.URLEncode)
' =====================================================================
Response.Charset = "utf-8"

Function H(v)
    H = Server.HTMLEncode(v & "")      ' & "" turns Null into "" (HTMLEncode(Null) errors)
End Function

Function U(v)
    U = Server.URLEncode(v & "")
End Function

' Escape for a single- or double-quoted JavaScript string.
' Every non-alphanumeric ASCII character becomes \xHH or \uHHHH, so quotes,
' backslashes, < > and line breaks can never end the string or the <script> block.
Function JS(v)
    Dim s, i, c, code, out
    s = v & ""
    out = ""
    For i = 1 To Len(s)
        c = Mid(s, i, 1)
        code = AscW(c)
        If code < 0 Then code = code + 65536
        If (code >= 48 And code <= 57) Or (code >= 65 And code <= 90) Or (code >= 97 And code <= 122) Or code = 32 Then
            out = out & c
        ElseIf code < 256 Then
            out = out & "\x" & Right("0" & Hex(code), 2)
        Else
            out = out & "\u" & Right("000" & Hex(code), 4)
        End If
    Next
    JS = out
End Function

Dim q, nm
q  = Left(Request.QueryString("q") & "", 200)
nm = Left(Request.QueryString("name") & "", 100)
%>
<!doctype html>
<html><head><meta charset="utf-8"><title>Search (FIXED)</title></head>
<body>
<h1>Search</h1>
<form method="get">
  <input name="q" value="<%= H(q) %>">
  <button>Search</button>
</form>
<p>You searched for: <%= H(q) %></p>
<p>Welcome back, <%= H(nm) %>!</p>
<p><a href="fixed.asp?q=<%= U(q) %>">Link to this search</a></p>
<script>var lastSearch = '<%= JS(q) %>';</script>
</body></html>
