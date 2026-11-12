<%@ Language="VBScript" CodePage="65001" %>
<% Option Explicit %>
<%
' =====================================================================
' BROKEN: Cross-site scripting (XSS). Visitor text is written into the
' page as raw HTML, so a link like
'   broken.asp?q=<script>alert('XSS')</script>
' runs JavaScript in every victim's browser (steals sessions, fakes forms).
' =====================================================================
Response.Charset = "utf-8"
Dim q, nm
q  = Request.QueryString("q")
nm = Request.QueryString("name")
%>
<!doctype html>
<html><head><meta charset="utf-8"><title>Search (BROKEN)</title></head>
<body>
<h1>Search</h1>
<form method="get">
  <input name="q" value="<%= q %>">          <!-- attribute injection: " onfocus=... -->
  <button>Search</button>
</form>
<p>You searched for: <%= q %></p>             <!-- raw HTML injection -->
<p>Welcome back, <% Response.Write nm %>!</p>
<script>var lastSearch = '<%= q %>';</script>  <!-- breaks out of the JS string -->
</body></html>
