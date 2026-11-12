<%@ Language="VBScript" CodePage="65001" %>
<%
' =====================================================================
' BROKEN: errors are hidden.
'  - "On Error Resume Next" at the top silently skips every failing line
'  - when it is removed, IIS shows only "500 - Internal server error"
'  - nothing is written to any log, so nobody knows what failed
' =====================================================================
On Error Resume Next
Response.Charset = "utf-8"
Const CONN_STR = "Provider=MSOLEDBSQL;Data Source=.\SQLEXPRESS;Initial Catalog=RescueDemo;Integrated Security=SSPI;"

Dim conn, rs, id
id = Request.QueryString("id")

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open CONN_STR
' Typo in the table name ("Orderz") - this fails, but the error is swallowed.
Set rs = conn.Execute("SELECT Total FROM Orderz WHERE OrderID = " & CLng(id))
%>
<!doctype html>
<html><head><meta charset="utf-8"><title>Order (BROKEN)</title></head>
<body>
<h1>Order <%= Server.HTMLEncode(id & "") %></h1>
<p>Order total: <%= rs("Total") %></p>   <!-- shows nothing, no clue why -->
</body></html>
