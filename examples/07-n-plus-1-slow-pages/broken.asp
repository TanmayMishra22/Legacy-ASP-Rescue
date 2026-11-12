<%@ Language="VBScript" CodePage="65001" %>
<% Option Explicit %>
<%
' =====================================================================
' BROKEN: the "N+1 queries" pattern. One query for the list, then one
' more query for EVERY row. 2,000 customers = 2,001 round trips.
' Fast on the developer's 10-row test database, painfully slow in production.
' Demo only. Needs the RescueDemo database from ../assets/demo-db.sql
' =====================================================================
Response.Charset = "utf-8"
Const CONN_STR = "Provider=MSOLEDBSQL;Data Source=.\SQLEXPRESS;Initial Catalog=RescueDemo;Integrated Security=SSPI;"

Dim conn, rsC, rsO, t0, queries
t0 = Timer()
queries = 0

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open CONN_STR
Set rsC = conn.Execute("SELECT CustomerID, Name FROM Customers ORDER BY Name")
queries = queries + 1
%>
<!doctype html>
<html><head><meta charset="utf-8"><title>Customer report (BROKEN)</title></head>
<body>
<h1>Customer report</h1>
<table border="1" cellpadding="6">
<tr><th>Customer</th><th>Orders</th><th>Total</th></tr>
<%
Do While Not rsC.EOF
    ' One extra database round trip per customer:
    Set rsO = conn.Execute("SELECT COUNT(*) AS N, ISNULL(SUM(Total), 0) AS T " & _
                           "FROM Orders WHERE CustomerID = " & CLng(rsC("CustomerID")))
    queries = queries + 1
    Response.Write "<tr><td>" & Server.HTMLEncode(rsC("Name") & "") & "</td><td>" & rsO("N") & _
                   "</td><td>" & FormatNumber(rsO("T"), 2) & "</td></tr>"
    rsO.Close
    rsC.MoveNext
Loop
rsC.Close : Set rsC = Nothing : Set rsO = Nothing
conn.Close : Set conn = Nothing
%>
</table>
<p><%= queries %> queries, <%= FormatNumber(Timer() - t0, 3) %> seconds</p>
</body></html>
