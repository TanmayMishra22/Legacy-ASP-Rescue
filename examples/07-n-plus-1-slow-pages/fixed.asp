<%@ Language="VBScript" CodePage="65001" %>
<% Option Explicit %>
<%
' =====================================================================
' FIXED: ONE query with a JOIN + GROUP BY does all the work in SQL Server,
' GetRows copies the result into an array, and the connection is closed
' before the HTML is built. 1 round trip no matter how many customers.
' Demo only. Needs the RescueDemo database from ../assets/demo-db.sql
' =====================================================================
Response.Charset = "utf-8"
Const CONN_STR = "Provider=MSOLEDBSQL;Data Source=.\SQLEXPRESS;Initial Catalog=RescueDemo;Integrated Security=SSPI;"

Dim conn, rs, sql, data, i, t0
t0 = Timer()

sql = "SELECT c.Name, COUNT(o.OrderID) AS OrderCount, ISNULL(SUM(o.Total), 0) AS OrderTotal " & _
      "FROM Customers c LEFT JOIN Orders o ON o.CustomerID = c.CustomerID " & _
      "GROUP BY c.CustomerID, c.Name " & _
      "ORDER BY c.Name"

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open CONN_STR
Set rs = conn.Execute(sql)
data = Empty
If Not rs.EOF Then data = rs.GetRows()   ' columns: 0=Name 1=OrderCount 2=OrderTotal
rs.Close : Set rs = Nothing
conn.Close : Set conn = Nothing          ' database work finished here
%>
<!doctype html>
<html><head><meta charset="utf-8"><title>Customer report (FIXED)</title></head>
<body>
<h1>Customer report</h1>
<table border="1" cellpadding="6">
<tr><th>Customer</th><th>Orders</th><th>Total</th></tr>
<%
If IsArray(data) Then
    For i = 0 To UBound(data, 2)
        Response.Write "<tr><td>" & Server.HTMLEncode(data(0, i) & "") & "</td><td>" & data(1, i) & _
                       "</td><td>" & FormatNumber(data(2, i), 2) & "</td></tr>"
    Next
End If
%>
</table>
<p>1 query, <%= FormatNumber(Timer() - t0, 3) %> seconds</p>
</body></html>
