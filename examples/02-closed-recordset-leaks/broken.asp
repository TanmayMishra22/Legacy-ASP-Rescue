<%@ Language="VBScript" CodePage="65001" %>
<% Option Explicit %>
<%
' =====================================================================
' BROKEN: "Operation is not allowed when the object is closed"
'         plus connections that are never closed (leaks).
' Demo only. Needs the RescueDemo database from ../assets/demo-db.sql
' =====================================================================
Response.Charset = "utf-8"
Const CONN_STR = "Provider=MSOLEDBSQL;Data Source=.\SQLEXPRESS;Initial Catalog=RescueDemo;Integrated Security=SSPI;"

Function GetCustomers()
    Dim conn, rs
    Set conn = Server.CreateObject("ADODB.Connection")
    conn.Open CONN_STR
    Set rs = conn.Execute("SELECT Name, City FROM Customers ORDER BY Name")
    conn.Close                 ' Closing the connection also closes the recordset...
    Set GetCustomers = rs      ' ...so the caller receives a dead recordset.
End Function

Function CountOrders()
    Dim conn, rs
    Set conn = Server.CreateObject("ADODB.Connection")
    conn.Open CONN_STR
    Set rs = conn.Execute("SELECT COUNT(*) AS N FROM Orders")
    CountOrders = rs("N")
    ' LEAK: rs and conn are never closed. Under load the pool runs dry and
    ' pages start timing out ("Timeout expired" / server getting slower).
End Function

Dim rs
%>
<!doctype html>
<html><head><meta charset="utf-8"><title>Customers (BROKEN)</title></head>
<body>
<h1>Customers</h1>
<p>Total orders: <%= CountOrders() %></p>
<ul>
<%
Set rs = GetCustomers()
Do While Not rs.EOF          ' <-- ADODB.Recordset error '800a0e78'
    Response.Write "<li>" & Server.HTMLEncode(rs("Name") & "") & "</li>"
    rs.MoveNext
Loop
%>
</ul>
</body></html>
