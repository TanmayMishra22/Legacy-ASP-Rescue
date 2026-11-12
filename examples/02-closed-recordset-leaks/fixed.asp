<%@ Language="VBScript" CodePage="65001" %>
<% Option Explicit %>
<%
' =====================================================================
' FIXED: copy the rows into an array with GetRows, then close and
'        release every ADO object in the same function that opened it.
' Demo only. Needs the RescueDemo database from ../assets/demo-db.sql
' =====================================================================
Response.Charset = "utf-8"
Const CONN_STR = "Provider=MSOLEDBSQL;Data Source=.\SQLEXPRESS;Initial Catalog=RescueDemo;Integrated Security=SSPI;"

' Returns a 2-D array (column, row) or Empty when there are no rows.
Function GetCustomers()
    Dim conn, rs
    GetCustomers = Empty
    Set conn = Server.CreateObject("ADODB.Connection")
    conn.Open CONN_STR
    Set rs = conn.Execute("SELECT Name, City FROM Customers ORDER BY Name")
    If Not rs.EOF Then GetCustomers = rs.GetRows()
    rs.Close   : Set rs = Nothing
    conn.Close : Set conn = Nothing
End Function

Function CountOrders()
    Dim conn, rs
    Set conn = Server.CreateObject("ADODB.Connection")
    conn.Open CONN_STR
    Set rs = conn.Execute("SELECT COUNT(*) AS N FROM Orders")
    CountOrders = CLng(rs("N"))    ' copy the value out BEFORE closing
    rs.Close   : Set rs = Nothing
    conn.Close : Set conn = Nothing  ' returns the connection to the pool
End Function

Dim data, i
data = GetCustomers()
%>
<!doctype html>
<html><head><meta charset="utf-8"><title>Customers (FIXED)</title></head>
<body>
<h1>Customers</h1>
<p>Total orders: <%= CountOrders() %></p>
<%
If IsArray(data) Then
    Response.Write "<ul>"
    For i = 0 To UBound(data, 2)
        Response.Write "<li>" & Server.HTMLEncode(data(0, i) & "") & " (" & Server.HTMLEncode(data(1, i) & "") & ")</li>"
    Next
    Response.Write "</ul>"
Else
    Response.Write "<p>No customers yet.</p>"
End If
%>
</body></html>
