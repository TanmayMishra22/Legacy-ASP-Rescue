<%@ Language="VBScript" CodePage="65001" %>
<% Option Explicit %>
<%
' =====================================================================
' BROKEN: SQL injection - user input is glued straight into the SQL text
' Demo only. Needs the RescueDemo database from ../assets/demo-db.sql
' =====================================================================
Response.Charset = "utf-8"
Const CONN_STR = "Provider=MSOLEDBSQL;Data Source=.\SQLEXPRESS;Initial Catalog=RescueDemo;Integrated Security=SSPI;"

Dim conn, rs, sql, city
city = Request.QueryString("city")
%>
<!doctype html>
<html><head><meta charset="utf-8"><title>Customer search (BROKEN)</title></head>
<body>
<h1>Customers by city</h1>
<form method="get"><input name="city" value="<%= Server.HTMLEncode(city) %>"> <button>Search</button></form>
<%
If Len(city) > 0 Then
    Set conn = Server.CreateObject("ADODB.Connection")
    conn.Open CONN_STR

    ' DANGER: whatever the visitor types becomes part of the SQL command.
    ' Typing   ' OR '1'='1   returns every customer.
    ' Typing   '             breaks the page with "Unclosed quotation mark".
    sql = "SELECT CustomerID, Name, City FROM Customers WHERE City = '" & city & "'"
    Set rs = conn.Execute(sql)

    Response.Write "<ul>"
    Do While Not rs.EOF
        Response.Write "<li>" & Server.HTMLEncode(rs("Name") & "") & " (" & Server.HTMLEncode(rs("City") & "") & ")</li>"
        rs.MoveNext
    Loop
    Response.Write "</ul>"

    rs.Close : Set rs = Nothing
    conn.Close : Set conn = Nothing
End If
%>
</body></html>
