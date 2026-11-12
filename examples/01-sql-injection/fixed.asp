<%@ Language="VBScript" CodePage="65001" %>
<% Option Explicit %>
<%
' =====================================================================
' FIXED: parameterised query with ADODB.Command
' The visitor's text is sent as DATA, never as part of the SQL command.
' Demo only. Needs the RescueDemo database from ../assets/demo-db.sql
' =====================================================================
Response.Charset = "utf-8"
Const CONN_STR = "Provider=MSOLEDBSQL;Data Source=.\SQLEXPRESS;Initial Catalog=RescueDemo;Integrated Security=SSPI;"

' ADO constants (or include adovbs.inc / use a METADATA typelib tag)
Const adCmdText = 1
Const adParamInput = 1
Const adVarWChar = 202

Dim conn, cmd, rs, city, found
city = Trim(Request.QueryString("city") & "")
If Len(city) > 100 Then city = Left(city, 100)   ' match the column size
%>
<!doctype html>
<html><head><meta charset="utf-8"><title>Customer search (FIXED)</title></head>
<body>
<h1>Customers by city</h1>
<form method="get"><input name="city" maxlength="100" value="<%= Server.HTMLEncode(city) %>"> <button>Search</button></form>
<%
If Len(city) > 0 Then
    Set conn = Server.CreateObject("ADODB.Connection")
    conn.Open CONN_STR

    Set cmd = Server.CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    cmd.CommandText = "SELECT CustomerID, Name, City FROM Customers WHERE City = ?"
    cmd.Parameters.Append cmd.CreateParameter("city", adVarWChar, adParamInput, 100, city)

    Set rs = cmd.Execute

    found = False
    Response.Write "<ul>"
    Do While Not rs.EOF
        found = True
        Response.Write "<li>" & Server.HTMLEncode(rs("Name") & "") & " (" & Server.HTMLEncode(rs("City") & "") & ")</li>"
        rs.MoveNext
    Loop
    Response.Write "</ul>"
    If Not found Then Response.Write "<p>No customers found in city: " & Server.HTMLEncode(city) & "</p>"

    rs.Close : Set rs = Nothing
    Set cmd = Nothing
    conn.Close : Set conn = Nothing
End If
%>
</body></html>
