<%@ Language="VBScript" CodePage="65001" %>
<% Option Explicit %>
<%
' =====================================================================
' FIXED: cart stored in SQL Server, keyed by a random token cookie.
' Survives app pool recycles, web gardens and load-balanced servers.
' Demo only. Needs the RescueDemo database from ../assets/demo-db.sql
' =====================================================================
Response.Charset = "utf-8"
Const CONN_STR = "Provider=MSOLEDBSQL;Data Source=.\SQLEXPRESS;Initial Catalog=RescueDemo;Integrated Security=SSPI;"
Const adCmdText = 1
Const adParamInput = 1
Const adInteger = 3
Const adChar = 129

Function IsValidToken(t)
    Dim re
    Set re = New RegExp
    re.Pattern = "^[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}$"
    re.IgnoreCase = True
    IsValidToken = re.Test(t & "")
End Function

Function NewToken(conn)
    Dim rs
    Set rs = conn.Execute("SELECT CONVERT(CHAR(36), NEWID()) AS T")
    NewToken = rs("T") & ""
    rs.Close : Set rs = Nothing
End Function

Function MakeCmd(conn, sql)
    Dim cmd
    Set cmd = Server.CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    cmd.CommandText = sql
    Set MakeCmd = cmd
End Function

Dim conn, cmd, rs, token, pid, found

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open CONN_STR

token = Request.Cookies("CartToken")
If Not IsValidToken(token) Then
    token = NewToken(conn)
    Response.Cookies("CartToken") = token
    Response.Cookies("CartToken").Expires = DateAdd("d", 30, Now())
    Response.Cookies("CartToken").Path = "/"
    ' On HTTPS sites also set: Response.Cookies("CartToken").Secure = True
End If

pid = Request.QueryString("add")
If Len(pid) > 0 And Len(pid) < 10 And IsNumeric(pid) Then
    Set cmd = MakeCmd(conn, _
        "UPDATE CartItems SET Qty = Qty + 1 WHERE CartToken = ? AND ProductID = ?; " & _
        "IF @@ROWCOUNT = 0 INSERT INTO CartItems (CartToken, ProductID, Qty) VALUES (?, ?, 1);")
    cmd.Parameters.Append cmd.CreateParameter("t1", adChar, adParamInput, 36, token)
    cmd.Parameters.Append cmd.CreateParameter("p1", adInteger, adParamInput, , CLng(pid))
    cmd.Parameters.Append cmd.CreateParameter("t2", adChar, adParamInput, 36, token)
    cmd.Parameters.Append cmd.CreateParameter("p2", adInteger, adParamInput, , CLng(pid))
    cmd.Execute
    Set cmd = Nothing
End If

Set cmd = MakeCmd(conn, "SELECT ProductID, Qty FROM CartItems WHERE CartToken = ? ORDER BY AddedAt")
cmd.Parameters.Append cmd.CreateParameter("t", adChar, adParamInput, 36, token)
Set rs = cmd.Execute
%>
<!doctype html>
<html><head><meta charset="utf-8"><title>Cart (FIXED)</title></head>
<body>
<h1>Your cart</h1>
<p><a href="?add=101">Add product 101</a> | <a href="?add=202">Add product 202</a> | <a href="?add=303">Add product 303</a></p>
<%
found = False
Response.Write "<ul>"
Do While Not rs.EOF
    found = True
    Response.Write "<li>Product " & CLng(rs("ProductID")) & " &times; " & CLng(rs("Qty")) & "</li>"
    rs.MoveNext
Loop
Response.Write "</ul>"
If Not found Then Response.Write "<p>Your cart is empty.</p>"

rs.Close : Set rs = Nothing
Set cmd = Nothing
conn.Close : Set conn = Nothing
%>
</body></html>
