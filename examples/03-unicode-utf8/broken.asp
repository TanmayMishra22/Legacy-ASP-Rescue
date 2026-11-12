<%@ Language="VBScript" %>
<% Option Explicit %>
<%
' =====================================================================
' BROKEN: garbled text (JosÃ© / ????). No code page, no charset,
'         and the value is sent to SQL Server as non-Unicode VARCHAR.
' Demo only. Needs the RescueDemo database from ../assets/demo-db.sql
' =====================================================================
Const CONN_STR = "Provider=MSOLEDBSQL;Data Source=.\SQLEXPRESS;Initial Catalog=RescueDemo;Integrated Security=SSPI;"
Const adCmdText = 1
Const adParamInput = 1
Const adVarChar = 129        ' <-- non-Unicode: Hindi, Chinese, emoji become "?"

Dim conn, cmd, rs, nm
nm = Request.Form("name")    ' decoded with the server's default ANSI code page

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open CONN_STR
If Len(nm) > 0 Then
    Set cmd = Server.CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    cmd.CommandText = "INSERT INTO Guestbook (Name) VALUES (?)"
    cmd.Parameters.Append cmd.CreateParameter("name", adVarChar, adParamInput, 200, nm)
    cmd.Execute
    Set cmd = Nothing
End If
Set rs = conn.Execute("SELECT TOP 10 Name FROM Guestbook ORDER BY EntryID DESC")
%>
<html><head><title>Guestbook (BROKEN)</title></head>
<body>
<h1>Guestbook - Namaste / नमस्ते</h1>
<form method="post"><input name="name"> <button>Sign</button></form>
<ul>
<% Do While Not rs.EOF %>
  <li><%= Server.HTMLEncode(rs("Name") & "") %></li>
<%   rs.MoveNext
   Loop
   rs.Close : Set rs = Nothing
   conn.Close : Set conn = Nothing %>
</ul>
</body></html>
