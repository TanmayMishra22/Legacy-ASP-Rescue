<%@ Language="VBScript" CodePage="65001" %>
<% Option Explicit %>
<%
' =====================================================================
' FIXED: UTF-8 end to end.
'  1. File saved as UTF-8 and CodePage="65001" on line 1
'  2. Response.CodePage + Response.Charset tell the browser it is UTF-8
'  3. NVARCHAR column + adVarWChar parameter (or N'...' literals)
' Demo only. Needs the RescueDemo database from ../assets/demo-db.sql
' =====================================================================
Response.CodePage = 65001
Response.Charset = "utf-8"

Const CONN_STR = "Provider=MSOLEDBSQL;Data Source=.\SQLEXPRESS;Initial Catalog=RescueDemo;Integrated Security=SSPI;"
Const adCmdText = 1
Const adParamInput = 1
Const adVarWChar = 202       ' Unicode string parameter (matches NVARCHAR)

Dim conn, cmd, rs, nm
nm = Trim(Request.Form("name") & "")   ' decoded as UTF-8 thanks to CodePage 65001
If Len(nm) > 200 Then nm = Left(nm, 200)

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open CONN_STR
If Len(nm) > 0 Then
    Set cmd = Server.CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    cmd.CommandText = "INSERT INTO Guestbook (Name) VALUES (?)"
    cmd.Parameters.Append cmd.CreateParameter("name", adVarWChar, adParamInput, 200, nm)
    cmd.Execute
    Set cmd = Nothing
End If
Set rs = conn.Execute("SELECT TOP 10 Name FROM Guestbook ORDER BY EntryID DESC")
%>
<!doctype html>
<html><head><meta charset="utf-8"><title>Guestbook (FIXED)</title></head>
<body>
<h1>Guestbook - Namaste / नमस्ते</h1>
<form method="post" accept-charset="utf-8"><input name="name" maxlength="200"> <button>Sign</button></form>
<ul>
<% Do While Not rs.EOF %>
  <li><%= Server.HTMLEncode(rs("Name") & "") %></li>
<%   rs.MoveNext
   Loop
   rs.Close : Set rs = Nothing
   conn.Close : Set conn = Nothing %>
</ul>
</body></html>
