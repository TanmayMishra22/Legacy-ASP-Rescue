<%@ Language="VBScript" CodePage="65001" %>
<% Option Explicit %>
<%
' =====================================================================
' BROKEN: dates depend on whoever configured the server.
'  - a free-text "dd/mm/yyyy" box, read with CDate(): 03/04/2026 becomes
'    4 March on a US-locale server (LCID 1033) but 3 April on an Indian/UK one
'  - the date is pasted into SQL as text: 13/04/2026 ->
'    "The conversion of a varchar data type to a datetime data type
'     resulted in an out-of-range value."
'  - output uses the server's regional format, which changes after a migration
' =====================================================================
Response.Charset = "utf-8"
Const CONN_STR = "Provider=MSOLEDBSQL;Data Source=.\SQLEXPRESS;Initial Catalog=RescueDemo;Integrated Security=SSPI;"

Dim conn, rs, fromText, d
fromText = Request.QueryString("from")      ' visitor types 03/04/2026 meaning 3 April
%>
<!doctype html>
<html><head><meta charset="utf-8"><title>Orders since (BROKEN)</title></head>
<body>
<h1>Orders since a date</h1>
<form method="get">From (dd/mm/yyyy): <input name="from" value="<%= Server.HTMLEncode(fromText) %>"> <button>Show</button></form>
<%
If Len(fromText) > 0 Then
    d = CDate(fromText)                       ' meaning depends on Session.LCID / server locale
    Set conn = Server.CreateObject("ADODB.Connection")
    conn.Open CONN_STR
    Set rs = conn.Execute("SELECT OrderID, OrderDate, Total FROM Orders " & _
                          "WHERE OrderDate >= '" & fromText & "' ORDER BY OrderDate")
    Response.Write "<p>Showing orders since " & d & "</p><ul>"      ' regional format
    Do While Not rs.EOF
        Response.Write "<li>#" & rs("OrderID") & " on " & rs("OrderDate") & " - " & rs("Total") & "</li>"
        rs.MoveNext
    Loop
    Response.Write "</ul>"
    rs.Close : Set rs = Nothing
    conn.Close : Set conn = Nothing
End If
%>
</body></html>
