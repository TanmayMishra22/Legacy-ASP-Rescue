<%@ Language="VBScript" CodePage="65001" %>
<% Option Explicit %>
<%
' =====================================================================
' FIXED: one unambiguous format everywhere.
'  - <input type="date"> always submits yyyy-mm-dd, whatever the PC's locale
'  - parsed by hand with DateSerial (no CDate guessing) and validated
'  - sent to SQL Server as a typed date PARAMETER, never as text
'  - shown with an explicit format (3 Apr 2026), not the server's locale
' Demo only. Needs the RescueDemo database from ../assets/demo-db.sql
' =====================================================================
Response.Charset = "utf-8"
Const CONN_STR = "Provider=MSOLEDBSQL;Data Source=.\SQLEXPRESS;Initial Catalog=RescueDemo;Integrated Security=SSPI;"
Const adCmdText = 1
Const adParamInput = 1
Const adDBTimeStamp = 135

' "2026-04-03" -> Date value, or Null if invalid
Function ParseIsoDate(s)
    Dim re, parts, y, m, d, dt
    ParseIsoDate = Null
    Set re = New RegExp
    re.Pattern = "^\d{4}-\d{2}-\d{2}$"
    If Not re.Test(s & "") Then Exit Function
    parts = Split(s, "-")
    y = CInt(parts(0)) : m = CInt(parts(1)) : d = CInt(parts(2))
    If y < 1900 Or y > 2100 Or m < 1 Or m > 12 Or d < 1 Or d > 31 Then Exit Function
    dt = DateSerial(y, m, d)
    If Day(dt) <> d Then Exit Function          ' 2026-02-31 would roll into March
    ParseIsoDate = dt
End Function

Function IsoDate(dt)
    IsoDate = Year(dt) & "-" & Right("0" & Month(dt), 2) & "-" & Right("0" & Day(dt), 2)
End Function

' Fixed English month names - MonthName() follows the server locale.
Function NiceDate(dt)
    Dim names
    names = Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
    NiceDate = Day(dt) & " " & names(Month(dt) - 1) & " " & Year(dt)
End Function

Dim conn, cmd, rs, fromText, fromDate, errText
fromText = Trim(Request.QueryString("from") & "")
If Len(fromText) > 0 Then
    fromDate = ParseIsoDate(fromText)
    If IsNull(fromDate) Then errText = "Please pick a valid date."
End If
%>
<!doctype html>
<html><head><meta charset="utf-8"><title>Orders since (FIXED)</title></head>
<body>
<h1>Orders since a date</h1>
<form method="get">From: <input type="date" name="from" value="<%= Server.HTMLEncode(fromText) %>"> <button>Show</button></form>
<%
If Len(errText) > 0 Then
    Response.Write "<p style=""color:#b00020"">" & errText & "</p>"
ElseIf Len(fromText) > 0 Then
    Set conn = Server.CreateObject("ADODB.Connection")
    conn.Open CONN_STR
    Set cmd = Server.CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    cmd.CommandText = "SELECT OrderID, OrderDate, Total FROM Orders WHERE OrderDate >= ? ORDER BY OrderDate"
    cmd.Parameters.Append cmd.CreateParameter("from", adDBTimeStamp, adParamInput, , fromDate)
    Set rs = cmd.Execute

    Response.Write "<p>Showing orders since " & NiceDate(fromDate) & " (" & IsoDate(fromDate) & ")</p><ul>"
    Do While Not rs.EOF
        Response.Write "<li>#" & rs("OrderID") & " on " & NiceDate(rs("OrderDate")) & " - " & FormatNumber(rs("Total"), 2) & "</li>"
        rs.MoveNext
    Loop
    Response.Write "</ul>"
    rs.Close : Set rs = Nothing
    Set cmd = Nothing
    conn.Close : Set conn = Nothing
End If
%>
</body></html>
