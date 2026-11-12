<%@ Language="VBScript" CodePage="65001" %>
<% Option Explicit %>
<%
' =====================================================================
' FIXED: errors are caught only where expected, logged with details
'        and a reference number, and the visitor sees a polite message.
' Needs a writable folder: <site root>\App_Data  (IIS blocks browsing it)
' Give the app pool identity Modify permission on that folder.
' Demo only. Needs the RescueDemo database from ../assets/demo-db.sql
' =====================================================================
Response.Charset = "utf-8"
Response.Buffer = True
Const CONN_STR = "Provider=MSOLEDBSQL;Data Source=.\SQLEXPRESS;Initial Catalog=RescueDemo;Integrated Security=SSPI;"
Const ForAppending = 8

Sub WriteLog(lineText)
    Dim fso, f
    On Error Resume Next            ' logging must never break the page
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    Set f = fso.OpenTextFile(Server.MapPath("/App_Data/asp-errors.log"), ForAppending, True)
    f.WriteLine lineText
    f.Close
    Set f = Nothing : Set fso = Nothing
    On Error GoTo 0
End Sub

' Call straight after a failing line, while Err still holds the details.
Sub FailPage(stepName)
    Dim errNum, errSrc, errDesc, ref
    errNum = Err.Number : errSrc = Err.Source : errDesc = Err.Description
    Err.Clear
    On Error GoTo 0
    Randomize
    ref = Year(Now()) & Right("0" & Month(Now()), 2) & Right("0" & Day(Now()), 2) & "-" & _
          Right("0" & Hour(Now()), 2) & Right("0" & Minute(Now()), 2) & Right("0" & Second(Now()), 2) & "-" & _
          Int(Rnd() * 9000 + 1000)
    WriteLog ref & vbTab & Request.ServerVariables("URL") & "?" & Request.QueryString & vbTab & _
             stepName & vbTab & "0x" & Hex(errNum) & vbTab & errSrc & vbTab & errDesc
    Response.Clear
    Response.Status = "500 Internal Server Error"
    Response.Write "<!doctype html><html><head><meta charset=""utf-8""><title>Sorry</title></head><body>"
    Response.Write "<h1>Sorry, something went wrong.</h1>"
    Response.Write "<p>We have logged the problem. Please quote reference <strong>" & ref & "</strong>.</p>"
    Response.Write "</body></html>"
    Response.End
End Sub

Dim conn, rs, id, total

id = Request.QueryString("id")
If Not IsNumeric(id) Or Len(id) = 0 Or Len(id) > 9 Then
    Response.Write "<p>Please supply a numeric order id, e.g. ?id=1</p>"
    Response.End
End If

Set conn = Server.CreateObject("ADODB.Connection")

On Error Resume Next                       ' guard ONLY the risky lines
conn.Open CONN_STR
If Err.Number <> 0 Then FailPage "Opening database"
On Error Resume Next
Set rs = conn.Execute("SELECT Total FROM Orderz WHERE OrderID = " & CLng(id))   ' same typo, now visible in the log
If Err.Number <> 0 Then FailPage "Loading order total"
On Error GoTo 0                            ' back to normal: other bugs show up loudly

If rs.EOF Then
    total = Empty
Else
    total = rs("Total")
End If
rs.Close : Set rs = Nothing
conn.Close : Set conn = Nothing
%>
<!doctype html>
<html><head><meta charset="utf-8"><title>Order (FIXED)</title></head>
<body>
<h1>Order <%= CLng(id) %></h1>
<% If IsEmpty(total) Then %>
<p>No order with that number.</p>
<% Else %>
<p>Order total: <%= FormatNumber(total, 2) %></p>
<% End If %>
</body></html>
