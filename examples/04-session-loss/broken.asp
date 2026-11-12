<%@ Language="VBScript" CodePage="65001" %>
<% Option Explicit %>
<%
' =====================================================================
' BROKEN: shopping cart kept only in Session().
' Session data lives in the memory of ONE worker process. It is lost when:
'   - the app pool recycles or goes idle (default idle timeout 20 min)
'   - a web garden / load balancer sends the next request elsewhere
'   - Session.Timeout passes, or the cookie domain changes (www vs no-www)
' =====================================================================
Response.Charset = "utf-8"
Dim cart, pid, items, i

pid = Request.QueryString("add")
If Len(pid) > 0 And IsNumeric(pid) Then
    Session("Cart") = Session("Cart") & CLng(pid) & ","
End If

cart = Session("Cart") & ""
%>
<!doctype html>
<html><head><meta charset="utf-8"><title>Cart (BROKEN)</title></head>
<body>
<h1>Your cart</h1>
<p><a href="?add=101">Add product 101</a> | <a href="?add=202">Add product 202</a> | <a href="?add=303">Add product 303</a></p>
<%
If Len(cart) = 0 Then
    Response.Write "<p>Your cart is empty.</p>"
Else
    items = Split(Left(cart, Len(cart) - 1), ",")
    Response.Write "<ul>"
    For i = 0 To UBound(items)
        Response.Write "<li>Product " & items(i) & "</li>"
    Next
    Response.Write "</ul>"
End If
%>
</body></html>
