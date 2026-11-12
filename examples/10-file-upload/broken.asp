<%@ Language="VBScript" CodePage="65001" %>
<%
' =====================================================================
' BROKEN: typical failing upload page.
'  - Request.Form cannot read multipart/form-data, so title is empty, and
'    touching Request.Form first makes BinaryRead fail:
'    "Request object error 'ASP 0206 : 80004005' Cannot call BinaryRead"
'  - files over 200,000 bytes (IIS default maxRequestEntityAllowed) fail with
'    "Request object error 'ASP 0104 : 80004005' Operation not Allowed"
'  - saves the WHOLE request (multipart headers included) under a name the
'    visitor controls, inside the public web folder: corrupt files + a way
'    to upload "shell.asp" and run code on the server
' =====================================================================
Dim title, data, stm
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    title = Request.Form("title")                  ' empty, and blocks BinaryRead below
    data = Request.BinaryRead(Request.TotalBytes)  ' ASP 0206 error here
    Set stm = Server.CreateObject("ADODB.Stream")
    stm.Type = 1
    stm.Open
    stm.Write data
    stm.SaveToFile Server.MapPath("uploads/" & Request.QueryString("name")), 2
    stm.Close
    Response.Write "Saved!"
End If
%>
<!doctype html>
<html><head><meta charset="utf-8"><title>Upload (BROKEN)</title></head>
<body>
<h1>Upload a photo</h1>
<form method="post" action="broken.asp?name=photo.jpg" enctype="multipart/form-data">
  Title: <input name="title"><br>
  <input type="file" name="photo"> <button>Upload</button>
</form>
</body></html>
