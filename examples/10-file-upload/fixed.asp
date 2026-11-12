<%@ Language="VBScript" CodePage="65001" %>
<% Option Explicit %>
<%
' =====================================================================
' FIXED: a small, safe single-file upload handler in pure Classic ASP.
'  - reads the body ONCE with BinaryRead (never touches Request.Form)
'  - checks the size first and gives a clear message
'  - finds the file part by its field name and copies ONLY the file bytes
'  - allows only listed extensions, ignores the visitor's file name and
'    saves to a folder that is NOT web-browsable (App_Data\uploads)
' Server setup (see README): raise ASP maxRequestEntityAllowed above
' MAX_BYTES, and give the app pool identity Modify on App_Data\uploads.
' =====================================================================
Response.Charset = "utf-8"

Const MAX_BYTES = 5242880              ' 5 MB
Const FILE_FIELD = "photo"
Const adTypeBinary = 1
Const adSaveCreateOverWrite = 2

' ASCII text -> byte string (for InStrB searches in the raw body)
Function ToBytes(s)
    Dim i, r
    r = ""
    For i = 1 To Len(s)
        r = r & ChrB(AscB(Mid(s, i, 1)))
    Next
    ToBytes = r
End Function

' byte string -> text (headers only, which are plain ASCII)
Function BytesToText(b)
    Dim i, r
    r = ""
    For i = 1 To LenB(b)
        r = r & Chr(AscB(MidB(b, i, 1)))
    Next
    BytesToText = r
End Function

Function GetBoundary()
    Dim ct, p, b
    GetBoundary = ""
    ct = Request.ServerVariables("HTTP_CONTENT_TYPE")
    If InStr(1, ct, "multipart/form-data", vbTextCompare) = 0 Then Exit Function
    p = InStr(1, ct, "boundary=", vbTextCompare)
    If p = 0 Then Exit Function
    b = Mid(ct, p + 9)
    If InStr(b, ";") > 0 Then b = Left(b, InStr(b, ";") - 1)
    b = Replace(Trim(b), """", "")
    GetBoundary = b
End Function

' Last occurrence of needle in haystack at or before position limit (byte strings)
Function FindLastB(haystack, needle, limit)
    Dim p, lastFound
    lastFound = 0
    p = InStrB(1, haystack, needle)
    Do While p > 0 And p < limit
        lastFound = p
        p = InStrB(p + 1, haystack, needle)
    Loop
    FindLastB = lastFound
End Function

Function AllowedExtension(fileName)
    Dim ext, p
    AllowedExtension = ""
    p = InStrRev(fileName, ".")
    If p = 0 Then Exit Function
    ext = LCase(Mid(fileName, p + 1))
    Select Case ext
        Case "jpg", "jpeg", "png", "gif", "pdf": AllowedExtension = ext
    End Select
End Function

Dim msg, total, boundary, body, fieldPos, partStart, hdrEnd, dataStart, dataEnd
Dim headerText, re, m, origName, ext, safeName, src, dst

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    total = Request.TotalBytes
    boundary = GetBoundary()
    If total = 0 Or Len(boundary) = 0 Then
        msg = "No file was sent. The form needs enctype=""multipart/form-data""."
    ElseIf total > MAX_BYTES Then
        msg = "That file is too large. The limit is " & (MAX_BYTES \ 1048576) & " MB."
    Else
        body = Request.BinaryRead(total)
        fieldPos = InStrB(1, body, ToBytes("name=""" & FILE_FIELD & """"))
        If fieldPos = 0 Then
            msg = "The upload field was not found."
        Else
            partStart = FindLastB(body, ToBytes("--" & boundary), fieldPos)
            hdrEnd = InStrB(fieldPos, body, ToBytes(vbCrLf & vbCrLf))
            If partStart = 0 Or hdrEnd = 0 Then
                msg = "The upload could not be read."
            Else
                headerText = BytesToText(MidB(body, partStart, hdrEnd - partStart))
                dataStart = hdrEnd + 4
                dataEnd = InStrB(dataStart, body, ToBytes(vbCrLf & "--" & boundary))
                Set re = New RegExp
                re.Pattern = "filename=""([^""]*)"""
                re.IgnoreCase = True
                Set m = re.Execute(headerText)
                origName = ""
                If m.Count > 0 Then origName = m(0).SubMatches(0)
                origName = Mid(origName, InStrRev(origName, "\") + 1)   ' old IE sends C:\path\file.jpg
                ext = AllowedExtension(origName)

                If Len(origName) = 0 Or dataEnd <= dataStart Then
                    msg = "Please choose a file."
                ElseIf Len(ext) = 0 Then
                    msg = "Only JPG, PNG, GIF or PDF files are allowed."
                Else
                    Randomize
                    safeName = Year(Now()) & Right("0" & Month(Now()), 2) & Right("0" & Day(Now()), 2) & "_" & _
                               Right("0" & Hour(Now()), 2) & Right("0" & Minute(Now()), 2) & Right("0" & Second(Now()), 2) & "_" & _
                               Int(Rnd() * 900000 + 100000) & "." & ext

                    Set src = Server.CreateObject("ADODB.Stream")
                    src.Type = adTypeBinary
                    src.Open
                    src.Write body
                    src.Position = dataStart - 1          ' Stream positions are 0-based

                    Set dst = Server.CreateObject("ADODB.Stream")
                    dst.Type = adTypeBinary
                    dst.Open
                    src.CopyTo dst, dataEnd - dataStart    ' ONLY the file bytes
                    dst.SaveToFile Server.MapPath("/App_Data/uploads/" & safeName), adSaveCreateOverWrite
                    dst.Close : Set dst = Nothing
                    src.Close : Set src = Nothing

                    msg = "Uploaded " & Server.HTMLEncode(origName) & " (" & FormatNumber(dataEnd - dataStart, 0) & _
                          " bytes), saved as " & safeName & "."
                End If
            End If
        End If
    End If
End If

%>
<!doctype html>
<html><head><meta charset="utf-8"><title>Upload (FIXED)</title></head>
<body>
<h1>Upload a photo</h1>
<% If Len(msg) > 0 Then Response.Write "<p><strong>" & msg & "</strong></p>" %>
<form method="post" enctype="multipart/form-data">
  <input type="file" name="photo" accept=".jpg,.jpeg,.png,.gif,.pdf"> <button>Upload</button>
</form>
<p>Max <%= MAX_BYTES \ 1048576 %> MB. JPG, PNG, GIF or PDF.</p>
</body></html>
