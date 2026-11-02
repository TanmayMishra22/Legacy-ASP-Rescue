<%
Option Explicit

Function ApiGet(url, bearerToken, timeoutMs)
    Dim http, result
    Set http = Server.CreateObject("MSXML2.ServerXMLHTTP.6.0")

    On Error Resume Next

    http.setTimeouts timeoutMs, timeoutMs, timeoutMs, timeoutMs
    http.Open "GET", url, False

    If Len(bearerToken) > 0 Then
        http.setRequestHeader "Authorization", "Bearer " & bearerToken
    End If

    http.setRequestHeader "Accept", "application/json"
    http.Send

    If Err.Number <> 0 Then
        result = "TRANSPORT_ERROR|number=" & Err.Number &
                 "|description=" & Err.Description
        Err.Clear
        On Error GoTo 0
        ApiGet = result
        Exit Function
    End If

    On Error GoTo 0

    result = "HTTP_STATUS|" & CStr(http.Status) &
             "|STATUS_TEXT|" & http.StatusText &
             "|BODY|" & http.ResponseText

    ApiGet = result
    Set http = Nothing
End Function

Function DiagnoseApiResult(result)
    Dim parts, statusCode

    If Left(result, 15) = "TRANSPORT_ERROR" Then
        DiagnoseApiResult = "Transport failure: check DNS, network, firewall, TLS and endpoint availability."
        Exit Function
    End If

    parts = Split(result, "|BODY|")
    statusCode = Replace(parts(0), "HTTP_STATUS|", "")

    Select Case CInt(statusCode)
        Case 200 To 299
            DiagnoseApiResult = "HTTP success. Validate the response body and expected JSON/data structure."
        Case 400
            DiagnoseApiResult = "HTTP 400: check parameters, request format and required fields."
        Case 401, 403
            DiagnoseApiResult = "Authentication/authorization failure. Check token, scopes and permissions."
        Case 404
            DiagnoseApiResult = "Endpoint not found. Check URL, API version and route."
        Case 429
            DiagnoseApiResult = "Rate limited. Check API limits and retry behavior."
        Case 500 To 599
            DiagnoseApiResult = "Remote server failure. Inspect the response before changing application code."
        Case Else
            DiagnoseApiResult = "Unexpected HTTP status " & statusCode & ". Inspect headers and response body."
    End Select
End Function

' Demo
Dim apiUrl, token, result

apiUrl = "https://api.example.com/v1/example"
token = ""  ' Never hard-code a real credential.

result = ApiGet(apiUrl, token, 10000)

Response.ContentType = "text/plain"
Response.Write "RAW RESULT:" & vbCrLf
Response.Write result & vbCrLf & vbCrLf
Response.Write "DIAGNOSIS:" & vbCrLf
Response.Write DiagnoseApiResult(result) & vbCrLf
%>
