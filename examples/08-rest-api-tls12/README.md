# Classic ASP fix 08: Calling modern REST/JSON APIs: TLS 1.2 errors and fragile parsing

Payment, SMS or shipping API calls fail with secure-channel errors or break on small changes. Fix: ServerXMLHTTP 6.0, TLS 1.2, timeouts, status checks.

**Explainer page (live):** https://tanmaymishra22.github.io/Legacy-ASP-Rescue/examples/08-rest-api-tls12/

| File | What it is |
|---|---|
| `broken.asp` | The typical broken version, runnable under IIS with Classic ASP enabled |
| `fixed.asp` | The corrected version |
| `index.html` | Plain-language explainer with side-by-side broken vs fixed comparison |

## Symptom
An integration that worked for years suddenly fails: *msxml3.dll error '80072f7d' An error occurred in the secure channel support*, or *msxml6.dll error '80072efd'*/*'80072ee2' The operation timed out*. Or the call "works" but the page shows half a value or nothing, because the API reply changed slightly.

## Why it happens
API providers now require TLS 1.2 or newer. Old code often uses `MSXML2.ServerXMLHTTP` (version 3.0) or the client-only `Microsoft.XMLHTTP`, and older Windows Server builds do not offer TLS 1.2 to WinHTTP by default. Missing timeouts let one slow API hang the page. Many APIs also reject requests without a `User-Agent` (GitHub returns 403). Picking values out of JSON with `InStr`/`Mid` breaks as soon as spacing, order or length changes.

## How to spot it
- `CreateObject("Microsoft.XMLHTTP")` or `"MSXML2.ServerXMLHTTP"` without `.6.0`.
- No `setTimeouts`, no check of `http.Status`, no `On Error` guard around `Send`.
- `SetOption 2, 13056` (ignore certificate errors) - a security hole, not a fix.
- Windows Server 2008 R2 / 2012 without the WinHTTP TLS 1.2 update and registry setting.

## The fix
In code: use `MSXML2.ServerXMLHTTP.6.0`, set timeouts, send `User-Agent` and `Accept` headers, guard `Send` and report the error number, check `Status` before reading the body, and read JSON values with a proper pattern or a proven JSON library. On the server: make sure TLS 1.2 is enabled in SCHANNEL and for WinHTTP (on 2008 R2/2012 install the WinHTTP TLS update KB3140245 and set `DefaultSecureProtocols` = `0xA00` under `HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\WinHttp`, and the `Wow6432Node` copy for 32-bit app pools), then reboot. Newer Windows versions support TLS 1.2 out of the box.

## Broken (excerpt)
```vb
Set http = Server.CreateObject("MSXML2.ServerXMLHTTP")  ' v3.0
http.Open "GET", "https://api.github.com/repos/" & _
          "TanmayMishra22/Legacy-ASP-Rescue", False
http.Send                    ' no timeout, no User-Agent
body = http.responseText     ' status never checked

p = InStr(body, """full_name"":")
repoName = Mid(body, p + 13, 40)   ' guesses the length
```

## Fixed (excerpt)
```vb
Set http = Server.CreateObject("MSXML2.ServerXMLHTTP.6.0")
http.setTimeouts 5000, 5000, 10000, 15000
On Error Resume Next
http.Open "GET", API_URL, False
http.setRequestHeader "User-Agent", "LegacyAspRescue-Demo/1.0"
http.setRequestHeader "Accept", "application/json"
http.Send
If Err.Number <> 0 Then
    errText = "Could not reach the API (0x" & Hex(Err.Number) & ")"
Else
    status = http.Status : body = http.responseText
End If
On Error GoTo 0
If status = 200 Then repoName = JsonString(body, "full_name")
```

## Checklist
- [ ] `MSXML2.ServerXMLHTTP.6.0` everywhere on the server side.
- [ ] `setTimeouts` set; transport errors caught and logged with their hex code.
- [ ] `Status` checked; unexpected replies logged (first part of the body).
- [ ] Certificate validation left ON (no 13056 option).
- [ ] Server confirmed to negotiate TLS 1.2 (test against the provider's endpoint after changes).

## Running it
1. Enable Classic ASP in IIS and point a site or virtual directory at this repository.
2. Create the demo database with [`../assets/demo-db.sql`](../assets/demo-db.sql) (SQL Server / Express). Pages use the
   *Microsoft OLE DB Driver for SQL Server* (`Provider=MSOLEDBSQL`) and Windows authentication; give the app pool identity
   access to the `RescueDemo` database, or edit `CONN_STR`.
3. Open `broken.asp` and `fixed.asp` side by side.

Teaching example with made-up demo data. Not taken from any client system.
