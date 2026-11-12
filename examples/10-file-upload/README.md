# Classic ASP fix 10: File uploads that fail, corrupt files or open a security hole

ASP 0104 / ASP 0206 errors, broken files and unsafe saving. Fix: correct limits, one BinaryRead, safe names and folders.

**Explainer page (live):** https://tanmaymishra22.github.io/Legacy-ASP-Rescue/examples/10-file-upload/

| File | What it is |
|---|---|
| `broken.asp` | The typical broken version, runnable under IIS with Classic ASP enabled |
| `fixed.asp` | The corrected version |
| `index.html` | Plain-language explainer with side-by-side broken vs fixed comparison |

## Symptom
Small files upload but larger ones fail with *Request object error 'ASP 0104 : 80004005' Operation not Allowed*, or the page dies with *ASP 0206 Cannot call BinaryRead*. Saved files will not open (corrupted), or after a migration uploads fail with *Write to file failed* (permissions).

## Why it happens
IIS limits ASP request bodies to 200,000 bytes by default (`maxRequestEntityAllowed`), and Request Filtering has its own limit (`maxAllowedContentLength`, about 28.6 MB, which returns 404.13). `Request.Form` cannot read `multipart/form-data`, and once it has been touched, `Request.BinaryRead` is no longer allowed. Saving the whole body writes the multipart headers into the file, which corrupts it. Using the visitor's file name inside the web folder lets someone upload *shell.asp* and run it.

## How to spot it
- Upload pages that use both `Request.Form` and `Request.BinaryRead`.
- IIS Manager > site > ASP > Limits Properties > *Maximum Requesting Entity Body Limit* still 200000.
- Files saved with the original name or extension, into a folder that can be browsed.
- App pool identity (e.g. `IIS AppPool\YourPool`) has no Modify permission on the upload folder.

## The fix
Raise the limits on purpose: `appcmd set config "Default Web Site" -section:system.webServer/asp /limits.maxRequestEntityAllowed:10485760 /commit:apphost` (and Request Filtering if files exceed 28.6 MB; on IIS 6 the metabase setting is `AspMaxRequestEntityAllowed`). In code, read the body once with `BinaryRead`, find the file part, copy only its bytes with `ADODB.Stream`, allow only listed extensions, generate your own file name and save outside the web root (e.g. `App_Data\uploads`), serving files back through a script that checks access.

## Broken (excerpt)
```vb
title = Request.Form("title")        ' blocks BinaryRead
data = Request.BinaryRead(Request.TotalBytes)   ' ASP 0206

Set stm = Server.CreateObject("ADODB.Stream")
stm.Type = 1 : stm.Open
stm.Write data                       ' whole body = corrupt file
stm.SaveToFile Server.MapPath("uploads/" & _
               Request.QueryString("name")), 2    ' visitor picks name
```

## Fixed (excerpt)
```vb
If total > MAX_BYTES Then msg = "That file is too large." ...
body = Request.BinaryRead(total)     ' read once, no Request.Form
fieldPos  = InStrB(1, body, ToBytes("name=""photo"""))
hdrEnd    = InStrB(fieldPos, body, ToBytes(vbCrLf & vbCrLf))
dataStart = hdrEnd + 4
dataEnd   = InStrB(dataStart, body, ToBytes(vbCrLf & "--" & boundary))
ext = AllowedExtension(origName)     ' jpg, png, gif, pdf only
safeName = <date>_<time>_<random> & "." & ext

src.Write body : src.Position = dataStart - 1
src.CopyTo dst, dataEnd - dataStart  ' ONLY the file bytes
dst.SaveToFile Server.MapPath("/App_Data/uploads/" & safeName), 2
```

## Checklist
- [ ] ASP `maxRequestEntityAllowed` (and Request Filtering) set just above your real maximum.
- [ ] The page reads the body once with `BinaryRead`; no `Request.Form` on that page.
- [ ] Size is checked first, with a clear message.
- [ ] Extension whitelist; server-generated file name; never the visitor's path.
- [ ] Upload folder is outside the web root (or not executable) and the app pool has Modify on it.

## Running it
1. Enable Classic ASP in IIS and point a site or virtual directory at this repository.
2. Create the demo database with [`../assets/demo-db.sql`](../assets/demo-db.sql) (SQL Server / Express). Pages use the
   *Microsoft OLE DB Driver for SQL Server* (`Provider=MSOLEDBSQL`) and Windows authentication; give the app pool identity
   access to the `RescueDemo` database, or edit `CONN_STR`.
3. Open `broken.asp` and `fixed.asp` side by side.

Teaching example with made-up demo data. Not taken from any client system.
