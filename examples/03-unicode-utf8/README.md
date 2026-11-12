# Classic ASP fix 03: Garbled text: JosÃ©, â€™ and ???? instead of real characters

Accents, Hindi and curly quotes break on the way in or out. Fix: UTF-8 end to end, NVARCHAR and N'' strings.

**Explainer page (live):** https://tanmaymishra22.github.io/Legacy-ASP-Rescue/examples/03-unicode-utf8/

| File | What it is |
|---|---|
| `broken.asp` | The typical broken version, runnable under IIS with Classic ASP enabled |
| `fixed.asp` | The corrected version |
| `index.html` | Plain-language explainer with side-by-side broken vs fixed comparison |

## Symptom
Names show up as *JosÃ© MÃ¼ller*, apostrophes as *â€™*, or Hindi/Chinese text is saved as *??????*. Often it starts after copying the site to a new server or after someone edits a file in a different editor.

## Why it happens
Text passes through four places, and each must agree on the encoding: the .asp file on disk, ASP's code page (how it reads form posts and writes output), the browser's charset, and the database column. Classic ASP defaults to the server's ANSI code page (for example 1252). If the file is UTF-8 but ASP thinks it is 1252, you get *Ã©*. If the column is VARCHAR or the parameter is `adVarChar`, characters outside the code page become *?* and are lost for good.

## How to spot it
- Line 1 of the page has no `CodePage="65001"`, and there is no `Response.Charset = "utf-8"`.
- Columns are `VARCHAR`/`TEXT` instead of `NVARCHAR`/`NVARCHAR(MAX)`.
- SQL literals without the `N` prefix: `'नमस्ते'` instead of `N'नमस्ते'`.
- Parameters created as `adVarChar` (129) instead of `adVarWChar` (202).
- Include files saved in a different encoding from the page that includes them.

## The fix
Make every step UTF-8: save all .asp and .inc files as UTF-8, put `CodePage="65001"` on line 1, set `Response.CodePage = 65001` and `Response.Charset = "utf-8"`, add ``, store text in NVARCHAR columns and send it with `adVarWChar` parameters (or `N'...'` literals). Text already saved as *?* cannot be recovered; text saved as *Ã©* often can be converted back.

## Broken (excerpt)
```vb
<%@ Language="VBScript" %>
<%
' no CodePage, no Charset -> server ANSI code page
Const adVarChar = 129    ' non-Unicode parameter

nm = Request.Form("name")
cmd.CommandText = "INSERT INTO Guestbook (Name) VALUES (?)"
cmd.Parameters.Append cmd.CreateParameter( _
    "name", adVarChar, adParamInput, 200, nm)
cmd.Execute
%>
```

## Fixed (excerpt)
```vb
<%@ Language="VBScript" CodePage="65001" %>
<%
Response.CodePage = 65001
Response.Charset  = "utf-8"
Const adVarWChar = 202   ' Unicode parameter (NVARCHAR)

nm = Trim(Request.Form("name") & "")
cmd.CommandText = "INSERT INTO Guestbook (Name) VALUES (?)"
cmd.Parameters.Append cmd.CreateParameter( _
    "name", adVarWChar, adParamInput, 200, nm)
cmd.Execute
%>
<meta charset="utf-8">   <!-- and Name is NVARCHAR(200) -->
```

## Checklist
- [ ] All .asp/.inc files saved as UTF-8 (same encoding everywhere).
- [ ] `` on line 1, plus `Response.Charset = "utf-8"`.
- [ ] `` in the HTML head.
- [ ] Text columns are NVARCHAR; literals use `N'...'`; parameters use `adVarWChar`.
- [ ] Tested with an accent (é), a curly quote (’) and a non-Latin script (नमस्ते).

## Running it
1. Enable Classic ASP in IIS and point a site or virtual directory at this repository.
2. Create the demo database with [`../assets/demo-db.sql`](../assets/demo-db.sql) (SQL Server / Express). Pages use the
   *Microsoft OLE DB Driver for SQL Server* (`Provider=MSOLEDBSQL`) and Windows authentication; give the app pool identity
   access to the `RescueDemo` database, or edit `CONN_STR`.
3. Open `broken.asp` and `fixed.asp` side by side.

Teaching example with made-up demo data. Not taken from any client system.
