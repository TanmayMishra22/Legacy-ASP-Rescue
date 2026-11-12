# Classic ASP fix 05: "500 - Internal server error" that hides the real problem

Generic 500 pages and On Error Resume Next hide what actually failed. Fix: see real errors safely, log them with a reference.

**Explainer page (live):** https://tanmaymishra22.github.io/Legacy-ASP-Rescue/examples/05-hidden-500-errors/

| File | What it is |
|---|---|
| `broken.asp` | The typical broken version, runnable under IIS with Classic ASP enabled |
| `fixed.asp` | The corrected version |
| `index.html` | Plain-language explainer with side-by-side broken vs fixed comparison |

## Symptom
Visitors get *500 - Internal server error. There is a problem with the resource you are looking for, and it cannot be displayed.* Or worse: the page loads but a value is blank, an email is never sent, an order is half saved, and there is no error at all.

## Why it happens
For security, IIS hides ASP error details from remote visitors, so the useful message (file, line, description) never reaches anyone. Many old pages also start with `On Error Resume Next`, which tells VBScript to skip every failing line silently for the rest of the page. Without logging, the cause is invisible.

## How to spot it
- `On Error Resume Next` near the top of a page or include, with no `If Err.Number <> 0` checks after risky lines.
- IIS log (`%SystemDrive%\inetpub\logs\LogFiles`) shows status 500 with sub-status and, for ASP, the line number and error text in the query field.
- Browsing the page *on the server itself* shows the detailed error that remote users never see.

## The fix
On a test server or for local requests, turn on detailed errors: in IIS Manager open ASP > Debugging Properties > *Send Errors To Browser* = True, and Error Pages > Edit Feature Settings > *Detailed errors for local requests*. In code, use `On Error Resume Next` only around a risky line, check `Err.Number` straight after, write the details to a log file outside the web folder, show the visitor a polite message with a reference number, and switch back with `On Error GoTo 0`.

## Broken (excerpt)
```vb
On Error Resume Next      ' hides EVERY error below

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open CONN_STR
Set rs = conn.Execute("SELECT Total FROM Orderz " & _
                      "WHERE OrderID = " & CLng(id))
%>
<p>Order total: <%= rs("Total") %></p>  <!-- blank -->
```

## Fixed (excerpt)
```vb
On Error Resume Next                   ' guard one step
Set rs = conn.Execute("SELECT Total FROM Orderz " & _
                      "WHERE OrderID = " & CLng(id))
If Err.Number <> 0 Then FailPage "Loading order total"
On Error GoTo 0                        ' back to normal

' FailPage: copies Err details, writes a line to
' /App_Data/asp-errors.log, sets status 500 and shows
' "Sorry, something went wrong. Reference 20261112-..."
```

## Checklist
- [ ] No page-wide `On Error Resume Next`; error checks sit right after risky lines.
- [ ] Errors are logged with time, URL, step, number, source and description.
- [ ] Visitors see a friendly message with a reference, never stack details.
- [ ] Detailed errors are on for local requests only (`appcmd set config -section:httpErrors -errorMode:DetailedLocalOnly`).
- [ ] Log folder (e.g. App_Data) is not browsable and the app pool can write to it.

## Running it
1. Enable Classic ASP in IIS and point a site or virtual directory at this repository.
2. Create the demo database with [`../assets/demo-db.sql`](../assets/demo-db.sql) (SQL Server / Express). Pages use the
   *Microsoft OLE DB Driver for SQL Server* (`Provider=MSOLEDBSQL`) and Windows authentication; give the app pool identity
   access to the `RescueDemo` database, or edit `CONN_STR`.
3. Open `broken.asp` and `fixed.asp` side by side.

Teaching example with made-up demo data. Not taken from any client system.
