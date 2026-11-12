# Classic ASP fix 09: Date bugs: 03/04 read as 4 March, and "out-of-range" datetime errors

Dates change meaning with server locale and break SQL. Fix: yyyy-mm-dd input, DateSerial parsing, typed date parameters.

**Explainer page (live):** https://tanmaymishra22.github.io/Legacy-ASP-Rescue/examples/09-dates-locale/

| File | What it is |
|---|---|
| `broken.asp` | The typical broken version, runnable under IIS with Classic ASP enabled |
| `fixed.asp` | The corrected version |
| `index.html` | Plain-language explainer with side-by-side broken vs fixed comparison |

## Symptom
Reports show the wrong month, bookings land on the wrong day, or days 13-31 crash with *The conversion of a varchar data type to a datetime data type resulted in an out-of-range value.* Often appears after moving to a new server or changing Windows regional settings.

## Why it happens
`CDate("03/04/2026")` is read using the session/server locale (`Session.LCID`): US English reads it as 4 March, Indian or UK English as 3 April. When the date is pasted into SQL as text, SQL Server applies its *own* language/date-format setting, which may differ again. Printing dates with plain `Response.Write` also uses the server's regional format.

## How to spot it
- `CDate(Request(...))`, `IsDate` on free-text input, or `"... '" & someDate & "'"` in SQL strings.
- `Session.LCID` set in some pages but not others.
- Errors only for days 13 and above, or only after a server move.

## The fix
Accept dates in one unambiguous format: `` always submits `yyyy-mm-dd`. Parse it yourself with `Split` + `DateSerial` and reject impossible dates (DateSerial turns 31 Feb into 3 March, so compare the day back). Send it to SQL Server as a typed date parameter (`adDBTimeStamp` 135). For display, format explicitly with fixed month names instead of relying on the locale.

## Broken (excerpt)
```vb
fromText = Request.QueryString("from")  ' "03/04/2026"
d = CDate(fromText)        ' 4 March or 3 April? Depends on LCID

Set rs = conn.Execute("SELECT OrderID, OrderDate, Total " & _
    "FROM Orders WHERE OrderDate >= '" & fromText & "'")

Response.Write "Showing orders since " & d   ' server format
```

## Fixed (excerpt)
```vb
fromText = Trim(Request.QueryString("from") & "")  ' "2026-04-03"
fromDate = ParseIsoDate(fromText)   ' regex + DateSerial + day check
If IsNull(fromDate) Then errText = "Please pick a valid date."

cmd.CommandText = "SELECT OrderID, OrderDate, Total " & _
                  "FROM Orders WHERE OrderDate >= ?"
cmd.Parameters.Append cmd.CreateParameter( _
    "from", adDBTimeStamp, adParamInput, , fromDate)   ' 135

Response.Write "Showing orders since " & NiceDate(fromDate)
' -> "3 Apr 2026" on every server
```

## Checklist
- [ ] Date inputs use `type="date"` or a clearly labelled yyyy-mm-dd format.
- [ ] No `CDate` on user text; parsing uses `DateSerial` with validation.
- [ ] Dates go to SQL as parameters (or, if you must use text, as `'yyyymmdd'`).
- [ ] Date ranges use `>= start AND < next day` so times are included.
- [ ] Display uses an explicit format, not the server's regional settings.

## Running it
1. Enable Classic ASP in IIS and point a site or virtual directory at this repository.
2. Create the demo database with [`../assets/demo-db.sql`](../assets/demo-db.sql) (SQL Server / Express). Pages use the
   *Microsoft OLE DB Driver for SQL Server* (`Provider=MSOLEDBSQL`) and Windows authentication; give the app pool identity
   access to the `RescueDemo` database, or edit `CONN_STR`.
3. Open `broken.asp` and `fixed.asp` side by side.

Teaching example with made-up demo data. Not taken from any client system.
