# Classic ASP fix 01: SQL injection in string-built queries

Visitor text glued into SQL lets anyone read or wipe your data. Fix: ADODB.Command parameters.

**Explainer page (live):** https://tanmaymishra22.github.io/Legacy-ASP-Rescue/examples/01-sql-injection/

| File | What it is |
|---|---|
| `broken.asp` | The typical broken version, runnable under IIS with Classic ASP enabled |
| `fixed.asp` | The corrected version |
| `index.html` | Plain-language explainer with side-by-side broken vs fixed comparison |

## Symptom
Search or login pages behave strangely when someone types a quote (`'`). One visitor sees every customer's record, a login works with no real password, or the page crashes with *Unclosed quotation mark after the character string*. In the worst case, data is changed or deleted.

## Why it happens
The page builds the SQL command by pasting the visitor's text into it: `"... WHERE City = '" & city & "'"`. SQL Server cannot tell which part was written by the developer and which part was typed by the visitor, so a typed quote ends the string early and the rest of the text runs as SQL.

## How to spot it
- Search the code for `Request(`, `Request.Form(` or `Request.QueryString(` values joined into a string that is later passed to `conn.Execute` or `rs.Open`.
- Look for home-made "protection" such as `Replace(x, "'", "''")`. It helps for text but not for numbers, and it is easy to forget once.
- Type a single quote `'` into a field on a test copy. A database error means the input reaches SQL as code.

## The fix
Use an `ADODB.Command` with `?` placeholders and add each value with `CreateParameter`. The value travels to SQL Server separately, as data, so quotes and SQL keywords are just ordinary characters. Also limit length to the column size and use a database login that has only the rights the site needs.

## Broken (excerpt)
```vb
city = Request.QueryString("city")

' Visitor text becomes part of the SQL command
sql = "SELECT CustomerID, Name, City FROM Customers " & _
      "WHERE City = '" & city & "'"

Set rs = conn.Execute(sql)
```

## Fixed (excerpt)
```vb
city = Trim(Request.QueryString("city") & "")
If Len(city) > 100 Then city = Left(city, 100)

Set cmd = Server.CreateObject("ADODB.Command")
Set cmd.ActiveConnection = conn
cmd.CommandType = adCmdText            ' 1
cmd.CommandText = "SELECT CustomerID, Name, City " & _
                  "FROM Customers WHERE City = ?"
cmd.Parameters.Append cmd.CreateParameter( _
    "city", adVarWChar, adParamInput, 100, city)   ' 202, 1

Set rs = cmd.Execute
```

## Checklist
- [ ] Every query that uses outside input uses `ADODB.Command` parameters.
- [ ] Parameter type and size match the column (e.g. `adVarWChar` 202 for NVARCHAR, `adInteger` 3 for INT).
- [ ] Numbers are checked with `IsNumeric` and converted with `CLng` before use.
- [ ] The site's SQL login is not `sa` and is not a db_owner.
- [ ] Detailed database errors are not shown to visitors (see example 05).

## Running it
1. Enable Classic ASP in IIS and point a site or virtual directory at this repository.
2. Create the demo database with [`../assets/demo-db.sql`](../assets/demo-db.sql) (SQL Server / Express). Pages use the
   *Microsoft OLE DB Driver for SQL Server* (`Provider=MSOLEDBSQL`) and Windows authentication; give the app pool identity
   access to the `RescueDemo` database, or edit `CONN_STR`.
3. Open `broken.asp` and `fixed.asp` side by side.

Teaching example with made-up demo data. Not taken from any client system.
