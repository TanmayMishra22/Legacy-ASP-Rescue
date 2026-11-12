# Classic ASP fix 02: "Operation is not allowed when the object is closed" and connection leaks

Recordsets used after their connection is closed, and connections never released. Fix: GetRows, then close everything.

**Explainer page (live):** https://tanmaymishra22.github.io/Legacy-ASP-Rescue/examples/02-closed-recordset-leaks/

| File | What it is |
|---|---|
| `broken.asp` | The typical broken version, runnable under IIS with Classic ASP enabled |
| `fixed.asp` | The corrected version |
| `index.html` | Plain-language explainer with side-by-side broken vs fixed comparison |

## Symptom
A page shows *ADODB.Recordset error '800a0e78' Operation is not allowed when the object is closed*. Or the site works in the morning and gets slower through the day until pages time out, and an app pool recycle "fixes" it for a while.

## Why it happens
A recordset belongs to its connection. When a helper function closes the connection and returns the recordset, the recordset is closed too, so the first `rs.EOF` fails. The opposite mistake is never closing: each forgotten connection stays busy until the script object is destroyed, and under load the connection pool runs out. A third common cause of the same error is a stored procedure or batch without `SET NOCOUNT ON`: the "rows affected" message comes back as a first, closed recordset.

## How to spot it
- Functions that do `conn.Close` and then `Set Fn = rs`.
- `Server.CreateObject("ADODB.Connection")` without a matching `.Close` and `Set ... = Nothing` on every path (watch early `Response.End` and `Exit Function`).
- Stored procedures that INSERT/UPDATE before their SELECT but have no `SET NOCOUNT ON`.
- In SQL Server, `sp_who2` or `sys.dm_exec_sessions` showing many sleeping sessions from the web server.

## The fix
Open, read and close in the same function. Copy the rows into a plain array with `rs.GetRows()` (or copy single values into variables), then close the recordset and connection and set them to `Nothing`. Return the array, not the recordset. Add `SET NOCOUNT ON` at the top of stored procedures.

## Broken (excerpt)
```vb
Function GetCustomers()
    Dim conn, rs
    Set conn = Server.CreateObject("ADODB.Connection")
    conn.Open CONN_STR
    Set rs = conn.Execute("SELECT Name, City FROM Customers")
    conn.Close            ' also closes rs!
    Set GetCustomers = rs
End Function

Set rs = GetCustomers()
Do While Not rs.EOF      ' error '800a0e78' here
```

## Fixed (excerpt)
```vb
Function GetCustomers()
    Dim conn, rs
    GetCustomers = Empty
    Set conn = Server.CreateObject("ADODB.Connection")
    conn.Open CONN_STR
    Set rs = conn.Execute("SELECT Name, City FROM Customers")
    If Not rs.EOF Then GetCustomers = rs.GetRows()
    rs.Close   : Set rs = Nothing
    conn.Close : Set conn = Nothing
End Function

data = GetCustomers()
If IsArray(data) Then
    For i = 0 To UBound(data, 2)
        Response.Write Server.HTMLEncode(data(0, i) & "")
    Next
End If
```

## Checklist
- [ ] Every `.Open` has a matching `.Close` and `Set x = Nothing`.
- [ ] Functions return arrays or values, never live recordsets.
- [ ] Single values (`COUNT(*)`, IDs) are copied out before closing.
- [ ] Stored procedures start with `SET NOCOUNT ON`.
- [ ] One connection string is used everywhere so pooling works.

## Running it
1. Enable Classic ASP in IIS and point a site or virtual directory at this repository.
2. Create the demo database with [`../assets/demo-db.sql`](../assets/demo-db.sql) (SQL Server / Express). Pages use the
   *Microsoft OLE DB Driver for SQL Server* (`Provider=MSOLEDBSQL`) and Windows authentication; give the app pool identity
   access to the `RescueDemo` database, or edit `CONN_STR`.
3. Open `broken.asp` and `fixed.asp` side by side.

Teaching example with made-up demo data. Not taken from any client system.
