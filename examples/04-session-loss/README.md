# Classic ASP fix 04: Users logged out and carts emptied at random (Session loss)

Session() data vanishes on app pool recycles, web gardens and load balancers. Fix: keep important state in the database.

**Explainer page (live):** https://tanmaymishra22.github.io/Legacy-ASP-Rescue/examples/04-session-loss/

| File | What it is |
|---|---|
| `broken.asp` | The typical broken version, runnable under IIS with Classic ASP enabled |
| `fixed.asp` | The corrected version |
| `index.html` | Plain-language explainer with side-by-side broken vs fixed comparison |

## Symptom
Users are suddenly logged out, a shopping cart or multi-step form empties itself, or it happens only on the live server and never on the developer's PC. Complaints often cluster around a certain time of day or after quiet periods.

## Why it happens
Classic ASP keeps `Session()` values in the memory of one IIS worker process. That memory is thrown away when the app pool recycles (by default every 1,740 minutes, on config changes, or on memory limits), when it shuts down after being idle (default 20 minutes), and it is not shared between processes in a web garden or between servers behind a load balancer. A change of host name (www vs no-www) also gives the visitor a new session cookie.

## How to spot it
- Important data (login, cart, wizard steps) stored only in `Session("...")`.
- App pool *Maximum Worker Processes* greater than 1 (web garden), or more than one server without sticky sessions.
- Event Viewer (System log, source WAS) shows recycles at the time users complain.
- `Session.Timeout`, the ASP session timeout and the app pool idle timeout do not match.

## The fix
Keep only small, rebuildable things in Session. Store anything users would be upset to lose in SQL Server, keyed by a random token kept in a cookie (here created with SQL Server's `NEWID()`). Validate the token's format before using it, and always pass it as a parameter. As short-term relief, align the timeouts, turn off the web garden and schedule recycles for quiet hours.

## Broken (excerpt)
```vb
pid = Request.QueryString("add")
If Len(pid) > 0 And IsNumeric(pid) Then
    ' Lives only in this worker process's memory
    Session("Cart") = Session("Cart") & CLng(pid) & ","
End If

cart = Session("Cart") & ""
If Len(cart) = 0 Then
    Response.Write "Your cart is empty."
End If
```

## Fixed (excerpt)
```vb
token = Request.Cookies("CartToken")
If Not IsValidToken(token) Then           ' GUID regex check
    token = NewToken(conn)                ' SELECT NEWID()
    Response.Cookies("CartToken") = token
    Response.Cookies("CartToken").Expires = DateAdd("d", 30, Now())
End If

Set cmd = MakeCmd(conn, _
  "UPDATE CartItems SET Qty = Qty + 1 " & _
  "WHERE CartToken = ? AND ProductID = ?; " & _
  "IF @@ROWCOUNT = 0 INSERT INTO CartItems " & _
  "(CartToken, ProductID, Qty) VALUES (?, ?, 1);")
' ...append the 4 parameters, then cmd.Execute
```

## Checklist
- [ ] Carts, logins and form progress are stored in the database, not only in Session.
- [ ] The cookie holds a random token only, never user IDs or prices.
- [ ] Token format is validated (regex) and used only in parameterised queries.
- [ ] App pool idle timeout, recycling schedule and web garden settings are reviewed.
- [ ] Old cart rows are cleaned up by a scheduled SQL job.

## Running it
1. Enable Classic ASP in IIS and point a site or virtual directory at this repository.
2. Create the demo database with [`../assets/demo-db.sql`](../assets/demo-db.sql) (SQL Server / Express). Pages use the
   *Microsoft OLE DB Driver for SQL Server* (`Provider=MSOLEDBSQL`) and Windows authentication; give the app pool identity
   access to the `RescueDemo` database, or edit `CONN_STR`.
3. Open `broken.asp` and `fixed.asp` side by side.

Teaching example with made-up demo data. Not taken from any client system.
