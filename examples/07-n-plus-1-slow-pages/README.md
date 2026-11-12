# Classic ASP fix 07: Slow pages from N+1 queries and row-by-row loops

One query per row turns a report into thousands of round trips. Fix: one JOIN/GROUP BY query plus GetRows.

**Explainer page (live):** https://tanmaymishra22.github.io/Legacy-ASP-Rescue/examples/07-n-plus-1-slow-pages/

| File | What it is |
|---|---|
| `broken.asp` | The typical broken version, runnable under IIS with Classic ASP enabled |
| `fixed.asp` | The corrected version |
| `index.html` | Plain-language explainer with side-by-side broken vs fixed comparison |

## Symptom
A report or list page that used to be quick now takes many seconds or times out (*Script timed out* / *The maximum amount of time for a script to execute was exceeded*). It is fine on the test database and slow on live data, and it gets worse every month as data grows.

## Why it happens
The page runs one query to get a list, then another query inside the loop for every row ("N+1 queries"). Each round trip has a fixed cost, so 2,000 customers means 2,001 trips. Reading `rs("Field")` many times in a long loop and building huge strings with `&` add more cost.

## How to spot it
- `conn.Execute` or `rs.Open` inside a `Do While Not rs.EOF` loop.
- SQL Server Profiler / Extended Events shows the same query repeated with different IDs.
- Add a counter and `Timer()` to the page (as in broken.asp) to see the number of queries and seconds.

## The fix
Let SQL Server do the work in one query: `LEFT JOIN` the child table and use `GROUP BY` with `COUNT`/`SUM`. Copy the result into an array with `GetRows`, close the connection, then build the HTML from the array. Check that the join columns are indexed (here `IX_Orders_CustomerID`).

## Broken (excerpt)
```vb
Set rsC = conn.Execute("SELECT CustomerID, Name FROM Customers")
Do While Not rsC.EOF
    ' one extra round trip per customer
    Set rsO = conn.Execute( _
      "SELECT COUNT(*) AS N, ISNULL(SUM(Total),0) AS T " & _
      "FROM Orders WHERE CustomerID = " & CLng(rsC("CustomerID")))
    Response.Write "<tr><td>" & rsC("Name") & "</td><td>" & rsO("N") ...
    rsO.Close
    rsC.MoveNext
Loop
```

## Fixed (excerpt)
```vb
sql = "SELECT c.Name, COUNT(o.OrderID) AS OrderCount, " & _
      "ISNULL(SUM(o.Total), 0) AS OrderTotal " & _
      "FROM Customers c " & _
      "LEFT JOIN Orders o ON o.CustomerID = c.CustomerID " & _
      "GROUP BY c.CustomerID, c.Name ORDER BY c.Name"
Set rs = conn.Execute(sql)
If Not rs.EOF Then data = rs.GetRows()
rs.Close : conn.Close        ' done with the database

For i = 0 To UBound(data, 2)
    Response.Write "<tr><td>" & Server.HTMLEncode(data(0, i) & "") ...
Next
```

## Checklist
- [ ] No queries inside loops over another recordset.
- [ ] Totals and counts are calculated in SQL (JOIN + GROUP BY), not in VBScript.
- [ ] Results are read with `GetRows`; the connection is closed before HTML output.
- [ ] Join and filter columns have indexes.
- [ ] Large lists are paged (`OFFSET ... FETCH` or TOP) instead of showing everything.

## Running it
1. Enable Classic ASP in IIS and point a site or virtual directory at this repository.
2. Create the demo database with [`../assets/demo-db.sql`](../assets/demo-db.sql) (SQL Server / Express). Pages use the
   *Microsoft OLE DB Driver for SQL Server* (`Provider=MSOLEDBSQL`) and Windows authentication; give the app pool identity
   access to the `RescueDemo` database, or edit `CONN_STR`.
3. Open `broken.asp` and `fixed.asp` side by side.

Teaching example with made-up demo data. Not taken from any client system.
