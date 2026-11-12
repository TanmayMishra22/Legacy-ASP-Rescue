# Classic ASP fix 06: Cross-site scripting (XSS) from unencoded output

Visitor text written straight into HTML can run scripts in other users' browsers. Fix: encode every output with Server.HTMLEncode.

**Explainer page (live):** https://tanmaymishra22.github.io/Legacy-ASP-Rescue/examples/06-xss-output-encoding/

| File | What it is |
|---|---|
| `broken.asp` | The typical broken version, runnable under IIS with Classic ASP enabled |
| `fixed.asp` | The corrected version |
| `index.html` | Plain-language explainer with side-by-side broken vs fixed comparison |

## Symptom
A search page or comment shows a pop-up, a strange login box or a redirect. Security scanners or a client's audit report flag "Reflected/Stored XSS". Sometimes it just shows up as broken layout when someone types `<` or a quote.

## Why it happens
`` writes whatever the visitor sent as HTML. If that text contains `` or breaks out of an attribute with a quote, the browser runs it as part of your site, with access to the user's cookies and the page. Stored XSS (saved in the database and shown to others) is the most dangerous kind.

## How to spot it
- `` or `Response.Write` of database or form values without encoding.
- Values placed inside `value="..."`, `href` or inline `` blocks.
- Test with `x`: if the x is bold, output is not encoded.

## The fix
Encode at output time, for the place the value lands: `Server.HTMLEncode` for HTML text and double-quoted attributes, `Server.URLEncode` inside URLs, and a JavaScript-string escaper for values inside ``. Wrap them in tiny helpers (`H()`, `U()`, `JS()`) so it is easy to use them everywhere. `& ""` avoids the error `Server.HTMLEncode(Null)` raises for empty database fields.

## Broken (excerpt)
```vb
q = Request.QueryString("q")
%>
<input name="q" value="<%= q %>">
<p>You searched for: <%= q %></p>
<script>var lastSearch = '<%= q %>';</script>
```

## Fixed (excerpt)
```vb
Function H(v)
    H = Server.HTMLEncode(v & "")   ' Null-safe
End Function

q = Left(Request.QueryString("q") & "", 200)
%>
<input name="q" value="<%= H(q) %>">
<p>You searched for: <%= H(q) %></p>
<a href="fixed.asp?q=<%= U(q) %>">Link</a>
<script>var lastSearch = '<%= JS(q) %>';</script>
```

## Checklist
- [ ] Every dynamic value in HTML goes through `H()` / `Server.HTMLEncode`.
- [ ] Attributes always use double quotes (HTMLEncode does not encode single quotes).
- [ ] URL parts use `Server.URLEncode`; script values use a JS escaper.
- [ ] Database values are encoded too (stored XSS), not only Request values.
- [ ] Cookies that hold logins are set with Secure on HTTPS sites.

## Running it
1. Enable Classic ASP in IIS and point a site or virtual directory at this repository.
2. Create the demo database with [`../assets/demo-db.sql`](../assets/demo-db.sql) (SQL Server / Express). Pages use the
   *Microsoft OLE DB Driver for SQL Server* (`Provider=MSOLEDBSQL`) and Windows authentication; give the app pool identity
   access to the `RescueDemo` database, or edit `CONN_STR`.
3. Open `broken.asp` and `fixed.asp` side by side.

Teaching example with made-up demo data. Not taken from any client system.
