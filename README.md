# Legacy-ASP-Rescue

Troubleshooting notes and proof-of-work for Classic ASP, IIS, SQL Server, APIs, and legacy web integrations.

**Live site:** https://tanmaymishra22.github.io/Legacy-ASP-Rescue/

**Classic ASP fixes (broken vs fixed):** https://tanmaymishra22.github.io/Legacy-ASP-Rescue/examples/

**Live local (IIS):** `http://localhost/Legacy-ASP-Rescue/`

## What's inside

| Path | What it is |
|------|------------|
| `index.html` | One-page portfolio: what Tanmay investigates and how he works |
| `examples/classic-asp-api-diagnostics/` | Classic ASP sample that calls an HTTP API, captures status/body, and prints a diagnosis |
| `examples/index.html` | Index of the 10 broken-vs-fixed examples |
| `examples/01-sql-injection/` | Fix 01: SQL injection in string-built queries |
| `examples/02-closed-recordset-leaks/` | Fix 02: "Operation is not allowed when the object is closed" and connection leaks |
| `examples/03-unicode-utf8/` | Fix 03: Garbled text: JosÃ©, â€™ and ???? instead of real characters |
| `examples/04-session-loss/` | Fix 04: Users logged out and carts emptied at random (Session loss) |
| `examples/05-hidden-500-errors/` | Fix 05: "500 - Internal server error" that hides the real problem |
| `examples/06-xss-output-encoding/` | Fix 06: Cross-site scripting (XSS) from unencoded output |
| `examples/07-n-plus-1-slow-pages/` | Fix 07: Slow pages from N+1 queries and row-by-row loops |
| `examples/08-rest-api-tls12/` | Fix 08: Calling modern REST/JSON APIs: TLS 1.2 errors and fragile parsing |
| `examples/09-dates-locale/` | Fix 09: Date bugs: 03/04 read as 4 March, and "out-of-range" datetime errors |
| `examples/10-file-upload/` | Fix 10: File uploads that fail, corrupt files or open a security hole |
| `examples/assets/` | Shared stylesheet and `demo-db.sql` (made-up demo data) |

## Classic ASP fixes: broken vs fixed

Each folder has a plain-language `index.html` explainer with broken and fixed code side by side, a `README.md`, and runnable `broken.asp` / `fixed.asp` pages. Create the demo database with [`examples/assets/demo-db.sql`](./examples/assets/demo-db.sql).

1. [SQL injection in string-built queries](./examples/01-sql-injection/) - [live page](https://tanmaymishra22.github.io/Legacy-ASP-Rescue/examples/01-sql-injection/)
2. ["Operation is not allowed when the object is closed" and connection leaks](./examples/02-closed-recordset-leaks/) - [live page](https://tanmaymishra22.github.io/Legacy-ASP-Rescue/examples/02-closed-recordset-leaks/)
3. [Garbled text: JosÃ©, â€™ and ???? instead of real characters](./examples/03-unicode-utf8/) - [live page](https://tanmaymishra22.github.io/Legacy-ASP-Rescue/examples/03-unicode-utf8/)
4. [Users logged out and carts emptied at random (Session loss)](./examples/04-session-loss/) - [live page](https://tanmaymishra22.github.io/Legacy-ASP-Rescue/examples/04-session-loss/)
5. ["500 - Internal server error" that hides the real problem](./examples/05-hidden-500-errors/) - [live page](https://tanmaymishra22.github.io/Legacy-ASP-Rescue/examples/05-hidden-500-errors/)
6. [Cross-site scripting (XSS) from unencoded output](./examples/06-xss-output-encoding/) - [live page](https://tanmaymishra22.github.io/Legacy-ASP-Rescue/examples/06-xss-output-encoding/)
7. [Slow pages from N+1 queries and row-by-row loops](./examples/07-n-plus-1-slow-pages/) - [live page](https://tanmaymishra22.github.io/Legacy-ASP-Rescue/examples/07-n-plus-1-slow-pages/)
8. [Calling modern REST/JSON APIs: TLS 1.2 errors and fragile parsing](./examples/08-rest-api-tls12/) - [live page](https://tanmaymishra22.github.io/Legacy-ASP-Rescue/examples/08-rest-api-tls12/)
9. [Date bugs: 03/04 read as 4 March, and "out-of-range" datetime errors](./examples/09-dates-locale/) - [live page](https://tanmaymishra22.github.io/Legacy-ASP-Rescue/examples/09-dates-locale/)
10. [File uploads that fail, corrupt files or open a security hole](./examples/10-file-upload/) - [live page](https://tanmaymishra22.github.io/Legacy-ASP-Rescue/examples/10-file-upload/)

## Example: Classic ASP API diagnostics

See [`examples/classic-asp-api-diagnostics/`](./examples/classic-asp-api-diagnostics/). Run the `.asp` under IIS (Classic ASP enabled). Uses `MSXML2.ServerXMLHTTP.6.0`. No real credentials are included.

## Contact

- GitHub: https://github.com/TanmayMishra22
- Email: devmishratanmaydev@gmail.com
