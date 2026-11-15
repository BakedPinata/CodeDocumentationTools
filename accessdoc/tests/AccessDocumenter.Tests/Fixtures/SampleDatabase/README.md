# The sample database, as text

`text/` is the sample database as Access itself wrote it (`SaveAsText`), by step N1's spike on 2026-11-15, kept byte
for byte: forms and macros in UTF-16, modules in code page 1252, each query's SQL, and a list of its tables.

The database is made up for accessdoc's tests and built by script through Access (`spikes/n1-access`, command
`build-sample`, until the spike is deleted; it's in the history after that). It has:

- `frmStart`, the start-up form, whose Load event leaves a mark, and an `AutoExec` macro that leaves two, so a
  reader can prove it ran neither;
- `frmCustomers`, with a combo box that has a row source, an option group, and a tab control whose first page holds
  a subform, `frmOrders`, which holds a subform of its own, `frmOrderLines`, a continuous form;
- events run by event procedures, by a macro (`cmdGreet`, `mcrGreet`) and by an expression (`cmdHello`,
  `=SayHello()`);
- a standard module (`basUtility`) and a class module (`clsOrderTotals`);
- local tables (`Customers`, `Orders`, `OrderLines`, `StartupLog`), a linked one (`Regions`), and two queries.

**Nothing personal is in it.** Access writes the Windows user's name into a new database as its Author; the script
replaces that with "Fixture Author" as soon as the database is made. The linked table's back end is in the public
documents folder (`C:\Users\Public\Documents\AccessDocSample\`), so its path names nobody. This text was read
through, and searched for a user name and for printer settings, before it was committed.

The database file itself (`Sample.accdb`) isn't committed yet: see step N1 in the generalisation plan.
