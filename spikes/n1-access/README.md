# N1 spike: getting forms and VBA out of an Access database

Throwaway code for step N1 of the [generalisation plan](../../GENERALISATION-PLAN.md), kept out of the tools and
their solutions. It tries the plan's two ways of reading a database (5.5) and reports what each got and how long it
took:

- **Choice A:** Access writes each form, report, module and macro as text (its own `SaveAsText`) from a copy of the
  database, hidden, and is closed afterwards. Tried two ways: a new, empty database **imports** each object from the
  copy, so the database itself is never opened; and the copy is **opened with code disabled**.
- **Choice C:** the database file is read directly, with no Access and no database engine.

It never changes the database it's given: it copies it to a temporary folder and works on the copy.

## Running it

From the repository root, on a machine with the .NET 10 SDK:

```bash
dotnet run --project spikes/n1-access -- try <database.accdb>
```

For a machine without the SDK, or where only an executable can be carried in, publish it as one file and copy
`accessspike.exe` across:

```bash
dotnet publish spikes/n1-access -c Release -r win-x64 -o <folder>
<folder>\accessspike.exe try <database.accdb>
```

| Command | What it does |
|---|---|
| `try <database>` | Tries both choices and prints the report. Add `--names` to list names in the report (leave it off for a real database: without it the report has only counts and times), and `--keep <folder>` to keep the text Access wrote, which is deleted otherwise |
| `try <sample> --as-a-person-would` | Also opens the copy the ordinary way, letting start-up run, to compare. Only for the sample |
| `build-sample <new.accdb>` | Builds the sample database through Access, from `sample/` and `SampleBuilder.cs` |
| `grow <database> <new.accdb> <n>` | A copy with its forms copied until there are `n`, to time a database the size of a real one |

## What to run, and what to send back

1. **At home,** on Northwind 2.0 made from Access's template into `accessdoc/private/`:
   `dotnet run --project spikes/n1-access -- try accessdoc/private/<Northwind file>.accdb --names`
2. **On a work machine,** with the published `accessspike.exe`, on the sample (`build-sample` there, or carry
   `Sample.accdb` across) and then on **one real database**, without `--names`.
3. Send back the report each printed. Without `--names` it holds no names, paths or code. What matters most:
   - does `accessspike.exe` start at all there (application control);
   - does "Choice A … imports" get everything, with `0 form(s) open, 0 TempVars set`;
   - does anything appear on screen while it runs (a security prompt, a log-in for a linked table, Access's window);
   - `Access processes left running by this` should be `0`;
   - how long "writing text" took for a real database;
   - what choice C says against choice A's code.

If a run is stuck for more than a few minutes, Access is probably showing a prompt it can't show while hidden. End
`MSACCESS.EXE` in Task Manager, and say where it stopped.

## What it found here (2026-11-15, Access 16.0 64-bit on Windows 11, Arm64)

**Choice A works, by importing.** A new, empty database imports each object from the copy
(`DoCmd.TransferDatabase`) and writes it out with `SaveAsText`. The database given is never the one open, so none of
its start-up can run: on the sample, whose start-up form and AutoExec macro each leave a mark, there were no forms
open, no TempVars set and no rows added to its log.

**Opening the copy with code disabled isn't enough.** With `AutomationSecurity` set to disable everything, the
sample's start-up form still opened and its AutoExec macro still ran its `SetTempVar` action: only VBA, and the
macro actions Access counts as unsafe, were stopped. A start-up form that opens also reads its record source, which
for a real database could be a linked table asking for a log-in. So this way is out, and importing is the one to use.

**Times.** Starting Access takes about 2 s and closing it about 3 s, whatever the database. Importing and writing
takes about 0.17 s an object: 1 s for the sample's 4 forms, **26 s for 150 forms**. Opening the copy writes 150
forms in under 1 s, so importing costs about 25 s on a database that size for running nothing of it.

**The text is what the plan expected:** forms and macros in UTF-16 with a byte order mark, modules in the system's
code page, a form's code after a `CodeBehindForm` line. A form of 13 controls is about 500 lines, with the things the
Version Control add-in strips (`Checksum`, `NameMap`, `GUID`, and so on) present. No printer settings were written
for forms made by script. Queries' SQL and tables (local or linked, with the connection) came from the database
engine without opening anything.

**Choice C gets the VBA, and not the forms.** A database keeps its VBA project in a system table,
`MSysAccessStorage`, laid out as a workbook's `vbaProject.bin` is, so about 300 lines reading the file's pages
(`JetFile.cs`) and vbadoc's own reading of the project gave every module's code, form modules included, in well
under a second, with no Access: 152 of 152 the same as Access's text on the 150-form database. The one difference
seen is the case of identifiers: the file keeps a name as it was typed (`quantity`), and Access writes it as the
project now spells it (`Quantity`). Each form's design is there too, as binary streams (`Blob`, `TypeInfo`,
`PropData`) in a format nobody has published, which is what the plan said: no layout this way.

**Access writes the Windows user's name into a new database** as its Author (and an organisation as its Company).
`build-sample` replaces it with "Fixture Author" as soon as the database is made; the built sample was then searched
for the name, in 8-bit and UTF-16, and has none.

**After building forms in design view, Access didn't quit when asked** and was ended by the spike (only in
`build-sample`; every `try` closed it cleanly). Reading never opens a form's design, so this doesn't touch accessdoc.

## Recommendation, before your runs

- **accessdoc gets a database's text by choice A with importing,** as above, and reads a folder of that text (its
  own, or the add-in's) with no Access.
- **Choice C is worth keeping for the VBA alone:** given a database on a machine where Access can't be started, the
  page of its VBA can still be made, and it gives the tests a way to read a database with nothing installed. It
  can't give forms, so it isn't a replacement.
- Both wait on what the work machine says: whether the executable may start Access there at all, and whether a real
  database imports without a prompt.
