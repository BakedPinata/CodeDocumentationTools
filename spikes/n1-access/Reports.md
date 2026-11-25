# Reports

## Work PC

### Northwind Database

Results of first test on Northwind Traders Developer Edition 2.5 on the work PC.
No MS Access window opened for this run.

`dotnet run --project spikes/n1-access -- try "<redacted>\TestDatabase.accdb"`

```
accessspike, step N1. Microsoft Windows 10.0.26100, X64 process.
Database: (name withheld), 12,196 KB

Choice A, Access writes text from a copy: a new, empty database imports each object from the copy (the database itself is never opened)
  Got: 40 forms, 15 reports, 18 modules, 1 macros, 65 queries, 29 local and 0 linked tables; 14,280 KB of text; 66 modules' code (form and report modules included)
  Time: starting Access 2.3 s, listing 0.4 s, writing text 16.5 s, closing 3.0 s
  Start-up: 0 form(s) open, 0 TempVars set
  Access processes left running by this: 0

Choice A, Access writes text from a copy: the copy is opened with all code disabled (AutomationSecurity)
  Got: 40 forms, 15 reports, 18 modules, 1 macros, 65 queries, 29 local and 0 linked tables; 13,880 KB of text; 66 modules' code (form and report modules included)
  Time: starting Access 2.8 s, listing 0.3 s, writing text 1.6 s, closing 3.2 s
  Start-up: 1 form(s) open, 0 TempVars set
  Access processes left running by this: 0

Choice C, the file read directly, with no Access:
  Got: 0 modules' code (form and report modules included), in 0.2 s
  Against choice A's code: 0 the same, 0 different or not in A, 66 in A and not found directly
  The file's own store (MSysAccessStorage): 333 entries, 187 with data, 2,733 KB
    MSysAccessStorage_ROOT: 1 entries, 0 KB
    MSysAccessStorage_ROOT/Cmdbars: 1 entries, 0 KB
    MSysAccessStorage_ROOT/CustomGroups: 1 entries, 0 KB
    MSysAccessStorage_ROOT/DataAccessPages: 1 entries, 0 KB
    MSysAccessStorage_ROOT/Databases: 4 entries, 0 KB
    MSysAccessStorage_ROOT/Forms: 201 entries, 1,147 KB
    MSysAccessStorage_ROOT/ImExSpecs: 1 entries, 0 KB
    MSysAccessStorage_ROOT/Modules: 40 entries, 0 KB
    MSysAccessStorage_ROOT/PropData: 1 entries, 0 KB
    MSysAccessStorage_ROOT/Reports: 73 entries, 1,583 KB
    MSysAccessStorage_ROOT/Scripts: 5 entries, 1 KB
    MSysAccessStorage_ROOT/VBA: 3 entries, 0 KB
    MSysAccessStorage_SCRATCH: 1 entries, 0 KB
```

### Northwind Database `--as-a-person-would`

Same as above.

`dotnet run --project spikes/n1-access -- try "<redacted>\TestDatabase.accdb" --as-a-person-would`

```
accessspike, step N1. Microsoft Windows 10.0.26100, X64 process.
Database: (name withheld), 12,504 KB

Choice A, Access writes text from a copy: a new, empty database imports each object from the copy (the database itself is never opened)
  Got: 40 forms, 15 reports, 18 modules, 1 macros, 65 queries, 29 local and 0 linked tables; 14,280 KB of text; 66 modules' code (form and report modules included)
  Time: starting Access 1.9 s, listing 0.4 s, writing text 13.2 s, closing 2.9 s
  Start-up: 0 form(s) open, 0 TempVars set
  Access processes left running by this: 0

Choice A, Access writes text from a copy: the copy is opened with all code disabled (AutomationSecurity)
  Got: 40 forms, 15 reports, 18 modules, 1 macros, 65 queries, 29 local and 0 linked tables; 13,880 KB of text; 66 modules' code (form and report modules included)
  Time: starting Access 1.9 s, listing 0.3 s, writing text 1.1 s, closing 2.9 s
  Start-up: 1 form(s) open, 0 TempVars set
  Access processes left running by this: 0

Choice A, Access writes text from a copy: the copy is opened as a person would open it, so its start-up runs (for comparison)
  Got: 40 forms, 15 reports, 18 modules, 1 macros, 65 queries, 29 local and 0 linked tables; 13,880 KB of text; 66 modules' code (form and report modules included)
  Time: starting Access 1.9 s, listing 0.3 s, writing text 1.1 s, closing 2.9 s
  Start-up: 0 form(s) open, 0 TempVars set
  Access processes left running by this: 0

Choice C, the file read directly, with no Access:
  Got: 0 modules' code (form and report modules included), in 0.2 s
  Against choice A's code: 0 the same, 0 different or not in A, 66 in A and not found directly
  The file's own store (MSysAccessStorage): 333 entries, 187 with data, 2,733 KB
    MSysAccessStorage_ROOT: 1 entries, 0 KB
    MSysAccessStorage_ROOT/Cmdbars: 1 entries, 0 KB
    MSysAccessStorage_ROOT/CustomGroups: 1 entries, 0 KB
    MSysAccessStorage_ROOT/DataAccessPages: 1 entries, 0 KB
    MSysAccessStorage_ROOT/Databases: 4 entries, 0 KB
    MSysAccessStorage_ROOT/Forms: 201 entries, 1,147 KB
    MSysAccessStorage_ROOT/ImExSpecs: 1 entries, 0 KB
    MSysAccessStorage_ROOT/Modules: 40 entries, 0 KB
    MSysAccessStorage_ROOT/PropData: 1 entries, 0 KB
    MSysAccessStorage_ROOT/Reports: 73 entries, 1,583 KB
    MSysAccessStorage_ROOT/Scripts: 5 entries, 1 KB
    MSysAccessStorage_ROOT/VBA: 3 entries, 0 KB
    MSysAccessStorage_SCRATCH: 1 entries, 0 KB
```

### Production Database

Run at work with a production database from the published spike.
Running this did seem to open an access window when the copy was opened with all code disabled.

`<path_redacted>\accessspike.exe try "<path_redacted>\<name_redacted>.accdb"`

```
accessspike, step N1. Microsoft Windows 10.0.26100, X64 process.
Database: (name withheld), 1,078,992 KB

Choice A, Access writes text from a copy: a new, empty database imports each object from the copy (the database itself is never opened)
  Got: 118 forms, 96 reports, 3 modules, 0 macros, 25 queries, 10 local and 294 linked tables; 92,215 KB of text; 196 modules' code (form and report modules included)
  Time: starting Access 2.1 s, listing 2.4 s, writing text 111.0 s, closing 4.1 s
  Start-up: 0 form(s) open, 0 TempVars set
  Access processes left running by this: 0

Choice A, Access writes text from a copy: the copy is opened with all code disabled (AutomationSecurity)
  Got: 118 forms, 96 reports, 3 modules, 0 macros, 25 queries, 10 local and 294 linked tables; 89,084 KB of text; 196 modules' code (form and report modules included)
  Time: starting Access 3.6 s, listing 2.6 s, writing text 5.0 s, closing 3.0 s
  Start-up: 1 form(s) open, 0 TempVars set
  Access processes left running by this: 0

Choice C, the file read directly, with no Access:
  Got: 0 modules' code (form and report modules included), in 1.2 s
  Against choice A's code: 0 the same, 0 different or not in A, 196 in A and not found directly
  The file's own store (MSysAccessStorage): 1534 entries, 1080 with data, 32,675 KB
    MSysAccessStorage_ROOT: 1 entries, 0 KB
    MSysAccessStorage_ROOT/Cmdbars: 1 entries, 0 KB
    MSysAccessStorage_ROOT/CustomGroups: 1 entries, 0 KB
    MSysAccessStorage_ROOT/DataAccessPages: 1 entries, 0 KB
    MSysAccessStorage_ROOT/Databases: 4 entries, 0 KB
    MSysAccessStorage_ROOT/Forms: 595 entries, 14,492 KB
    MSysAccessStorage_ROOT/ImExSpecs: 4 entries, 0 KB
    MSysAccessStorage_ROOT/Modules: 11 entries, 0 KB
    MSysAccessStorage_ROOT/PropData: 1 entries, 0 KB
    MSysAccessStorage_ROOT/Reports: 495 entries, 3,415 KB
    MSysAccessStorage_ROOT/Scripts: 5 entries, 0 KB
    MSysAccessStorage_ROOT/VBA: 414 entries, 14,764 KB
    MSysAccessStorage_SCRATCH: 1 entries, 0 KB
```