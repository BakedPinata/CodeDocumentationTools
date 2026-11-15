# The Version Control add-in's test source

`Testing.accdb.src/` is the exported source of the test database of the
[Access Version Control add-in](https://github.com/joyfullservice/msaccess-vcs-addin) (msaccess-vcs-addin), copied
unchanged from its `Testing/Testing.accdb.src` folder at commit `55eb68290df3d5678711bf8c11d2adde25e8ba5c`
(2026-10-28). It's here as real `SaveAsText` text for accessdoc to read: forms (`.form` with their code in `.cls`),
reports, a macro, modules, queries, and local and linked table definitions.

- **Licence:** the add-in's, in [LICENSE.txt](LICENSE.txt) beside this file (BSD-style: use, copy and distribute
  with the copyright notice and permission notice kept). It's compatible with this repository's MIT licence.
- **Left out:** the folder's `AGENTS.md` and `vcs-agent-docs/`, which are the add-in's notes for its own
  development and not part of the database's source.
- **Never changed.** Fixtures are kept byte for byte (`.gitattributes`), so a form's text is as Access and the
  add-in wrote it. To take a newer copy, replace the folder whole and update the commit above.
- The database was made for testing by the add-in's authors and holds nothing of ours.
