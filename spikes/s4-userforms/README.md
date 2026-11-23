# S4a spike: UserForms written into a workbook

Throwaway code for step S4a of the [generalisation plan](../../GENERALISATION-PLAN.md), kept out of the tools and
their solutions. It builds new workbooks that hold the sample's two UserForms, each made a different way, so Excel can
say which ways it accepts before `vbadoc import` takes forms (S4b).

Each workbook is a new workbook written by vbadoc (`TargetWorkbook.New`). The spike adds the forms to its project, with
a module `SpikeTest` that has two procedures:

- `Check`, a function that loads both forms without showing them and says what it found: the caption, the controls by
  name, and the rows `UserForm_Initialize` put in the list box. Run it from the Immediate window as `?Check`.
- `ShowForms`, which shows the first form and then the second.

## The workbooks

| Workbook | What it tries |
|---|---|
| `F1-files.xlsm` | From the VBA editor's `.frm` and `.frx`: the designer copied from the `.frx`, the `.frm`'s `Begin` block as the form's `VBFrame`, the code with a `VB_Base` of new GUIDs, and a reference to `FM20.DLL` |
| `F2-no-vb-base.xlsm` | As F1, without `VB_Base` |
| `F3-no-reference.xlsm` | As F1, with no reference to the MSForms library |
| `F4-designer-as-workbook.xlsm` | As F1, with the designer's class id and `CompObj` as a workbook keeps them (zeros, and no ProgID) |
| `C1-copied.xlsm` | From the sample workbook: the designer storage, `VBFrame`, the code (with its `VB_Base`) and the MSForms reference copied as they are |

## Building them

From the repository root, with the sample in `vbadoc/private/s4/` (`Userforms-Spike.xlsm`, and `frmDemoForm1` and
`frmDemoForm2` exported from it as `.frm` and `.frx`):

```bash
dotnet run --project spikes/s4-userforms -- build vbadoc/private/s4 vbadoc/private/s4/out
```

Add the start of a workbook's name to build only some (`… out F1 C1`). `inspect <workbook or .frx>` prints what's in a
project or a `.frx`, `compare <workbook> <form.frx>` compares a form's storage with its `.frx`, stream by stream, and
`dir <workbook>` prints the project's `dir` records.

The workbooks are in `vbadoc/private/`, which is never committed. `C1-copied.xlsm` has the sample's MSForms reference,
whose path is in your Temp folder, so it has your Windows user name in it, as the sample does.

## What to run, and what to send back

Carry the five workbooks in `vbadoc/private/s4/out/` to each machine: at home, and then a work machine. For each one:

1. Open it, and enable macros if asked.
2. Open the VBA editor (Alt+F11). Are both forms under Forms, and does each open in the designer (double-click it)
   with its controls: a label, a text box, a list box and a button on the first; a MultiPage, a frame and an image on
   the second?
3. In the Immediate window (Ctrl+G), type `?Check` and press Enter. Send back what it prints, or the error.
4. Run `ShowForms` (put the cursor in it and press F5). Does each form show as it did in the sample, with three rows
   in the list box? Close the first to see the second.
5. For F1 only: save it, close Excel, open it again and do step 3 once more.

F2 is expected to fail: say what the message is, and where it appears (on opening, in the VBA editor, or on `?Check`).

## What it found here (2026-11-23, Excel 16.0 64-bit on Windows 11, through a hidden Excel)

**A form written from its two files works** (F1): both forms load with every control, the first form's
`UserForm_Initialize` fills its list, and a copy saved by Excel opens and works again. `?Check` gave, for F1, F3, F4
and C1:

```text
frmDemoForm1: UserForm1, 4 controls (lblDemoLabel txbDemoTextBox lbxDemoListBox btnDemoButton), list of 3 rows
frmDemoForm2: UserForm2, 3 controls (MultiPage1 Frame1 Image1)
```

**A form needs its `VB_Base`** (F2): without it, Excel stopped on a dialog that a hidden Excel can't show, as a class
without its attributes did (see vbadoc's DECISIONS, on classes). Its value doesn't matter: Excel gave every form two
new GUIDs when it compiled the project, the copied ones included, so a writer can make up its own. The VBA editor does
the same on import.

**What else the project needs for a form**, found by reading the sample and confirmed by F1:

- A designer storage at the top of the project, named after the form, holding everything in the `.frx` after its
  24-byte header. That's a whole tree, not three streams: a MultiPage, a frame or an image is a storage of its own
  inside it.
- A `VBFrame` stream (`\x03VBFrame`) in that storage: the `.frm`'s `Begin … End` block without its `OleObjectBlob`
  line. The workbook's has `TypeInfoVer` in its place, which Excel leaves out when it saves a form written without it.
- A line `BaseClass=<form>` in `PROJECT`, and one `Package={AC9F2F90-E877-11CE-9F68-00AA00574A4F}` line for all forms.
- The form's `dir` records are a class's with `MODULEPRIVATE` (`0x0028`) before the end.
- The code has `VB_Base`, `VB_TemplateDerived` and `VB_Customizable` added, as for a class.

**How the `.frx` differs from the workbook's storage doesn't matter** (F4): the `.frx` gives its root the form's class
id and its `CompObj` the ProgID `Forms.Form.1`, where the workbook has zeros and no ProgID. Both work. Excel keeps what
it was given when it saves. The `f` streams also differ in a few padding bytes, which hold nothing.

**The MSForms reference:** a reference to `FM20.DLL` by its GUID works, and Excel keeps it as it is (F1). With none,
the forms still work (F3) and Excel doesn't add one, but a form's code that names an MSForms type, as an event like
`KeyPress` does (`MSForms.ReturnInteger`), needs it. The sample's own reference is a control reference to a cache of
MSForms made on each machine, with a path in the user's Temp folder, and copied as it is it works too (C1).

**The `.frx` header** is `LB`, `08 00`, the size of the rest, eight zeros, and two numbers that look like the form's
outer size: the `.frm`'s `ClientWidth` plus 192 and `ClientHeight` plus 512 in both samples. Only S4c, writing a
`.frx`, needs them.

## For S4b, before your runs

As the plan has it, with these:

- A form from a folder: the designer from the `.frx` as it is, `VBFrame` from the `Begin` block, new GUIDs for
  `VB_Base`, the records and lines above, and a reference to `FM20.DLL` when the project has no MSForms reference.
- A form from a workbook: its storage and `VBFrame` copied as they are, its code as stored, and its MSForms reference
  copied as any other reference is.
