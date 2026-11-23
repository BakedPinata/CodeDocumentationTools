# UserForms

Two UserForms made in Excel for step S4 of the generalisation plan, with made-up code only, because this repository is
public, and exported from the VBA editor (File > Export File). Each is a `.frm` (the designer's `Begin` block, then
the code) and a `.frx` (a 24-byte header, then a compound file holding the form and its controls).

| Form | What's on it |
|---|---|
| `frmDemoForm1` | A label, a text box, a list box and a button. Its `UserForm_Initialize` fills the list box with three rows of two columns |
| `frmDemoForm2` | A MultiPage, with a frame and an image on its first page. It has no code. The frame and the image each have a picture, which makes the `.frx` a tree of storages rather than three streams |

**The pictures were replaced.** The sample's frame and image held a collage of product photographs that weren't ours
to publish. Each was replaced with a plain picture of the same size (860 × 645) saying "vbadoc fixture image", padded
with JPEG comment segments to exactly the same number of bytes, so every stream, length and offset in the `.frx` is
as the VBA editor wrote it. This was done with the S4 spike's `replace-image` command, in the repository's history
(`spikes/s4-userforms`). Nothing else was changed, and none of the four files holds a user name or a path.

SHA-256 of the files as committed:

```
f7180ce8e0e6c1c07f5685ac7f454d748c3b064d0535baa1814fc31c15bf404a  frmDemoForm1.frm
40494a300035cce7cde45bd7280c81b02153606b68801feadbe2418e03215d21  frmDemoForm1.frx
72415a3d388447fb1065c0190f2a2b2f5e8fafb1b5c6135fcbad010b5b4a5390  frmDemoForm2.frm
ba70af411b46fa7ee4d86434ca2389bb3c3fa741392f72c6eb8c662daa946571  frmDemoForm2.frx
```
