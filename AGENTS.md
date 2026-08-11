# Repository Instructions

## VBA source files

- Save every `.bas`, `.cls`, and `.frm` file as UTF-8 with CRLF line endings.
- Use only characters representable in Windows code page 936 (CP936); prefer common Chinese, ASCII, and standard Chinese punctuation over emoji or uncommon Unicode symbols.
- After editing VBA source files, verify every changed file:
  1. decodes as strict UTF-8;
  2. contains no lone LF line endings;
  3. encodes to CP936 without replacement or data loss.

Keep repository source files in UTF-8. `sync-vba.ps1` performs the CP936 conversion only at the VBE import boundary.

## VBA synchronization

- Preserve `sync-vba.ps1` as UTF-8 with BOM so Windows PowerShell 5.1 reads its Chinese text correctly.
- Run VBA synchronization only when the script reports Windows ANSI code page CP936; do not bypass this guard.
- Treat a synchronization as successful only when import and VBE read-back validation both complete without errors.
- Use a temporary workbook copy when testing synchronization changes.

See `README.md` under **VBA 同步** for operator setup and usage.
