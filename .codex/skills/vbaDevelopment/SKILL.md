---
name: vbaDevelopment
description: Use when editing or generating VBA source files such as .bas, .cls, or .frm, especially when Chinese text, Excel VBA import/export, code synchronization, line-continuation syntax, or mojibake/encoding issues may occur.
---

# VBA Development

## Overview
Use this skill when working on VBA source files that will be imported back into Excel. The main failure mode is source integrity rather than logic: VBA is sensitive to encoding, continuation lines, and aggressive syntax shrinking. Treat importability and readable Chinese text as part of correctness.

## When to Use
- Editing `.bas`, `.cls`, or `.frm` files.
- Fixing Chinese garbling, mojibake, or sync issues between editor and Excel VBA.
- Rewriting exported VBA under `VBA_Export`.
- Refactoring `If ... Then`, continuation `_`, or compact block syntax that may break after formatting.

## Required Workflow
1. Read the whole VBA file before editing. Do not assume a local replacement is safe.
2. Rewrite the full source file when changing VBA code in exported modules.
3. Keep continuation lines contiguous. If a line ends with ` _`, the next physical line must be the continuation line. Do not insert blank lines there.
4. Preserve explicit block syntax. Do not over-compress `If ... End If`, `For ... Next`, `With ... End With`, or error-handling blocks.
5. After writing any `.bas`, `.cls`, or `.frm` under `VBA_Export`, run:

```powershell
python convert.py
```

6. Confirm the file was converted to GBK/ANSI so Chinese text displays correctly after importing into the VBA editor.

## Syntax Guardrails
- Never leave a trailing `_` followed by a blank line.
- Never convert a block `If` into malformed single-line syntax.
- Every `If` must have a matching `End If` unless it is a valid single-line `If`.
- Every `For` must have a matching `Next`.
- Prefer clear multi-line VBA over aggressive syntax shrinking.

## Encoding Guardrails
- `VBA_Export` source files must remain compatible with Excel VBA import on Windows Chinese environments.
- UTF-8 text that looks correct in the editor can still import as garbled text in VBA.
- `python convert.py` is mandatory after writing VBA source files because it normalizes CRLF and converts UTF-8 output to GBK.

## Verification
- Re-open the written module and inspect Chinese comments and strings.
- Check every line ending in ` _` to ensure the continuation is on the next line.
- Search for broken constructs such as separated `If ... Then` continuations or missing `End If`.
- If the user reports “乱码” or “导入后中文不对”, assume encoding is wrong until `python convert.py` has been run.

## Example
Bad:

```vb
If InStr(1, txt, kw & CStr(numArr(ii)), vbTextCompare) > 0 And _

   InStr(1, txt, kw & "#" & CStr(numArr(ii)), vbTextCompare) = 0 Then
    sa = CStr(k)
End If
```

Good:

```vb
If InStr(1, txt, kw & CStr(numArr(ii)), vbTextCompare) > 0 And _
   InStr(1, txt, kw & "#" & CStr(numArr(ii)), vbTextCompare) = 0 Then
    sa = CStr(k)
End If
```
