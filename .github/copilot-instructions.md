# MFCMAPI coding conventions

## Blank lines after closing braces

- Put a blank line after a closing `}` that ends a control-flow block before
  the next statement or declaration. This is the "linefeeds after closing }"
  convention; CRLF line endings alone do not satisfy it.
- Keep associated `else`, `catch`, and do/while clauses together, and do not
  insert blank lines between consecutive closing braces.
- Check this spacing explicitly in changed code. clang-format does not insert
  every required blank line, so a successful format-check is not sufficient.

## Line endings

- Preserve CRLF line endings in the Windows working tree. Do not introduce LF-only
  lines, mixed line endings, or whole-file line-ending churn when editing.
- Respect [.gitattributes](../.gitattributes) and existing file encodings. Git may
  store LF in the index while the working-tree file uses CRLF; that is normal.
- Before finishing, check edited files with `git ls-files --eol -- <paths>` and
  `git diff --check`. Check new files' line endings too.

## Error checking

- Use the existing EC/WC wrappers in [error.h](../core/utility/error.h) for
  fallible Win32, GDI, COM, and MAPI calls. Do not add unchecked calls or discard
  failures with `static_cast<void>`.
- Follow surrounding code when choosing EC (log and display an error dialog)
  versus WC (log without displaying a dialog). Painting paths should not start
  displaying error dialogs.
- Match the wrapper to the API's documented return contract: `_H` for HRESULT,
  `_MAPI` for MAPI HRESULT, `_W32` for Win32 error codes, `_B` for BOOL success,
  and `_D(TYPE, ...)` when the native result is needed and zero means failure.
- `EC_B` and `WC_B` return **HRESULT**, not BOOL. Check them with
  `FAILED`/`SUCCEEDED`; never Boolean truth tests. `EC_D`/`WC_D` preserve the
  native result. Use the `_S` variants when only logging is needed.
- Do not blindly wrap APIs for which zero, false, or null is a valid result.
  Preserve their documented semantics and use the repository's error-reporting
  helpers for actual failures.

## Formatting and validation

- Follow [.editorconfig](../.editorconfig) and [.clang-format](../.clang-format).
  Use the repository formatter rather than approximating its output manually.
- After the final C/C++ edit, format only the changed files:
  `& .\scripts\clang.ps1 -Files @('UI\UIFunctions.cpp')`
- Then check the same files:
  `& .\scripts\clang.ps1 -Check -Files @('UI\UIFunctions.cpp')`
  Replace the example path with the actual changed C/C++ paths.
- Before opening or updating a PR with C/C++ changes, clang checks must pass for
  every C/C++ file changed by the complete PR, not just the latest edits. Fix
  formatting failures before proceeding.
- Code changes must also pass one representative Release build, normally
  `Release_Unicode|x64`, including affected application and unit-test targets.
  A Debug build alone is insufficient: Release optimization can expose additional
  warnings and link failures. One local Release configuration is enough; leave
  the remaining build matrix to CI.
- Run the smallest relevant existing build/tests for code changes. Formatting,
  line-ending checks, and EC/WC usage are separate requirements; a successful
  build does not replace them.
