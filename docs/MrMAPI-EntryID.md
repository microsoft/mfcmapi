# Read one message property by EntryID

```
MrMAPI -EntryID <hex-entryid> -Profile "<profile-name>" -Store <index> <property>
```

For example, replace the placeholder with the full hexadecimal `PR_ENTRYID`
(Outlook's `MailItem.EntryID`) of a message:

```
MrMAPI -EntryID 00000000... -Profile "Test profile" -Store 0 PR_SUBJECT_W
MrMAPI -EntryID 00000000... -Profile "Test profile" -Store 0 0x0037001F
```

The ellipsis is illustrative, not valid input. EntryIDs must contain 1–8192 bytes,
encoded as contiguous pairs of hexadecimal digits, without a `0x` prefix,
whitespace, or separators. This accepts opaque provider EntryIDs; syntactically
valid bytes need not identify a message. The provider checks that when opening.

Both a nonempty profile name and a zero-based decimal store-table index are
required. Use the existing `-Store` listing mode with the intended `-Profile`
to identify that index; store order can change between sessions. There is no
default-profile or default-store fallback in EntryID mode. The selected store
must own the EntryID. Folder paths, store EntryID selection, internet Message-IDs,
EWS IDs, `.msg` files and attachment IDs are not supported by this mode.

Specify exactly one known property name (case insensitive), or a full eight-digit
hexadecimal property tag, optionally prefixed with `0x`. A property ID alone,
unspecified type, error type, or object type is not accepted. Named-property
names/GUIDs are not resolved; a numeric named-property tag is store-specific.

The store and message are opened without write access, dialogs, or deferred
message-open errors. No properties, read flags or mailbox content are changed.
Existing profile/provider availability and matching MAPI/Outlook bitness are
still required. Providers can have normal logon/network side effects.

The requested property is read with `GetProps` and printed using MrMAPI's
existing property XML representation inside `<messageprops>` on stdout, not as
a bare value. The console uses UTF-16 output. Missing properties produce the
provider's `PT_ERROR` output and a failure exit status. Stream fallback and
all-property export are intentionally not supported.

Only `-NoAddins`, `-Verbose`, `-Online`, and `-Version` may additionally be used.
Verbose/provider diagnostics may accompany the XML. Other mode/filter/output
switches are rejected; this mode does not write files. Normal switch-prefix
matching applies; prefer full switch names in scripts.

Exit status is **0** on successful retrieval (or explicit help), **2** for invalid
EntryID command input before MAPI initialization/logon, and **1** for store,
message or property lookup failure. Existing commands retain their behavior.
As elsewhere, `-Version 0` lists available MAPI implementations and exits
without initializing MAPI or logging on.

Mailbox-free regression tests are in `mmentryidtest::EntryIDTest`, including
real command-line parsing and fake store/message interfaces. Run with the
existing `scripts\vstest.ps1` runner and `/TestCaseFilter:FullyQualifiedName~mmentryidtest`.
Process-level parsing/exit-code checks (also without logon) can be run with:

```powershell
pwsh -File .\UnitTest\scripts\Test-MrMAPIEntryID.ps1 -MrMAPI .\bin\x64\MrMAPI\Debug_Unicode\MrMAPI.exe
```
