# Viewing template attachments

The attachment table can inspect Outlook template (`.oft`) and message (`.msg`)
files attached **by value**. These files contain MSG compound storage in
`PR_ATTACH_DATA_BIN`, not an embedded `IMessage` in `PR_ATTACH_DATA_OBJ`.

Double-click a recognized attachment to display its message, or enable **View
attachments as embedded message** to inspect it in the property pane. Recognition
uses the long filename (falling back to the short filename), with a
case-insensitive, exact `.oft` or `.msg` extension. The payload must also open as
OLE compound storage and an MSG message; an extension alone is not sufficient.
Failures are reported through the normal error dialog.

The payload is copied into memory without saving or executing a file. Changes to
this in-memory message do not write back to the original attachment. Genuine
embedded messages retain their existing editing behavior. Other attachments
retain their existing attachment-property and save-to-file behavior.
