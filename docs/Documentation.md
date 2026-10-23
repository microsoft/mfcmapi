# MFCMAPI Resources
[![MFCMAPI on Facebook](https://badge.facebook.com/badge/26764016480.2776.1538253884.png)](https://www.facebook.com/MFCMAPI/)<br/>
[A Brief History of MFCMAPI](http://web.archive.org/web/20190117143305/https://blogs.msdn.microsoft.com/stephen_griffin/2008/01/04/a-brief-history-of-mfcmapi/)<br/>
[MFCMAPI Articles](MFCMAPIArticles.md)<br/>
[Blog entries on MFCMAPI/MAPI Editor](https://docs.microsoft.com/en-us/archive/blogs/stephen_griffin/) (archived)<br/>

## MFCMAPI Add-Ins
[MFCMAPI Addin Development](MFCMAPIAddinDevelopment.md)<br/>
[Sample add-in for creating contacts, appointments, tasks, etc.](CreateOutlookItemsAddin.md)

## MrMAPI: another user's Exchange mailbox

Add `-Mailbox <primary SMTP address>` and an explicit `-Profile <profile name>`
to a store/folder command. For example:

```text
MrMAPI -Mailbox user@example.com -Profile "Service profile" -Store
MrMAPI -Mailbox user@example.com -Profile "Service profile" -Contents -Folder 13 -List
MrMAPI -Mailbox user@example.com -Profile "Service profile" -NamedProps
MrMAPI -Mailbox user@example.com -Profile "Service profile" -ChildFolders -Folder "Inbox"
```

Run under the Windows identity that owns the existing Exchange profile. The switch
does not accept credentials, impersonate another identity, grant permissions, or
request administrative privileges. The service/delegate account must already have
the required Exchange mailbox permissions. Match MrMAPI's bitness to Outlook.

MrMAPI resolves the supplied address without an address-book dialog, verifies the
resolved Exchange mail user's **primary** SMTP address, then uses the profile's
primary Exchange store and existing `CreateStoreEntryID2`/mailbox-logon helpers.
Aliases, ambiguous/near matches, non-Exchange recipients and missing identity
properties are rejected. The provider must expose the recipient's primary SMTP
address and support the Exchange manage-store interfaces. Authentication support,
cross-server routing and permissions remain provider/server dependent; this is not
a promise that every modern Exchange tenant allows this operation.

Store indices, `-ReceiveFolder` indices and `#store` folder prefixes cannot be
combined with `-Mailbox`: every selected folder must belong to the requested
mailbox. With `-Contents`, the default folder is the Inbox (13). Existing folder
paths and default-folder numbers otherwise retain their usual meaning.
The store is temporary (not added to the profile); these inspection commands do
not write mailbox data. The existing mailbox-logon helper requests best available
access, not an enforced read-only store handle.

Exit code 2 means invalid mailbox options; 1 means initialization/logon,
resolution, mailbox-open or folder-open failure; 0 means the command was
dispatched. Errors occurring inside an existing inspection command keep that
command's existing output/exit semantics. No failed request falls back to the
profile owner's mailbox. Commands without `-Mailbox` are unchanged.

## General MAPI/Exchange Development Resources
[MSDN MAPI Documentation](https://msdn.microsoft.com/en-us/library/office/cc815646.aspx)<br/>
[Exchange Protocol Documentation](https://msdn.microsoft.com/en-us/library/cc307725.aspx)<br/>
[MAPI Mail List](http://peach.ease.lsoft.com/archives/mapi-l.html)<br/>
[ExchangeDev Blog](https://docs.microsoft.com/en-us/archive/blogs/exchangedev/) (archived)<br/>
[EHLO Blog](https://blogs.technet.microsoft.com/exchange/)<br/>
![EWSEditor](https://github.com/dseph/EwsEditor/raw/master/EWSEditor/Resources/EWSEditor2.png)[EWSEditor](https://github.com/dseph/EwsEditor)
