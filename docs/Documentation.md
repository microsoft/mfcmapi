# MFCMAPI Resources
[![MFCMAPI on Facebook](https://badge.facebook.com/badge/26764016480.2776.1538253884.png)](https://www.facebook.com/MFCMAPI/)<br/>
[A Brief History of MFCMAPI](http://web.archive.org/web/20190117143305/https://blogs.msdn.microsoft.com/stephen_griffin/2008/01/04/a-brief-history-of-mfcmapi/)<br/>
[MFCMAPI Articles](MFCMAPIArticles.md)<br/>
[Blog entries on MFCMAPI/MAPI Editor](https://docs.microsoft.com/en-us/archive/blogs/stephen_griffin/) (archived)<br/>

## MFCMAPI Add-Ins
[MFCMAPI Addin Development](MFCMAPIAddinDevelopment.md)<br/>
[Sample add-in for creating contacts, appointments, tasks, etc.](CreateOutlookItemsAddin.md)

## Debugging symbols

When debugging MAPI or Outlook, add Microsoft's public symbol server to your debugger's symbol path:

```text
https://msdl.microsoft.com/download/symbols
```

For MFCMAPI and MrMAPI, matching PDB archives are available alongside the executable archives in the [GitHub releases](https://github.com/microsoft/mfcmapi/releases). Download the PDB archive for the same release and architecture as the executable, then extract the PDB alongside the EXE.

### Visual Studio

Under **Tools > Options > Debugging > Symbols**, add the URL as a symbol-file location. While debugging, use **Debug > Windows > Modules** to load symbols and inspect **Symbol Load Information**.

### WinDbg

Append the server to your symbol path, using a local cache directory, then reload symbols:

```text
.sympath+ srv*C:\Symbols*https://msdl.microsoft.com/download/symbols
.reload
```

## General MAPI/Exchange Development Resources
[MSDN MAPI Documentation](https://msdn.microsoft.com/en-us/library/office/cc815646.aspx)<br/>
[Exchange Protocol Documentation](https://msdn.microsoft.com/en-us/library/cc307725.aspx)<br/>
[MAPI Mail List](http://peach.ease.lsoft.com/archives/mapi-l.html)<br/>
[ExchangeDev Blog](https://docs.microsoft.com/en-us/archive/blogs/exchangedev/) (archived)<br/>
[EHLO Blog](https://blogs.technet.microsoft.com/exchange/)<br/>
![EWSEditor](https://github.com/dseph/EwsEditor/raw/master/EWSEditor/Resources/EWSEditor2.png)[EWSEditor](https://github.com/dseph/EwsEditor)
