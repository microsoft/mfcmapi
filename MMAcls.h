#pragma once
// MMAcls.h : Acl table dumping for MrMAPI

void DoAcls(_In_ MYOPTIONS ProgOpts);
void DumpExchangeTable(_In_z_ LPWSTR lpszProfile, _In_ ULONG ulPropTag, _In_ ULONG ulFolder, _In_z_ LPWSTR lpszFolder);