#pragma once

enum eAceType
{
	acetypeContainer,
	acetypeMessage,
	acetypeFreeBusy
};

_Check_return_ wstring GetTextualSid(_In_ PSID pSid);
_Check_return_ HRESULT SDToString(_In_count_(cbBuf) const BYTE* lpBuf, size_t cbBuf, eAceType acetype, _In_ wstring& SDString, _In_ wstring& sdInfo);