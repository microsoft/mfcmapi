#pragma once

namespace sid
{
	enum eAceType
	{
		acetypeContainer,
		acetypeMessage,
		acetypeFreeBusy
	};

	_Check_return_ std::wstring GetTextualSid(_In_ PSID pSid);
	_Check_return_ std::wstring LookupAccountSid(PSID SidStart, _In_ std::wstring& sidDomain);
	_Check_return_ std::wstring
	SDToString(_In_count_(cbBuf) const BYTE* lpBuf, size_t cbBuf, eAceType acetype, _In_ std::wstring& sdInfo);
}