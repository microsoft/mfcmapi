#pragma once

namespace sid
{
	enum eAceType
	{
		acetypeContainer,
		acetypeMessage,
		acetypeFreeBusy
	};

	struct SidAccount
	{
		std::wstring domain;
		std::wstring name;
	};

	struct SecurityDescriptor
	{
		std::wstring dacl;
		std::wstring info;
	};

	_Check_return_ std::wstring GetTextualSid(_In_ PSID pSid);
	_Check_return_ SidAccount LookupAccountSid(PSID SidStart);
	_Check_return_ SecurityDescriptor SDToString(_In_count_(cbBuf) const BYTE* lpBuf, size_t cbBuf, eAceType acetype);
} // namespace sid