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
	public:
		SidAccount() = default;
		SidAccount(std::wstring _domain, std::wstring _name) noexcept
			: domain(std::move(_domain)), name(std::move(_name)){};
		_Check_return_ std::wstring getDomain() const;
		_Check_return_ std::wstring getName() const;

	private:
		std::wstring domain;
		std::wstring name;
	};

	struct SecurityDescriptor
	{
		std::wstring dacl;
		std::wstring info;
	};

	_Check_return_ std::wstring GetTextualSid(_In_ PSID pSid);
	_Check_return_ std::wstring GetTextualSid(std::vector<BYTE> buf);
	_Check_return_ SidAccount LookupAccountSid(PSID SidStart);
	_Check_return_ SidAccount LookupAccountSid(std::vector<BYTE> buf);
	_Check_return_ SecurityDescriptor SDToString(const std::vector<BYTE>& buf, eAceType acetype);
} // namespace sid