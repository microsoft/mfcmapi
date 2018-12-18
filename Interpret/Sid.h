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
		SidAccount(const std::wstring& _domain, const std::wstring& _name) : domain(_domain), name(_name){};
		_Check_return_ std::wstring getDomain() const
		{
			return !domain.empty() ? domain : strings::formatmessage(IDS_NODOMAIN);
		}
		_Check_return_ std::wstring getName() const
		{
			return !name.empty() ? name : strings::formatmessage(IDS_NONAME);
		}

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
	_Check_return_ SecurityDescriptor SDToString(const std::vector<BYTE> buf, eAceType acetype);
} // namespace sid