#pragma once

namespace sid
{
	enum class aceType
	{
		Container,
		Message,
		FreeBusy
	};

	struct SidAccount
	{
	public:
		SidAccount() = default;
		SidAccount(std::wstring _domain, std::wstring _name) noexcept
			: domain(std::move(_domain)), name(std::move(_name)) {};
		_Check_return_ std::wstring getDomain() const;
		_Check_return_ std::wstring getName() const;

	private:
		std::wstring domain;
		std::wstring name;
	};

	_Check_return_ std::wstring LookupIdentifierAuthority(const SID_IDENTIFIER_AUTHORITY& authority);
	_Check_return_ std::wstring IdentifierAuthorityToString(const SID_IDENTIFIER_AUTHORITY& authority);
	_Check_return_ SidAccount LookupAccountSid(std::vector<BYTE> buf);
	_Check_return_ std::wstring NTSDToString(const std::vector<BYTE>& buf, aceType acetype);
} // namespace sid