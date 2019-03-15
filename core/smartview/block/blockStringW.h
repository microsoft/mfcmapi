#pragma once
#include <core/smartview/block/block.h>

namespace smartview
{
	class blockStringW : public block
	{
	public:
		blockStringW() = default;
		blockStringW(const blockStringW&) = delete;
		blockStringW& operator=(const blockStringW&) = delete;
		void setData(const std::wstring& _data)
		{
			if (this) data = _data;
		}
		operator const std::wstring&() const { return this ? data : strings::emptystring; }
		_NODISCARD _Ret_z_ const wchar_t* c_str() const noexcept { return this ? data.c_str() : L""; }
		_NODISCARD std::wstring::size_type length() const noexcept { return this ? data.length() : 0; }
		_NODISCARD bool empty() const noexcept { return this ? data.empty() : true; }

		blockStringW(std::shared_ptr<binaryParser> parser, size_t cchChar = -1)
		{
			if (cchChar == static_cast<size_t>(-1))
			{
				cchChar = wcsnlen_s(
							  reinterpret_cast<LPCWSTR>(parser->GetCurrentAddress()),
							  (parser->RemainingBytes()) / sizeof WCHAR) +
						  1;
			}

			if (cchChar && parser->CheckRemainingBytes(sizeof WCHAR * cchChar))
			{
				setOffset(parser->GetCurrentOffset());
				setData(strings::RemoveInvalidCharactersW(
					std::wstring(reinterpret_cast<LPCWSTR>(parser->GetCurrentAddress()), cchChar)));
				setSize(sizeof WCHAR * cchChar);
				parser->advance(sizeof WCHAR * cchChar);
			}
		}

		static std::shared_ptr<blockStringW> parse(std::shared_ptr<binaryParser> parser, size_t cchChar = -1)
		{
			return std::make_shared<blockStringW>(parser, cchChar);
		}

	private:
		std::wstring data;
	};
} // namespace smartview