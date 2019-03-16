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

		virtual bool isSet() const { return set; }

		// Mimic std::wstring
		operator const std::wstring&() const { return data; }
		_NODISCARD _Ret_z_ const wchar_t* c_str() const noexcept { return data.c_str(); }
		_NODISCARD std::wstring::size_type length() const noexcept { return data.length(); }
		_NODISCARD bool empty() const noexcept { return data.empty(); }

		blockStringW(std::shared_ptr<binaryParser> parser, size_t cchChar = -1)
		{
			set = true;
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
				data = strings::RemoveInvalidCharactersW(
					std::wstring(reinterpret_cast<LPCWSTR>(parser->GetCurrentAddress()), cchChar));
				setSize(sizeof WCHAR * cchChar);
				parser->advance(sizeof WCHAR * cchChar);
			}
		}

		blockStringW(const std::wstring& _data, size_t _offset, size_t _size)
		{
			set = true;
			data = _data;
			setOffset(_offset);
			setSize(_size);
		}

		static std::shared_ptr<blockStringW> parse(std::shared_ptr<binaryParser> parser, size_t cchChar = -1)
		{
			return std::make_shared<blockStringW>(parser, cchChar);
		}

	private:
		std::wstring data;
		bool set{false};
	};

	inline std::shared_ptr<blockStringW> emptySW() { return std::make_shared<blockStringW>(); }
} // namespace smartview