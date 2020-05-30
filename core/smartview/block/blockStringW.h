#pragma once
#include <core/smartview/block/smartViewParser.h>

namespace smartview
{
	class blockStringW : public smartViewParser
	{
	public:
		blockStringW() = default;
		blockStringW(const blockStringW&) = delete;
		blockStringW& operator=(const blockStringW&) = delete;

		bool isSet() const noexcept override { return set; }

		// Mimic std::wstring
		operator const std::wstring&() const noexcept { return data; }
		_NODISCARD _Ret_z_ const wchar_t* c_str() const noexcept { return data.c_str(); }
		_NODISCARD std::wstring::size_type length() const noexcept { return data.length(); }
		_NODISCARD bool empty() const noexcept { return data.empty(); }

		blockStringW(const std::shared_ptr<binaryParser>& parser, size_t cchChar = -1)
		{
			if (cchChar == static_cast<size_t>(-1))
			{
				cchChar =
					wcsnlen_s(reinterpret_cast<LPCWSTR>(parser->getAddress()), (parser->getSize()) / sizeof WCHAR) + 1;
			}

			if (cchChar && parser->checkSize(sizeof WCHAR * cchChar))
			{
				setOffset(parser->getOffset());
				data = strings::RemoveInvalidCharactersW(
					std::wstring(reinterpret_cast<LPCWSTR>(parser->getAddress()), cchChar));
				setSize(sizeof WCHAR * cchChar);
				parser->advance(sizeof WCHAR * cchChar);
				set = true;
			}
		}

		static std::shared_ptr<blockStringW> parse(const std::shared_ptr<binaryParser>& parser, size_t cchChar = -1)
		{
			return std::make_shared<blockStringW>(parser, cchChar);
		}

		void parse() override{};
		void parseBlocks() override{};

	private:
		std::wstring toStringInternal() const override { return data; }
		std::wstring data;
		bool set{false};
	};

	inline std::shared_ptr<blockStringW> emptySW() { return std::make_shared<blockStringW>(); }
} // namespace smartview