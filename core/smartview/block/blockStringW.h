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

		bool isSet() const noexcept override { return parsed; }

		// Mimic std::wstring
		operator const std::wstring&() const noexcept { return data; }
		_NODISCARD _Ret_z_ const wchar_t* c_str() const noexcept { return data.c_str(); }
		_NODISCARD std::wstring::size_type length() const noexcept { return data.length(); }
		_NODISCARD bool empty() const noexcept { return data.empty(); }

		static std::shared_ptr<blockStringW> parse(const std::wstring& _data, size_t _size, size_t _offset)
		{
			auto ret = std::make_shared<blockStringW>();
			ret->parsed = true;
			ret->data = _data;
			ret->setSize(_size);
			ret->setOffset(_offset);
			return ret;
		}

		static std::shared_ptr<blockStringW> parse(const std::shared_ptr<binaryParser>& parser, size_t cchChar = -1)
		{
			auto ret = std::make_shared<blockStringW>();
			ret->m_Parser = parser;
			ret->cchChar = cchChar;
			ret->ensureParsed();
			return ret;
		}

		void parse() override
		{
			parsed = false;
			if (cchChar == static_cast<size_t>(-1))
			{
				cchChar =
					wcsnlen_s(reinterpret_cast<LPCWSTR>(m_Parser->getAddress()), (m_Parser->getSize()) / sizeof WCHAR) +
					1;
			}

			if (cchChar && m_Parser->checkSize(sizeof WCHAR * cchChar))
			{
				data = strings::RemoveInvalidCharactersW(
					std::wstring(reinterpret_cast<LPCWSTR>(m_Parser->getAddress()), cchChar));
				m_Parser->advance(sizeof WCHAR * cchChar);
				parsed = true;
			}
		};

	private:
		std::wstring toStringInternal() const override { return data; }
		std::wstring data;

		size_t cchChar{};
	};

	inline std::shared_ptr<blockStringW> emptySW() { return std::make_shared<blockStringW>(); }
} // namespace smartview