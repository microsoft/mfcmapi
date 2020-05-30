#pragma once
#include <core/smartview/block/smartViewParser.h>

namespace smartview
{
	class blockStringA : public smartViewParser
	{
	public:
		blockStringA() = default;
		blockStringA(const blockStringA&) = delete;
		blockStringA& operator=(const blockStringA&) = delete;

		bool isSet() const noexcept override { return set; }

		// Mimic std::string
		operator const std::string&() const noexcept { return data; }
		_NODISCARD _Ret_z_ const char* c_str() const noexcept { return data.c_str(); }
		_NODISCARD std::string::size_type length() const noexcept { return data.length(); }
		_NODISCARD bool empty() const noexcept { return data.empty(); }

		blockStringA(const std::shared_ptr<binaryParser>& parser, size_t _cchChar = -1)
		{
			m_Parser = parser;
			cchChar = _cchChar;
			ensureParsed();
		}

		static std::shared_ptr<blockStringA> parse(const std::shared_ptr<binaryParser>& parser, size_t cchChar = -1)
		{
			return std::make_shared<blockStringA>(parser, cchChar);
		}

		void parse() override
		{
			if (cchChar == static_cast<size_t>(-1))
			{
				cchChar =
					strnlen_s(reinterpret_cast<LPCSTR>(m_Parser->getAddress()), (m_Parser->getSize()) / sizeof CHAR) +
					1;
			}

			if (cchChar && m_Parser->checkSize(sizeof CHAR * cchChar))
			{
				setOffset(m_Parser->getOffset());
				data = strings::RemoveInvalidCharactersA(
					std::string(reinterpret_cast<LPCSTR>(m_Parser->getAddress()), cchChar));
				setSize(sizeof CHAR * cchChar);
				m_Parser->advance(sizeof CHAR * cchChar);
				set = true;
			}
		};
		void parseBlocks() override{};

	private:
		std::wstring toStringInternal() const override { return strings::stringTowstring(data); }
		std::string data;
		bool set{false};

		size_t cchChar{};
	};

	inline std::shared_ptr<blockStringA> emptySA() { return std::make_shared<blockStringA>(); }
} // namespace smartview