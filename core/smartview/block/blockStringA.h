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

		bool isSet() const noexcept override { return parsed; }

		// Mimic std::string
		operator const std::string&() const noexcept { return data; }
		_NODISCARD _Ret_z_ const char* c_str() const noexcept { return data.c_str(); }
		_NODISCARD std::string::size_type length() const noexcept { return data.length(); }
		_NODISCARD bool empty() const noexcept { return data.empty(); }

		static std::shared_ptr<blockStringA> parse(const std::shared_ptr<binaryParser>& parser, size_t cchChar = -1)
		{
			auto ret = std::make_shared<blockStringA>();
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
					strnlen_s(reinterpret_cast<LPCSTR>(m_Parser->getAddress()), (m_Parser->getSize()) / sizeof CHAR) +
					1;
			}

			if (cchChar && m_Parser->checkSize(sizeof CHAR * cchChar))
			{
				data = strings::RemoveInvalidCharactersA(
					std::string(reinterpret_cast<LPCSTR>(m_Parser->getAddress()), cchChar));
				m_Parser->advance(sizeof CHAR * cchChar);
				parsed = true;
			}
		};

	private:
		std::wstring toStringInternal() const override { return strings::stringTowstring(data); }
		std::string data;

		size_t cchChar{};
	};

	inline std::shared_ptr<blockStringA> emptySA() { return std::make_shared<blockStringA>(); }
} // namespace smartview