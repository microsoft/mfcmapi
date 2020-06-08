#pragma once
#include <core/smartview/block/block.h>

namespace smartview
{
	class smartViewParser : public block
	{
	public:
		smartViewParser() = default;
		virtual ~smartViewParser() = default;
		smartViewParser(const smartViewParser&) = delete;
		smartViewParser& operator=(const smartViewParser&) = delete;

		virtual void parse(const std::shared_ptr<binaryParser>& binaryParser, bool bDoJunk)
		{
			parse(binaryParser, 0, bDoJunk);
		}

		virtual void parse(const std::shared_ptr<binaryParser>& binaryParser, size_t cbBin, bool _enableJunk)
		{
			m_Parser = binaryParser;
			m_Parser->setCap(cbBin);
			enableJunk = _enableJunk;
			ensureParsed();
			m_Parser->clearCap();
		}

		template <typename T>
		static std::shared_ptr<T> parse(const std::shared_ptr<binaryParser>& binaryParser, bool _enableJunk)
		{
			return parse<T>(binaryParser, 0, _enableJunk);
		}

		template <typename T>
		static std::shared_ptr<T>
		parse(const std::shared_ptr<binaryParser>& binaryParser, size_t cbBin, bool _enableJunk)
		{
			auto ret = std::make_shared<T>();
			ret->smartViewParser::parse(binaryParser, cbBin, _enableJunk);
			return ret;
		}

		std::wstring toString();

	protected:
		void ensureParsed();
		// Temp until we finish moving everything
		void parse() override {}
	};
} // namespace smartview