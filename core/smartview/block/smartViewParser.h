#pragma once
#include <core/smartview/block/binaryParser.h>
#include <core/smartview/block/block.h>

#define _MaxBytes 0xFFFF
#define _MaxDepth 25
#define _MaxEID 500
#define _MaxEntriesSmall 500
#define _MaxEntriesLarge 1000
#define _MaxEntriesExtraLarge 1500
#define _MaxEntriesEnormous 10000

namespace smartview
{
	class smartViewParser : public block
	{
	public:
		smartViewParser() = default;
		virtual ~smartViewParser() = default;
		smartViewParser(const smartViewParser&) = delete;
		smartViewParser& operator=(const smartViewParser&) = delete;

		void init(size_t _cb, _In_count_(_cb) const BYTE* _bin)
		{
			m_Parser = std::make_shared<binaryParser>(_cb, _bin);
			parsed = false;
			enableJunk = true;
		}

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

		bool isSet() const noexcept override { return parsed; }

	protected:
		void ensureParsed();
		std::shared_ptr<binaryParser> m_Parser;
		bool parsed{false};

	private:
		// Consume binaryParser and populate (which may also inherit from smarViewParser)
		virtual void parse() = 0;
		// (optional) Stiches smartViewParser submembers into a tree
		virtual void parseBlocks(){};

		bool enableJunk{true};
	};
} // namespace smartview