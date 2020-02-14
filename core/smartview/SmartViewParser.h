#pragma once
#include <core/smartview/binaryParser.h>
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
	class smartViewParser
	{
	public:
		smartViewParser() = default;
		virtual ~smartViewParser() = default;
		smartViewParser(const smartViewParser&) = delete;
		smartViewParser& operator=(const smartViewParser&) = delete;

		void init(size_t cb, _In_count_(cb) const BYTE* _bin)
		{
			m_Parser = std::make_shared<binaryParser>(cb, _bin);
			parsed = false;
			data = std::make_shared<block>();
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

		std::wstring toString();

		std::shared_ptr<block>& getBlock() noexcept { return data; }
		bool hasData() const { return data->hasData(); }

	protected:
		std::shared_ptr<binaryParser> m_Parser;

		// Nu style parsing data
		template <typename... Args> void setRoot(const std::wstring& text, const Args... args)
		{
			data->setText(text, args...);
		}

		void setRoot(const std::shared_ptr<block>& _data) noexcept { data = _data; }

		template <typename... Args>
		void setRoot(const std::shared_ptr<block>& _data, const std::wstring& text, const Args... args)
		{
			data = _data;
			data->setText(text, args...);
		}

		template <typename... Args> void addHeader(const std::wstring& text, const Args... args)
		{
			data->addHeader(text, args...);
		}

		void addChild(const std::shared_ptr<block>& _block, const std::wstring& text) { data->addChild(_block, text); }
		template <typename... Args>
		void addChild(const std::shared_ptr<block>& _block, const std::wstring& text, const Args... args)
		{
			data->addChild(_block, text, args...);
		}
		void addLabledChild(const std::wstring& _text, const std::shared_ptr<block>& _block)
		{
			data->addLabledChild(_text, _block);
		}
		void addChild(const std::shared_ptr<block>& child) { data->addChild(child); }
		void terminateBlock() { data->terminateBlock(); }
		void addBlankLine() { data->addBlankLine(); }

	private:
		void ensureParsed();
		virtual void parse() = 0;
		virtual void parseBlocks() = 0;

		std::shared_ptr<block> data = std::make_shared<block>();
		bool enableJunk{true};
		bool parsed{false};
	};
} // namespace smartview