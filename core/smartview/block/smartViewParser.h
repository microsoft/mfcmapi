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

		std::wstring toString();

	protected:
		std::shared_ptr<binaryParser> m_Parser;

		// Nu style parsing data
		template <typename... Args> void setRoot(const std::wstring& text, const Args... args)
		{
			block::setText(text, args...);
		}

		void setRoot(const std::shared_ptr<block>& _data) noexcept { addChild(_data); }

		template <typename... Args>
		void setRoot(const std::shared_ptr<block>& _data, const std::wstring& text, const Args... args)
		{
			addChild(_data, text, args...);
		}

		template <typename... Args> void addHeader(const std::wstring& text, const Args... args)
		{
			block::addHeader(text, args...);
		}

		void addChild(const std::shared_ptr<block>& _block, const std::wstring& _text)
		{
			block::addChild(_block, _text);
		}
		template <typename... Args>
		void addChild(const std::shared_ptr<block>& _block, const std::wstring& _text, const Args... args)
		{
			block::addChild(_block, _text, args...);
		}
		void addLabeledChild(const std::wstring& _text, const std::shared_ptr<block>& _block)
		{
			block::addLabeledChild(_text, _block);
		}
		void addChild(const std::shared_ptr<block>& child) { block::addChild(child); }
		void terminateBlock() { block::terminateBlock(); }
		void addBlankLine() { block::addBlankLine(); }

	private:
		void ensureParsed();
		virtual void parse() = 0;
		virtual void parseBlocks() = 0;

		bool enableJunk{true};
		bool parsed{false};
	};
} // namespace smartview