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
	class SmartViewParser;
	typedef SmartViewParser FAR* LPSMARTVIEWPARSER;

	class SmartViewParser
	{
	public:
		SmartViewParser() = default;
		virtual ~SmartViewParser() = default;
		SmartViewParser(const SmartViewParser&) = delete;
		SmartViewParser& operator=(const SmartViewParser&) = delete;

		void init(size_t cbBin, _In_count_(cbBin) const BYTE* lpBin)
		{
			m_Parser = std::make_shared<binaryParser>(cbBin, lpBin);
			m_bParsed = false;
			data = {};
			m_bEnableJunk = true;
		}

		void parse(std::shared_ptr<binaryParser> binaryParser, bool bDoJunk) { parse(binaryParser, 0, bDoJunk); }

		void parse(std::shared_ptr<binaryParser> binaryParser, size_t cbBin, bool bEnableJunk)
		{
			m_Parser = binaryParser;
			m_Parser->setCap(cbBin);
			m_bEnableJunk = bEnableJunk;
			EnsureParsed();
			m_Parser->clearCap();
		}

		_Check_return_ std::wstring ToString();

		block& getBlock() { return data; }
		bool hasData() const { return data.hasData(); }

	protected:
		std::shared_ptr<binaryParser> m_Parser;

		// Nu style parsing data
		template <typename... Args> void setRoot(const std::wstring& text, const Args... args)
		{
			data.setText(text, args...);
		}

		void setRoot(const block& _data) { data = _data; }

		template <typename... Args> void setRoot(const block& _data, const std::wstring& text, const Args... args)
		{
			data = _data;
			data.setText(text, args...);
		}

		template <typename... Args> void addHeader(const std::wstring& text, const Args... args)
		{
			data.addHeader(text, args...);
		}

		void addBlock(block& _block, const std::wstring& text) { data.addBlock(_block, text); }
		template <typename... Args> void addBlock(block& _block, const std::wstring& text, const Args... args)
		{
			data.addBlock(_block, text, args...);
		}
		void addBlock(block& child) { data.addBlock(child); }
		void terminateBlock() { data.terminateBlock(); }
		void addBlankLine() { data.addBlankLine(); }

	private:
		void EnsureParsed();
		virtual void Parse() = 0;
		virtual void ParseBlocks() = 0;

		block data;
		bool m_bEnableJunk{true};
		bool m_bParsed{false};
	};
} // namespace smartview