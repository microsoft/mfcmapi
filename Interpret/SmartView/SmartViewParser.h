#pragma once
#include <Interpret/SmartView/BinaryParser.h>

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
		SmartViewParser();
		virtual ~SmartViewParser() = default;

		void init(size_t cbBin, _In_count_(cbBin) const BYTE* lpBin);
		_Check_return_ std::wstring ToString();

		void DisableJunkParsing();
		size_t GetCurrentOffset() const;
		void EnsureParsed();
		const block& getBlock() const { return data; }
		bool hasData() const { return data.hasData(); }

	protected:
		CBinaryParser m_Parser;

		// Nu style parsing data
		block data;
		template <typename... Args> void setRoot(const std::wstring& text, const Args... args)
		{
			data.setText(text, args...);
		}

		template <typename... Args> void addHeader(const std::wstring& text, const Args... args)
		{
			data.addHeader(text, args...);
		}

		void addBlock(const block& _block, const std::wstring& text) { data.addBlock(_block, text); }
		template <typename... Args> void addBlock(const block& _block, const std::wstring& text, const Args... args)
		{
			data.addBlock(_block, text, args...);
		}
		void addBlock(const block& child) { data.addBlock(child); }
		void addLine() { data.addLine(); }

	private:
		virtual void Parse() = 0;
		virtual void ParseBlocks() = 0;

		bool m_bEnableJunk;
		bool m_bParsed;
	};
} // namespace smartview