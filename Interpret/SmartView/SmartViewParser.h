#pragma once
#include <Interpret/SmartView/BinaryParser.h>

#define _MaxBytes 0xFFFF
#define _MaxDepth 50
#define _MaxEID 500
#define _MaxEntriesSmall 500
#define _MaxEntriesLarge 1000
#define _MaxEntriesExtraLarge 1500
#define _MaxEntriesEnormous 10000

namespace smartview
{
	class block
	{
	public:
		block() : header(true), offset(0), cb(0), text(L"") {}
		block(size_t _offset, size_t _cb, std::wstring _text) : header(false), offset(_offset), cb(_cb), text(_text) {}
		block(std::wstring _text) : header(true), offset(0), cb(0), text(_text) {}
		std::wstring ToString()
		{
			std::vector<std::wstring> items;
			items.emplace_back(text);
			for (auto item : children)
			{
				items.emplace_back(item.ToString());
			}

			return strings::join(items, strings::emptystring);
		}
		size_t GetSize() { return cb; }
		size_t GetOffset() { return offset; }
		void AddChild(block child) { children.push_back(child); }

	private:
		std::vector<block> children;
		bool header;
		size_t offset;
		size_t cb;
		std::wstring text;
	};

	class SmartViewParser;
	typedef SmartViewParser FAR* LPSMARTVIEWPARSER;

	class SmartViewParser
	{
	public:
		SmartViewParser();
		virtual ~SmartViewParser() = default;

		void Init(size_t cbBin, _In_count_(cbBin) const BYTE* lpBin);
		_Check_return_ std::wstring ToString();

		void DisableJunkParsing();
		size_t GetCurrentOffset() const;
		void EnsureParsed();

	protected:
		_Check_return_ std::wstring JunkDataToString(const std::vector<BYTE>& lpJunkData) const;
		_Check_return_ std::wstring
		JunkDataToString(size_t cbJunkData, _In_count_(cbJunkData) const BYTE* lpJunkData) const;
		_Check_return_ LPSPropValue BinToSPropValue(DWORD dwPropCount, bool bStringPropsExcludeLength);

		// These functions return pointers to memory backed and cleaned up by SmartViewParser
		LPBYTE GetBYTES(size_t cbBytes);
		LPSTR GetStringA(size_t cchChar = -1);
		LPWSTR GetStringW(size_t cchChar = -1);
		LPBYTE Allocate(size_t cbBytes);
		LPBYTE AllocateArray(size_t cArray, size_t cbEntry);

		CBinaryParser m_Parser;

		// Nu style parsing data
		block data;
		size_t dataOffset;

		std::wstring dataToString() { return data.ToString(); }

		void addHeader(std::wstring text) { data.AddChild(block(text)); }
		void addData(size_t cb, std::wstring text)
		{
			data.AddChild(block(dataOffset, cb, text));
			dataOffset += cb;
		}

		void addBytes(const std::vector<BYTE>& bytes, bool bPrependCB = true)
		{
			const auto cb = bytes.size() * sizeof(BYTE);
			data.AddChild(block(dataOffset, cb, strings::BinToHexString(bytes, bPrependCB)));
			dataOffset += cb;
		}

		void addBlock(block& _block) { data.AddChild(_block); }

	private:
		virtual void Parse() = 0;
		virtual _Check_return_ std::wstring ToStringInternal() { return dataToString(); }

		bool m_bEnableJunk;
		bool m_bParsed;

		// We use list instead of vector so our nodes never get reallocated
		std::list<std::string> m_stringCache;
		std::list<std::wstring> m_wstringCache;
		std::list<std::vector<BYTE>> m_binCache;
	};
}