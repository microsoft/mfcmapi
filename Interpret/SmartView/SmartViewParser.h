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
	struct FILETIMEBLock
	{
		blockT<DWORD> dwLowDateTime;
		blockT<DWORD> dwHighDateTime;
		operator FILETIME() const { return FILETIME{dwLowDateTime, dwHighDateTime}; }
	};

	struct SBinaryBlock
	{
		blockT<ULONG> cb;
		blockBytes lpb;
	};

	struct StringArrayA
	{
		blockT<ULONG> cValues;
		std::vector<blockStringA> lppszA;
	};

	struct StringArrayW
	{
		blockT<ULONG> cValues;
		std::vector<blockStringW> lppszW;
	};

	struct PVBlock
	{
		blockT<LONG> l; /* case PT_LONG */
		blockT<unsigned short int> b; /* case PT_BOOLEAN */
		FILETIMEBLock ft; /* case PT_SYSTIME */
		blockStringA lpszA; /* case PT_STRING8 */
		SBinaryBlock bin; /* case PT_BINARY */
		blockStringW lpszW; /* case PT_UNICODE */
		SBinaryArray MVbin; /* case PT_MV_BINARY */
		StringArrayA MVszA; /* case PT_MV_STRING8 */
		StringArrayW MVszW; /* case PT_MV_UNICODE */
		blockT<SCODE> err; /* case PT_ERROR */
	};

	struct SPropValueStruct
	{
		blockT<WORD> PropType;
		blockT<WORD> PropID;
		blockT<ULONG> ulPropTag;
		ULONG dwAlignPad;
		PVBlock Value;

		SPropValue const getData()
		{
			auto prop = SPropValue{};
			prop.ulPropTag = ulPropTag;
			prop.dwAlignPad = dwAlignPad;
			switch (PropType)
			{
			case PT_LONG:
				prop.Value.l = Value.l;
				break;
			case PT_BOOLEAN:
				prop.Value.b = Value.b;
				break;
			case PT_SYSTIME:
				prop.Value.ft = Value.ft;
				break;
			case PT_STRING8:
				prop.Value.lpszA = const_cast<LPSTR>(Value.lpszA.c_str());
				break;
			case PT_BINARY:
				prop.Value.bin.cb = Value.bin.cb;
				prop.Value.bin.lpb = const_cast<LPBYTE>(Value.bin.lpb.data());
				break;
			case PT_UNICODE:
				prop.Value.lpszW = const_cast<LPWSTR>(Value.lpszW.c_str());
				break;
				//SBinaryArray MVbin; /* case PT_MV_BINARY */
				//StringArrayA MVszA; /* case PT_MV_STRING8 */
				//StringArrayW MVszW; /* case PT_MV_UNICODE */
			case PT_ERROR:
				prop.Value.err = Value.err;
				break;
			}

			return prop;
		}
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
		block getBlock() const { return data; }
		blockBytes getJunkData() { return m_Parser.GetBlockRemainingData(); }
		bool hasData() { return data.hasData(); }

	protected:
		_Check_return_ std::wstring JunkDataToString(const std::vector<BYTE>& lpJunkData) const;
		_Check_return_ std::wstring
		JunkDataToString(size_t cbJunkData, _In_count_(cbJunkData) const BYTE* lpJunkData) const;
		_Check_return_ LPSPropValue BinToSPropValue(DWORD dwPropCount, bool bStringPropsExcludeLength);
		_Check_return_ SPropValueStruct BinToSPropValueStruct(bool bStringPropsExcludeLength);

		// These functions return pointers to memory backed and cleaned up by SmartViewParser
		LPBYTE GetBYTES(size_t cbBytes);
		LPSTR GetStringA(size_t cchChar = -1);
		LPWSTR GetStringW(size_t cchChar = -1);
		LPBYTE Allocate(size_t cbBytes);
		LPBYTE AllocateArray(size_t cArray, size_t cbEntry);

		CBinaryParser m_Parser;

		// Nu style parsing data
		block data;
		template <typename... Args> void addHeader(const std::wstring& text, Args... args)
		{
			data.addHeader(text, args...);
		}
		void addBlock(const block& _block, const std::wstring& text) { data.addBlock(_block, text); }
		template <typename... Args> void addBlock(const block& _block, const std::wstring& text, Args... args)
		{
			data.addBlock(_block, text, args...);
		}
		void addBlock(const block& child) { data.addBlock(child); }
		void addBlockBytes(const blockBytes& _block) { data.addBlockBytes(_block); }
		void addLine() { data.addLine(); }

	private:
		virtual void Parse() = 0;
		// TODO: make this = 0 to ensure everyone has a ParseBlocks implementation
		virtual void ParseBlocks() {}
		virtual _Check_return_ std::wstring ToStringInternal() { return data.ToString(); }

		bool m_bEnableJunk;
		bool m_bParsed;

		// We use list instead of vector so our nodes never get reallocated
		std::list<std::string> m_stringCache;
		std::list<std::wstring> m_wstringCache;
		std::list<std::vector<BYTE>> m_binCache;
	};
}