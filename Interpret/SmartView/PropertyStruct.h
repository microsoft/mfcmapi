#pragma once
#include <Interpret/SmartView/SmartViewParser.h>
#include <MAPIDefs.h>

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

	class PropertyStruct : public SmartViewParser
	{
	public:
		PropertyStruct();
		void SetMaxEntries(DWORD maxEntries) { m_MaxEntries = maxEntries; }

	private:
		void Parse() override;
		void ParseBlocks() override;

		DWORD m_MaxEntries = _MaxEntriesSmall;
		std::vector<SPropValueStruct> m_Props;

		_Check_return_ SPropValueStruct BinToSPropValueStruct(bool bStringPropsExcludeLength);
	};

	_Check_return_ std::wstring PropsToString(DWORD PropCount, LPSPropValue Prop);
}