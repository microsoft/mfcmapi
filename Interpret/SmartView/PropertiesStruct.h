#pragma once
#include <Interpret/SmartView/SmartViewParser.h>
#include <MAPIDefs.h>
#include <Interpret/InterpretProp.h>
#include <Interpret/SmartView/SmartView.h>

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

	struct SBinaryArrayBlock
	{
		blockT<ULONG> cValues;
		std::vector<SBinaryBlock> lpbin;
	};

	struct CountedStringA
	{
		blockT<DWORD> cb;
		blockStringA str;
	};

	struct CountedStringW
	{
		blockT<DWORD> cb;
		blockStringW str;
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
		blockT<WORD> i; /* case PT_I2 */
		blockT<LONG> l; /* case PT_LONG */
		blockT<WORD> b; /* case PT_BOOLEAN */
		blockT<float> flt; /* case PT_R4 */
		blockT<double> dbl; /* case PT_DOUBLE */
		FILETIMEBLock ft; /* case PT_SYSTIME */
		CountedStringA lpszA; /* case PT_STRING8 */
		SBinaryBlock bin; /* case PT_BINARY */
		CountedStringW lpszW; /* case PT_UNICODE */
		blockT<GUID> lpguid; /* case PT_CLSID */
		blockT<LARGE_INTEGER> li; /* case PT_I8 */
		SBinaryArrayBlock MVbin; /* case PT_MV_BINARY */
		StringArrayA MVszA; /* case PT_MV_STRING8 */
		StringArrayW MVszW; /* case PT_MV_UNICODE */
		blockT<SCODE> err; /* case PT_ERROR */
	};

	struct SPropValueStruct
	{
		blockT<WORD> PropType;
		blockT<WORD> PropID;
		blockT<ULONG> ulPropTag;
		ULONG dwAlignPad{};
		PVBlock Value;
		_Check_return_ const std::wstring PropString()
		{
			EnsurePropStrings();
			return propString;
		}
		_Check_return_ const std::wstring AltPropString()
		{
			EnsurePropStrings();
			return altPropString;
		}

		// TODO: Fill in missing cases with test coverage
		SPropValue const getData()
		{
			auto prop = SPropValue{};
			prop.ulPropTag = ulPropTag;
			prop.dwAlignPad = dwAlignPad;
			switch (PropType)
			{
			case PT_I2:
				prop.Value.i = Value.i;
				break;
			case PT_LONG:
				prop.Value.l = Value.l;
				break;
			case PT_R4:
				prop.Value.flt = Value.flt;
				break;
			case PT_DOUBLE:
				prop.Value.dbl = Value.dbl;
				break;
			case PT_BOOLEAN:
				prop.Value.b = Value.b;
				break;
			case PT_I8:
				prop.Value.li = Value.li.getData();
				break;
			case PT_SYSTIME:
				prop.Value.ft = Value.ft;
				break;
			case PT_STRING8:
				prop.Value.lpszA = const_cast<LPSTR>(Value.lpszA.str.c_str());
				break;
			case PT_BINARY:
				prop.Value.bin.cb = Value.bin.cb;
				prop.Value.bin.lpb = const_cast<LPBYTE>(Value.bin.lpb.data());
				break;
			case PT_UNICODE:
				prop.Value.lpszW = const_cast<LPWSTR>(Value.lpszW.str.c_str());
				break;
			case PT_CLSID:
				guid = Value.lpguid.getData();
				prop.Value.lpguid = &guid;
				break;
			//case PT_MV_STRING8:
			//case PT_MV_UNICODE:
			//case PT_MV_BINARY:
			case PT_ERROR:
				prop.Value.err = Value.err;
				break;
			}

			return prop;
		}

		void EnsurePropStrings()
		{
			if (propStringsGenerated) return;
			auto sProp = getData();
			interpretprop::InterpretProp(&sProp, &propString, &altPropString);

			propString = strings::RemoveInvalidCharactersW(propString, false);
			altPropString = strings::RemoveInvalidCharactersW(altPropString, false);
			propStringsGenerated = true;
		}

		_Check_return_ std::wstring PropNum()
		{
			switch (PROP_TYPE(ulPropTag))
			{
			case PT_LONG:
				return smartview::InterpretNumberAsString(Value.l, ulPropTag, 0, nullptr, nullptr, false);
				break;
			case PT_I2:
				return smartview::InterpretNumberAsString(Value.i, ulPropTag, 0, nullptr, nullptr, false);
				break;
			case PT_I8:
				return smartview::InterpretNumberAsString(
					Value.li.getData().QuadPart, ulPropTag, 0, nullptr, nullptr, false);
				break;
			}

			return strings::emptystring;
		}

		// Any data we need to cache for getData can live here
	private:
		GUID guid{};
		std::wstring propString;
		std::wstring altPropString;
		bool propStringsGenerated{};
	};

	class PropertiesStruct : public SmartViewParser
	{
	public:
		void SetMaxEntries(DWORD maxEntries) { m_MaxEntries = maxEntries; }
		void EnableNickNameParsing() { m_NickName = true; }
		void EnableRuleConditionParsing() { m_RuleCondition = true; }
		_Check_return_ std::vector<SPropValueStruct> Props() const { return m_Props; } // TODO: add const

	private:
		void Parse() override;
		void ParseBlocks() override;

		bool m_NickName{};
		bool m_RuleCondition{};
		DWORD m_MaxEntries{_MaxEntriesSmall};
		std::vector<SPropValueStruct> m_Props;

		_Check_return_ SPropValueStruct BinToSPropValueStruct();
	};
}