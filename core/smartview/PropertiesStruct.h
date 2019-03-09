#pragma once
#include <core/smartview/SmartViewParser.h>
#include <core/smartview/SmartView.h>
#include <core/property/parseProperty.h>

namespace smartview
{
	struct FILETIMEBLock
	{
		blockT<DWORD> dwLowDateTime;
		blockT<DWORD> dwHighDateTime;
		operator FILETIME() const { return FILETIME{dwLowDateTime, dwHighDateTime}; }
		size_t getSize() const { return dwLowDateTime.getSize() + dwHighDateTime.getSize(); }
		size_t getOffset() const { return dwLowDateTime.getOffset(); }
	};

	struct SBinaryBlock
	{
		blockT<ULONG> cb;
		blockBytes lpb;
		size_t getSize() const { return cb.getSize() + lpb.getSize(); }
		size_t getOffset() const { return cb.getOffset() ? cb.getOffset() : lpb.getOffset(); }
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
		size_t getSize() const { return cb.getSize() + str.getSize(); }
		size_t getOffset() const { return cb.getOffset(); }
	};

	struct CountedStringW
	{
		blockT<DWORD> cb;
		blockStringW str;
		size_t getSize() const { return cb.getSize() + str.getSize(); }
		size_t getOffset() const { return cb.getOffset(); }
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
		_Check_return_ blockStringW& PropBlock()
		{
			EnsurePropBlocks();
			return propBlock;
		}
		_Check_return_ blockStringW& AltPropBlock()
		{
			EnsurePropBlocks();
			return altPropBlock;
		}
		_Check_return_ blockStringW& SmartViewBlock()
		{
			EnsurePropBlocks();
			return smartViewBlock;
		}

		// TODO: Fill in missing cases with test coverage
		void EnsurePropBlocks()
		{
			if (propStringsGenerated) return;
			auto prop = SPropValue{};
			auto size = size_t{};
			auto offset = size_t{};
			prop.ulPropTag = ulPropTag;
			prop.dwAlignPad = dwAlignPad;
			switch (PropType)
			{
			case PT_I2:
				prop.Value.i = Value.i;
				size = Value.i.getSize();
				offset = Value.i.getOffset();
				break;
			case PT_LONG:
				prop.Value.l = Value.l;
				size = Value.l.getSize();
				offset = Value.l.getOffset();
				break;
			case PT_R4:
				prop.Value.flt = Value.flt;
				size = Value.flt.getSize();
				offset = Value.flt.getOffset();
				break;
			case PT_DOUBLE:
				prop.Value.dbl = Value.dbl;
				size = Value.dbl.getSize();
				offset = Value.dbl.getOffset();
				break;
			case PT_BOOLEAN:
				prop.Value.b = Value.b;
				size = Value.b.getSize();
				offset = Value.b.getOffset();
				break;
			case PT_I8:
				prop.Value.li = Value.li.getData();
				size = Value.li.getSize();
				offset = Value.li.getOffset();
				break;
			case PT_SYSTIME:
				prop.Value.ft = Value.ft;
				size = Value.ft.getSize();
				offset = Value.ft.getOffset();
				break;
			case PT_STRING8:
				prop.Value.lpszA = const_cast<LPSTR>(Value.lpszA.str.c_str());
				size = Value.lpszA.getSize();
				offset = Value.lpszA.getOffset();
				break;
			case PT_BINARY:
				prop.Value.bin.cb = Value.bin.cb;
				prop.Value.bin.lpb = const_cast<LPBYTE>(Value.bin.lpb.data());
				size = Value.bin.getSize();
				offset = Value.bin.getOffset();
				break;
			case PT_UNICODE:
				prop.Value.lpszW = const_cast<LPWSTR>(Value.lpszW.str.c_str());
				size = Value.lpszW.getSize();
				offset = Value.lpszW.getOffset();
				break;
			case PT_CLSID:
				guid = Value.lpguid.getData();
				prop.Value.lpguid = &guid;
				size = Value.lpguid.getSize();
				offset = Value.lpguid.getOffset();
				break;
			//case PT_MV_STRING8:
			//case PT_MV_UNICODE:
			//case PT_MV_BINARY:
			case PT_ERROR:
				prop.Value.err = Value.err;
				size = Value.err.getSize();
				offset = Value.err.getOffset();
				break;
			}

			auto propString = std::wstring{};
			auto altPropString = std::wstring{};
			property::parseProperty(&prop, &propString, &altPropString);

			propBlock.setData(strings::RemoveInvalidCharactersW(propString, false));
			propBlock.setSize(size);
			propBlock.setOffset(offset);

			altPropBlock.setData(strings::RemoveInvalidCharactersW(altPropString, false));
			altPropBlock.setSize(size);
			altPropBlock.setOffset(offset);

			const auto smartViewString = parsePropertySmartView(&prop, nullptr, nullptr, nullptr, false, false);
			smartViewBlock.setData(smartViewString);
			smartViewBlock.setSize(size);
			smartViewBlock.setOffset(offset);

			propStringsGenerated = true;
		}

		_Check_return_ std::wstring PropNum() const
		{
			switch (PROP_TYPE(ulPropTag))
			{
			case PT_LONG:
				return InterpretNumberAsString(Value.l, ulPropTag, 0, nullptr, nullptr, false);
			case PT_I2:
				return InterpretNumberAsString(Value.i, ulPropTag, 0, nullptr, nullptr, false);
			case PT_I8:
				return InterpretNumberAsString(Value.li.getData().QuadPart, ulPropTag, 0, nullptr, nullptr, false);
			}

			return strings::emptystring;
		}

		// Any data we need to cache for getData can live here
	private:
		GUID guid{};
		blockStringW propBlock;
		blockStringW altPropBlock;
		blockStringW smartViewBlock;
		bool propStringsGenerated{};
	};

	class PropertiesStruct : public SmartViewParser
	{
	public:
		void init(std::shared_ptr<binaryParser> parser, DWORD cValues, bool bRuleCondition);
		void SetMaxEntries(DWORD maxEntries) { m_MaxEntries = maxEntries; }
		void EnableNickNameParsing() { m_NickName = true; }
		void EnableRuleConditionParsing() { m_RuleCondition = true; }
		_Check_return_ std::vector<SPropValueStruct>& Props() { return m_Props; }

	private:
		void Parse() override;
		void ParseBlocks() override;

		bool m_NickName{};
		bool m_RuleCondition{};
		DWORD m_MaxEntries{_MaxEntriesSmall};
		std::vector<SPropValueStruct> m_Props;

		_Check_return_ SPropValueStruct BinToSPropValueStruct();
	};
} // namespace smartview