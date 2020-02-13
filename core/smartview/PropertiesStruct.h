#pragma once
#include <core/smartview/smartViewParser.h>
#include <core/smartview/SmartView.h>
#include <core/property/parseProperty.h>
#include <core/smartview/block/blockStringA.h>
#include <core/smartview/block/blockStringW.h>
#include <core/smartview/block/blockBytes.h>
#include <core/smartview/block/blockT.h>

namespace smartview
{
	struct FILETIMEBLock
	{
		std::shared_ptr<blockT<DWORD>> dwLowDateTime = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> dwHighDateTime = emptyT<DWORD>();
		operator FILETIME() const { return FILETIME{*dwLowDateTime, *dwHighDateTime}; }
		size_t getSize() const { return dwLowDateTime->getSize() + dwHighDateTime->getSize(); }
		size_t getOffset() const { return dwHighDateTime->getOffset(); }
	};

	struct SBinaryBlock
	{
		std::shared_ptr<blockT<ULONG>> cb = emptyT<ULONG>();
		std::shared_ptr<blockBytes> lpb = emptyBB();
		size_t getSize() const { return cb->getSize() + lpb->getSize(); }
		size_t getOffset() const { return cb->getOffset() ? cb->getOffset() : lpb->getOffset(); }

		SBinaryBlock(const std::shared_ptr<binaryParser>& parser);
		SBinaryBlock() noexcept {};
	};

	struct SBinaryArrayBlock
	{
		std::shared_ptr<blockT<ULONG>> cValues = emptyT<ULONG>();
		std::vector<std::shared_ptr<SBinaryBlock>> lpbin;
	};

	struct CountedStringA
	{
		std::shared_ptr<blockT<DWORD>> cb = emptyT<DWORD>();
		std::shared_ptr<blockStringA> str = emptySA();
		size_t getSize() const { return cb->getSize() + str->getSize(); }
		size_t getOffset() const { return cb->getOffset(); }
	};

	struct CountedStringW
	{
		std::shared_ptr<blockT<DWORD>> cb = emptyT<DWORD>();
		std::shared_ptr<blockStringW> str = emptySW();
		size_t getSize() const { return cb->getSize() + str->getSize(); }
		size_t getOffset() const { return cb->getOffset(); }
	};

	struct StringArrayA
	{
		std::shared_ptr<blockT<ULONG>> cValues = emptyT<ULONG>();
		std::vector<std::shared_ptr<blockStringA>> lppszA;
	};

	struct StringArrayW
	{
		std::shared_ptr<blockT<ULONG>> cValues = emptyT<ULONG>();
		std::vector<std::shared_ptr<blockStringW>> lppszW;
	};

	struct PVBlock
	{
		std::shared_ptr<blockT<WORD>> i = emptyT<WORD>(); /* case PT_I2 */
		std::shared_ptr<blockT<LONG>> l = emptyT<LONG>(); /* case PT_LONG */
		std::shared_ptr<blockT<WORD>> b = emptyT<WORD>(); /* case PT_BOOLEAN */
		std::shared_ptr<blockT<float>> flt = emptyT<float>(); /* case PT_R4 */
		std::shared_ptr<blockT<double>> dbl = emptyT<double>(); /* case PT_DOUBLE */
		FILETIMEBLock ft; /* case PT_SYSTIME */
		CountedStringA lpszA; /* case PT_STRING8 */
		SBinaryBlock bin; /* case PT_BINARY */
		CountedStringW lpszW; /* case PT_UNICODE */
		std::shared_ptr<blockT<GUID>> lpguid = emptyT<GUID>(); /* case PT_CLSID */
		std::shared_ptr<blockT<LARGE_INTEGER>> li = emptyT<LARGE_INTEGER>(); /* case PT_I8 */
		SBinaryArrayBlock MVbin; /* case PT_MV_BINARY */
		StringArrayA MVszA; /* case PT_MV_STRING8 */
		StringArrayW MVszW; /* case PT_MV_UNICODE */
		std::shared_ptr<blockT<SCODE>> err = emptyT<SCODE>(); /* case PT_ERROR */
	};

	struct SPropValueStruct
	{
		std::shared_ptr<blockT<WORD>> PropType = emptyT<WORD>();
		std::shared_ptr<blockT<WORD>> PropID = emptyT<WORD>();
		std::shared_ptr<blockT<ULONG>> ulPropTag = emptyT<ULONG>();
		ULONG dwAlignPad{};
		PVBlock Value;
		_Check_return_ std::shared_ptr<blockStringW> PropBlock()
		{
			EnsurePropBlocks();
			return propBlock;
		}
		_Check_return_ std::shared_ptr<blockStringW> AltPropBlock()
		{
			EnsurePropBlocks();
			return altPropBlock;
		}
		_Check_return_ std::shared_ptr<blockStringW> SmartViewBlock()
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
			prop.ulPropTag = *ulPropTag;
			prop.dwAlignPad = dwAlignPad;
			switch (*PropType)
			{
			case PT_I2:
				prop.Value.i = *Value.i;
				size = Value.i->getSize();
				offset = Value.i->getOffset();
				break;
			case PT_LONG:
				prop.Value.l = *Value.l;
				size = Value.l->getSize();
				offset = Value.l->getOffset();
				break;
			case PT_R4:
				prop.Value.flt = *Value.flt;
				size = Value.flt->getSize();
				offset = Value.flt->getOffset();
				break;
			case PT_DOUBLE:
				prop.Value.dbl = *Value.dbl;
				size = Value.dbl->getSize();
				offset = Value.dbl->getOffset();
				break;
			case PT_BOOLEAN:
				prop.Value.b = *Value.b;
				size = Value.b->getSize();
				offset = Value.b->getOffset();
				break;
			case PT_I8:
				prop.Value.li = Value.li->getData();
				size = Value.li->getSize();
				offset = Value.li->getOffset();
				break;
			case PT_SYSTIME:
				prop.Value.ft = Value.ft;
				size = Value.ft.getSize();
				offset = Value.ft.getOffset();
				break;
			case PT_STRING8:
				prop.Value.lpszA = const_cast<LPSTR>(Value.lpszA.str->c_str());
				size = Value.lpszA.getSize();
				offset = Value.lpszA.getOffset();
				break;
			case PT_BINARY:
				prop.Value.bin.cb = *Value.bin.cb;
				prop.Value.bin.lpb = const_cast<LPBYTE>(Value.bin.lpb->data());
				size = Value.bin.getSize();
				offset = Value.bin.getOffset();
				break;
			case PT_UNICODE:
				prop.Value.lpszW = const_cast<LPWSTR>(Value.lpszW.str->c_str());
				size = Value.lpszW.getSize();
				offset = Value.lpszW.getOffset();
				break;
			case PT_CLSID:
				guid = Value.lpguid->getData();
				prop.Value.lpguid = &guid;
				size = Value.lpguid->getSize();
				offset = Value.lpguid->getOffset();
				break;
			//case PT_MV_STRING8:
			//case PT_MV_UNICODE:
			//case PT_MV_BINARY:
			case PT_ERROR:
				prop.Value.err = *Value.err;
				size = Value.err->getSize();
				offset = Value.err->getOffset();
				break;
			}

			auto propString = std::wstring{};
			auto altPropString = std::wstring{};
			property::parseProperty(&prop, &propString, &altPropString);

			propBlock =
				std::make_shared<blockStringW>(strings::RemoveInvalidCharactersW(propString, false), size, offset);

			altPropBlock =
				std::make_shared<blockStringW>(strings::RemoveInvalidCharactersW(altPropString, false), size, offset);

			const auto smartViewString = parsePropertySmartView(&prop, nullptr, nullptr, nullptr, false, false);
			smartViewBlock = std::make_shared<blockStringW>(smartViewString, size, offset);

			propStringsGenerated = true;
		}

		_Check_return_ std::wstring PropNum() const
		{
			switch (PROP_TYPE(*ulPropTag))
			{
			case PT_LONG:
				return InterpretNumberAsString(*Value.l, *ulPropTag, 0, nullptr, nullptr, false);
			case PT_I2:
				return InterpretNumberAsString(*Value.i, *ulPropTag, 0, nullptr, nullptr, false);
			case PT_I8:
				return InterpretNumberAsString(Value.li->getData().QuadPart, *ulPropTag, 0, nullptr, nullptr, false);
			}

			return strings::emptystring;
		}

		SPropValueStruct(const std::shared_ptr<binaryParser>& parser, bool doNickname, bool doRuleProcessing);

		// Any data we need to cache for getData can live here
	private:
		GUID guid{};
		std::shared_ptr<blockStringW> propBlock = emptySW();
		std::shared_ptr<blockStringW> altPropBlock = emptySW();
		std::shared_ptr<blockStringW> smartViewBlock = emptySW();
		bool propStringsGenerated{};
	};

	class PropertiesStruct : public smartViewParser
	{
	public:
		void parse(const std::shared_ptr<binaryParser>& parser, DWORD cValues, bool bRuleCondition);
		void SetMaxEntries(DWORD maxEntries) noexcept { m_MaxEntries = maxEntries; }
		void EnableNickNameParsing() noexcept { m_NickName = true; }
		void EnableRuleConditionParsing() noexcept { m_RuleCondition = true; }
		_Check_return_ std::vector<std::shared_ptr<SPropValueStruct>>& Props() noexcept { return m_Props; }

	private:
		void parse() override;
		void parseBlocks() override;

		bool m_NickName{};
		bool m_RuleCondition{};
		DWORD m_MaxEntries{_MaxEntriesSmall};
		std::vector<std::shared_ptr<SPropValueStruct>> m_Props;
	};
} // namespace smartview