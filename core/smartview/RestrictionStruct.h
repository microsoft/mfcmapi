#pragma once
#include <core/smartview/SmartViewParser.h>
#include <core/smartview/PropertiesStruct.h>
#include <core/smartview/block/blockT.h>

namespace smartview
{
	class RestrictionStruct;

	struct SAndRestrictionStruct
	{
		blockT<DWORD> cRes;
		std::vector<std::shared_ptr<RestrictionStruct>> lpRes;
	};

	struct SOrRestrictionStruct
	{
		blockT<DWORD> cRes;
		std::vector<std::shared_ptr<RestrictionStruct>> lpRes;
	};

	struct SNotRestrictionStruct
	{
		blockT<DWORD> ulReserved;
		std::shared_ptr<RestrictionStruct> lpRes;
	};

	struct SContentRestrictionStruct
	{
		blockT<DWORD> ulFuzzyLevel;
		blockT<DWORD> ulPropTag;
		PropertiesStruct lpProp;
	};

	struct SBitMaskRestrictionStruct
	{
		blockT<DWORD> relBMR;
		blockT<DWORD> ulPropTag;
		blockT<DWORD> ulMask;
	};

	struct SPropertyRestrictionStruct
	{
		blockT<DWORD> relop;
		blockT<DWORD> ulPropTag;
		PropertiesStruct lpProp;
	};

	struct SComparePropsRestrictionStruct
	{
		blockT<DWORD> relop;
		blockT<DWORD> ulPropTag1;
		blockT<DWORD> ulPropTag2;
	};

	struct SSizeRestrictionStruct
	{
		blockT<DWORD> relop;
		blockT<DWORD> ulPropTag;
		blockT<DWORD> cb;
	};

	struct SExistRestrictionStruct
	{
		blockT<DWORD> ulReserved1;
		blockT<DWORD> ulPropTag;
		blockT<DWORD> ulReserved2;
	};

	struct SSubRestrictionStruct
	{
		blockT<DWORD> ulSubObject;
		std::shared_ptr<RestrictionStruct> lpRes;
	};

	struct SCommentRestrictionStruct
	{
		blockT<DWORD> cValues; /* # of properties in lpProp */
		std::shared_ptr<RestrictionStruct> lpRes;
		PropertiesStruct lpProp;
	};

	struct SAnnotationRestrictionStruct
	{
		blockT<DWORD> cValues; /* # of properties in lpProp */
		std::shared_ptr<RestrictionStruct> lpRes;
		PropertiesStruct lpProp;
	};

	struct SCountRestrictionStruct
	{
		blockT<DWORD> ulCount;
		std::shared_ptr<RestrictionStruct> lpRes;
	};

	class RestrictionStruct : public SmartViewParser
	{
	public:
		RestrictionStruct(bool bRuleCondition, bool bExtendedCount)
			: m_bRuleCondition(bRuleCondition), m_bExtendedCount(bExtendedCount)
		{
		}
		RestrictionStruct(std::shared_ptr<binaryParser> parser, ULONG ulDepth, bool bRuleCondition, bool bExtendedCount)
			: m_bRuleCondition(bRuleCondition), m_bExtendedCount(bExtendedCount)
		{
			m_Parser = parser;
			parse(ulDepth);
		}

	private:
		void Parse() override { parse(0); }
		void parse(ULONG ulDepth);
		void ParseBlocks() override
		{
			setRoot(L"Restriction:\r\n");
			parseBlocks(0);
		};
		void parseBlocks(ULONG ulTabLevel);

		std::shared_ptr<blockT<DWORD>> rt = emptyT<DWORD>(); /* Restriction type */
		SComparePropsRestrictionStruct resCompareProps;
		SAndRestrictionStruct resAnd;
		SOrRestrictionStruct resOr;
		SNotRestrictionStruct resNot;
		SContentRestrictionStruct resContent;
		SPropertyRestrictionStruct resProperty;
		SBitMaskRestrictionStruct resBitMask;
		SSizeRestrictionStruct resSize;
		SExistRestrictionStruct resExist;
		SSubRestrictionStruct resSub;
		SCommentRestrictionStruct resComment;
		SAnnotationRestrictionStruct resAnnotation;
		SCountRestrictionStruct resCount;

		bool m_bRuleCondition{};
		bool m_bExtendedCount{};
	};
} // namespace smartview