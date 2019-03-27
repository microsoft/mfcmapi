#pragma once
#include <core/smartview/smartViewParser.h>
#include <core/smartview/PropertiesStruct.h>
#include <core/smartview/block/blockT.h>

namespace smartview
{
	class RestrictionStruct;

	struct SAndRestrictionStruct
	{
		std::shared_ptr<blockT<DWORD>> cRes = emptyT<DWORD>();
		std::vector<std::shared_ptr<RestrictionStruct>> lpRes;
	};

	struct SOrRestrictionStruct
	{
		std::shared_ptr<blockT<DWORD>> cRes = emptyT<DWORD>();
		std::vector<std::shared_ptr<RestrictionStruct>> lpRes;
	};

	struct SNotRestrictionStruct
	{
		std::shared_ptr<blockT<DWORD>> ulReserved = emptyT<DWORD>();
		std::shared_ptr<RestrictionStruct> lpRes;
	};

	struct SContentRestrictionStruct
	{
		std::shared_ptr<blockT<DWORD>> ulFuzzyLevel = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> ulPropTag = emptyT<DWORD>();
		PropertiesStruct lpProp;
	};

	struct SBitMaskRestrictionStruct
	{
		std::shared_ptr<blockT<DWORD>> relBMR = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> ulPropTag = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> ulMask = emptyT<DWORD>();
	};

	struct SPropertyRestrictionStruct
	{
		std::shared_ptr<blockT<DWORD>> relop = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> ulPropTag = emptyT<DWORD>();
		PropertiesStruct lpProp;
	};

	struct SComparePropsRestrictionStruct
	{
		std::shared_ptr<blockT<DWORD>> relop = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> ulPropTag1 = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> ulPropTag2 = emptyT<DWORD>();
	};

	struct SSizeRestrictionStruct
	{
		std::shared_ptr<blockT<DWORD>> relop = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> ulPropTag = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> cb = emptyT<DWORD>();
	};

	struct SExistRestrictionStruct
	{
		std::shared_ptr<blockT<DWORD>> ulReserved1 = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> ulPropTag = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> ulReserved2 = emptyT<DWORD>();
	};

	struct SSubRestrictionStruct
	{
		std::shared_ptr<blockT<DWORD>> ulSubObject = emptyT<DWORD>();
		std::shared_ptr<RestrictionStruct> lpRes;
	};

	struct SCommentRestrictionStruct
	{
		std::shared_ptr<blockT<DWORD>> cValues = emptyT<DWORD>(); /* # of properties in lpProp */
		std::shared_ptr<RestrictionStruct> lpRes;
		PropertiesStruct lpProp;
	};

	struct SAnnotationRestrictionStruct
	{
		std::shared_ptr<blockT<DWORD>> cValues = emptyT<DWORD>(); /* # of properties in lpProp */
		std::shared_ptr<RestrictionStruct> lpRes;
		PropertiesStruct lpProp;
	};

	struct SCountRestrictionStruct
	{
		std::shared_ptr<blockT<DWORD>> ulCount = emptyT<DWORD>();
		std::shared_ptr<RestrictionStruct> lpRes;
	};

	class RestrictionStruct : public smartViewParser
	{
	public:
		RestrictionStruct(bool bRuleCondition, bool bExtendedCount)
			: m_bRuleCondition(bRuleCondition), m_bExtendedCount(bExtendedCount)
		{
		}
		RestrictionStruct(const std::shared_ptr<binaryParser>& parser, ULONG ulDepth, bool bRuleCondition, bool bExtendedCount)
			: m_bRuleCondition(bRuleCondition), m_bExtendedCount(bExtendedCount)
		{
			m_Parser = parser;
			parse(ulDepth);
		}

	private:
		void parse() override { parse(0); }
		void parse(ULONG ulDepth);
		void parseBlocks() override
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