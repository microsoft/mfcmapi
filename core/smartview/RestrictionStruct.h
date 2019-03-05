#pragma once
#include <core/smartview/SmartViewParser.h>
#include <core/smartview/PropertiesStruct.h>

namespace smartview
{
	struct SRestrictionStruct;

	struct SAndRestrictionStruct
	{
		blockT<DWORD> cRes;
		std::vector<SRestrictionStruct> lpRes;
	};

	struct SOrRestrictionStruct
	{
		blockT<DWORD> cRes;
		std::vector<SRestrictionStruct> lpRes;
	};

	struct SNotRestrictionStruct
	{
		blockT<DWORD> ulReserved;
		std::vector<SRestrictionStruct> lpRes;
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
		std::vector<SRestrictionStruct> lpRes;
	};

	struct SCommentRestrictionStruct
	{
		blockT<DWORD> cValues; /* # of properties in lpProp */
		std::vector<SRestrictionStruct> lpRes;
		PropertiesStruct lpProp;
	};

	struct SAnnotationRestrictionStruct
	{
		blockT<DWORD> cValues; /* # of properties in lpProp */
		std::vector<SRestrictionStruct> lpRes;
		PropertiesStruct lpProp;
	};

	struct SCountRestrictionStruct
	{
		blockT<DWORD> ulCount;
		std::vector<SRestrictionStruct> lpRes;
	};

	struct SRestrictionStruct
	{
		blockT<DWORD> rt; /* Restriction type */
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

		SRestrictionStruct() = default;
		SRestrictionStruct(
			std::shared_ptr<binaryParser> parser,
			ULONG ulDepth,
			bool bRuleCondition,
			bool bExtendedCount)
		{
			init(parser, ulDepth, bRuleCondition, bExtendedCount);
		}

		void init(std::shared_ptr<binaryParser> parser, ULONG ulDepth, bool bRuleCondition, bool bExtendedCount);
	};

	class RestrictionStruct : public SmartViewParser
	{
	public:
		void init(bool bRuleCondition, bool bExtended);

	private:
		void Parse() override;
		void ParseBlocks() override;

		void ParseRestriction(_In_ const SRestrictionStruct& lpRes, ULONG ulTabLevel);

		bool m_bRuleCondition{};
		bool m_bExtended{};
		SRestrictionStruct m_lpRes;
	};
} // namespace smartview