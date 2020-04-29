#pragma once
#include <core/smartview/smartViewParser.h>
#include <core/smartview/PropertiesStruct.h>
#include <core/smartview/block/blockT.h>

namespace smartview
{
	class RestrictionStruct;

	// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxcdata/3d0e5e14-7464-4659-b83d-a601b81fbdf5
	// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxocfg/001660fe-9839-41b9-9d5f-cd2b1a577e3b
	struct SAndRestrictionStruct
	{
		std::shared_ptr<blockT<DWORD>> cRes = emptyT<DWORD>();
		std::vector<std::shared_ptr<RestrictionStruct>> lpRes;
	};

	// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxcdata/50a969db-a794-4e3c-9fc0-28f37bc1d7a2
	// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxocfg/09499990-ab00-4d4f-9c0d-cff61d9cddff
	struct SOrRestrictionStruct
	{
		std::shared_ptr<blockT<DWORD>> cRes = emptyT<DWORD>();
		std::vector<std::shared_ptr<RestrictionStruct>> lpRes;
	};

	// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxcdata/0abc5c41-9db7-4e6c-8d4d-b5c7e51d5355
	// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxocfg/f10fbc18-b384-4cd8-9490-a6955035e2ec
	struct SNotRestrictionStruct
	{
		std::shared_ptr<RestrictionStruct> lpRes;
	};

	// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxcdata/c1526deb-d05d-4d42-af68-d0233e4cd064
	// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxocfg/e6ded216-f49e-49b9-8b26-e1743e76c897
	struct SContentRestrictionStruct
	{
		std::shared_ptr<blockT<DWORD>> ulFuzzyLevel = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> ulPropTag = emptyT<DWORD>();
		PropertiesStruct lpProp;
	};

	// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxcdata/0c7fe9fe-b16a-49d1-ba9b-053524b5da97
	// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxocfg/46c50b53-f6b1-4d2d-ae79-ced052b261d1
	struct SBitMaskRestrictionStruct
	{
		std::shared_ptr<blockT<DWORD>> relBMR = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> ulPropTag = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> ulMask = emptyT<DWORD>();
	};

	// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxcdata/285d931b-e1ef-4390-aa84-9a3ee98b6e28
	// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxocfg/4e66da53-032b-46f0-8b84-39ee1e96de80
	struct SPropertyRestrictionStruct
	{
		std::shared_ptr<blockT<DWORD>> relop = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> ulPropTag = emptyT<DWORD>();
		PropertiesStruct lpProp;
	};

	// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxcdata/486b4b66-a6f0-43d3-98ee-91f1e0e29907
	// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxocfg/e0dc196d-e041-4d80-bc44-89b94fd71158
	struct SComparePropsRestrictionStruct
	{
		std::shared_ptr<blockT<DWORD>> relop = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> ulPropTag1 = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> ulPropTag2 = emptyT<DWORD>();
	};

	// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxcdata/b81c6f60-0e17-43e0-8d1f-c3015c7cd9de
	// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxocfg/c6915f61-debe-466c-9145-ffadee0cb28e
	struct SSizeRestrictionStruct
	{
		std::shared_ptr<blockT<DWORD>> relop = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> ulPropTag = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> cb = emptyT<DWORD>();
	};

	// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxcdata/c4da9872-8903-4b70-9b23-9980f6bc21c7
	// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxocfg/116661e8-6c17-417c-a75d-261ac0ac29ba
	struct SExistRestrictionStruct
	{
		std::shared_ptr<blockT<DWORD>> ulPropTag = emptyT<DWORD>();
	};

	// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxcdata/9e87191a-fe53-48a1-aba4-55ba1f8c221f
	// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxocfg/550dc0ac-01cf-4ee2-b15c-799faa852120
	struct SSubRestrictionStruct
	{
		std::shared_ptr<blockT<DWORD>> ulSubObject = emptyT<DWORD>();
		std::shared_ptr<RestrictionStruct> lpRes;
	};

	// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxcdata/ceb8bd59-90e5-4c6f-8632-8fc9ef639766
	// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxocfg/2ffa8c7a-38c1-4446-8fe1-13b52854460e
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

	// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxcdata/66b892da-032a-4f66-812a-01c11457d5ce
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
		RestrictionStruct(
			const std::shared_ptr<binaryParser>& parser,
			ULONG ulDepth,
			bool bRuleCondition,
			bool bExtendedCount)
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