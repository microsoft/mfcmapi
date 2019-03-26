#pragma once
#include <core/smartview/SmartViewParser.h>
#include <core/smartview/RestrictionStruct.h>
#include <core/smartview/block/blockStringW.h>
#include <core/smartview/block/blockT.h>

namespace smartview
{
	// [MS-OXORULE] 2.2.1.3.1.9 PidTagRuleCondition Property
	// https://msdn.microsoft.com/en-us/library/ee204420(v=exchg.80).aspx

	// [MS-OXORULE] 2.2.4.1.10 PidTagExtendedRuleMessageCondition Property
	// https://msdn.microsoft.com/en-us/library/ee200930(v=exchg.80).aspx

	// RuleRestriction
	// https://msdn.microsoft.com/en-us/library/ee201126(v=exchg.80).aspx

	// [MS-OXCDATA] 2.6.1 PropertyName Structure
	// http://msdn.microsoft.com/en-us/library/ee158295.aspx
	//   This structure specifies a Property Name
	//
	struct PropertyName
	{
		std::shared_ptr<blockT<BYTE>> Kind = emptyT<BYTE>();
		std::shared_ptr<blockT<GUID>> Guid = emptyT<GUID>();
		std::shared_ptr<blockT<DWORD>> LID = emptyT<DWORD>();
		std::shared_ptr<blockT<BYTE>> NameSize = emptyT<BYTE>();
		std::shared_ptr<blockStringW> Name = emptySW();

		PropertyName(const std::shared_ptr<binaryParser>& parser);
	};

	// [MS-OXORULE] 2.2.4.2 NamedPropertyInformation Structure
	// https://msdn.microsoft.com/en-us/library/ee159014(v=exchg.80).aspx
	// =====================
	//   This structure specifies named property information for a rule condition
	//
	struct NamedPropertyInformation
	{
		std::shared_ptr<blockT<WORD>> NoOfNamedProps = emptyT<WORD>();
		std::vector<std::shared_ptr<blockT<WORD>>> PropId;
		std::shared_ptr<blockT<DWORD>> NamedPropertiesSize = emptyT<DWORD>();
		std::vector<std::shared_ptr<PropertyName>> PropertyName;
	};

	class RuleCondition : public SmartViewParser
	{
	public:
		void Init(bool bExtended);

	private:
		void Parse() override;
		void ParseBlocks() override;

		NamedPropertyInformation m_NamedPropertyInformation;
		std::shared_ptr<RestrictionStruct> m_lpRes;
		bool m_bExtended{};
	};
} // namespace smartview