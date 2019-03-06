#pragma once
#include <core/smartview/SmartViewParser.h>
#include <core/smartview/RestrictionStruct.h>

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
		blockT<BYTE> Kind{};
		blockT<GUID> Guid{};
		blockT<DWORD> LID{};
		blockT<BYTE> NameSize{};
		blockStringW Name;
	};

	// [MS-OXORULE] 2.2.4.2 NamedPropertyInformation Structure
	// https://msdn.microsoft.com/en-us/library/ee159014(v=exchg.80).aspx
	// =====================
	//   This structure specifies named property information for a rule condition
	//
	struct NamedPropertyInformation
	{
		blockT<WORD> NoOfNamedProps{};
		std::vector<blockT<WORD>> PropId;
		blockT<DWORD> NamedPropertiesSize{};
		std::vector<PropertyName> PropertyName;
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