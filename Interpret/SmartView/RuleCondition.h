#pragma once
#include <Interpret/SmartView/SmartViewParser.h>
#include <Interpret/SmartView/RestrictionStruct.h>

namespace smartview
{
	// http://msdn.microsoft.com/en-us/library/ee158295.aspx
	// http://msdn.microsoft.com/en-us/library/ee179073.aspx

	// [MS-OXCDATA]
	// PropertyName
	// =====================
	//   This structure specifies a Property Name
	//
	struct PropertyName
	{
		BYTE Kind{};
		GUID Guid{};
		DWORD LID{};
		BYTE NameSize{};
		std::wstring Name;
	};

	// [MS-OXORULE]
	// NamedPropertyInformation
	// =====================
	//   This structure specifies named property information for a rule condition
	//
	struct NamedPropertyInformation
	{
		WORD NoOfNamedProps{};
		std::vector<WORD> PropId;
		DWORD NamedPropertiesSize{};
		std::vector<PropertyName> PropertyName;
	};

	class RuleCondition : public SmartViewParser
	{
	public:
		RuleCondition();

		void Init(bool bExtended);

	private:
		void Parse() override;
		_Check_return_ std::wstring ToStringInternal() override;

		NamedPropertyInformation m_NamedPropertyInformation;
		RestrictionStruct m_lpRes;
		bool m_bExtended;
	};
}