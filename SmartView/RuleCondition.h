#pragma once
#include "SmartViewParser.h"
#include "RestrictionStruct.h"

// http://msdn.microsoft.com/en-us/library/ee158295.aspx
// http://msdn.microsoft.com/en-us/library/ee179073.aspx

// [MS-OXCDATA]
// PropertyNameStruct
// =====================
//   This structure specifies a Property Name
//
struct PropertyNameStruct
{
	BYTE Kind;
	GUID Guid;
	DWORD LID;
	BYTE NameSize;
	LPWSTR Name;
};

// [MS-OXORULE]
// NamedPropertyInformationStruct
// =====================
//   This structure specifies named property information for a rule condition
//
struct NamedPropertyInformationStruct
{
	WORD NoOfNamedProps;
	WORD* PropId;
	DWORD NamedPropertiesSize;
	PropertyNameStruct* PropertyName;
};

class RuleCondition : public SmartViewParser
{
public:
	RuleCondition(bool bExtended);
	~RuleCondition();

private:
	void Parse() override;
	_Check_return_ wstring ToStringInternal() override;

	NamedPropertyInformationStruct m_NamedPropertyInformation;
	RestrictionStruct* m_lpRes;
	bool m_bExtended;
};