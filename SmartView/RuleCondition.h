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
	RuleCondition(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin, bool bExtended);
	~RuleCondition();

private:
	void Parse();
	_Check_return_ wstring ToStringInternal();

	NamedPropertyInformationStruct m_NamedPropertyInformation;
	RestrictionStruct* m_lpRes;
	bool m_bExtended;
};