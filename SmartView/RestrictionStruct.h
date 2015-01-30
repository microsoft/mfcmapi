#pragma once
#include "SmartViewParser.h"

// Caller allocates with new. Clean up with DeleteRestriction and delete[].
bool BinToRestriction(ULONG ulDepth, ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin, _Out_ size_t* lpcbBytesRead, _In_ LPSRestriction psrRestriction, bool bRuleCondition, bool bExtendedCount);
// Neuters an SRestriction - caller must use delete to delete the SRestriction
void DeleteRestriction(_In_ LPSRestriction lpRes);

class RestrictionStruct : public SmartViewParser
{
public:
	RestrictionStruct(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin);
	~RestrictionStruct();

private:
	void Parse();
	_Check_return_ wstring ToStringInternal();

	LPSRestriction m_lpRes;
};