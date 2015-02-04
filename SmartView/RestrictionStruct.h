#pragma once
#include "SmartViewParser.h"

class RestrictionStruct : public SmartViewParser
{
public:
	RestrictionStruct(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin, bool bRuleCondition, bool bExtended);
	~RestrictionStruct();

private:
	void Parse();
	_Check_return_ wstring ToStringInternal();

	// Caller allocates with new. Clean up with DeleteRestriction and delete[].
	bool BinToRestriction(ULONG ulDepth, _In_ LPSRestriction psrRestriction, bool bRuleCondition, bool bExtendedCount);
	// Neuters an SRestriction - caller must use delete to delete the SRestriction
	void DeleteRestriction(_In_ LPSRestriction lpRes);

	bool m_bRuleCondition;
	bool m_bExtended;
	LPSRestriction m_lpRes;
};