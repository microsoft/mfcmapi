#pragma once
#include "SmartViewParser.h"

class RestrictionStruct : public SmartViewParser
{
public:
	RestrictionStruct();

	void Init(bool bRuleCondition, bool bExtended);
private:
	void Parse() override;
	_Check_return_ std::wstring ToStringInternal() override;

	bool BinToRestriction(ULONG ulDepth, _In_ LPSRestriction psrRestriction, bool bRuleCondition, bool bExtendedCount);

	bool m_bRuleCondition;
	bool m_bExtended;
	LPSRestriction m_lpRes;
};