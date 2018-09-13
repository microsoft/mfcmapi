#pragma once
#include <Interpret/SmartView/SmartViewParser.h>
#include <Interpret/SmartView/TimeZone.h>

namespace smartview
{
	// TZRule
	// =====================
	//   This structure represents both a description when a daylight.
	//   savings shift occurs, and in addition, the year in which that
	//   timezone rule came into effect.
	//
	struct TZRule
	{
		blockT<BYTE> bMajorVersion;
		blockT<BYTE> bMinorVersion;
		blockT<WORD> wReserved;
		blockT<WORD> wTZRuleFlags;
		blockT<WORD> wYear;
		blockBytes X; // 14 bytes
		blockT<DWORD> lBias; // offset from GMT
		blockT<DWORD> lStandardBias; // offset from bias during standard time
		blockT<DWORD> lDaylightBias; // offset from bias during daylight time
		SYSTEMTIMEBlock stStandardDate; // time to switch to standard time
		SYSTEMTIMEBlock stDaylightDate; // time to switch to daylight time
	};

	// TimeZoneDefinition
	// =====================
	//   This represents an entire timezone including all historical, current
	//   and future timezone shift rules for daylight savings time, etc.
	class TimeZoneDefinition : public SmartViewParser
	{
	private:
		void Parse() override;
		void ParseBlocks() override;

		blockT<BYTE> m_bMajorVersion;
		blockT<BYTE> m_bMinorVersion;
		blockT<WORD> m_cbHeader;
		blockT<WORD> m_wReserved;
		blockT<WORD> m_cchKeyName;
		blockStringW m_szKeyName;
		blockT<WORD> m_cRules;
		std::vector<TZRule> m_lpTZRule;
	};
} // namespace smartview