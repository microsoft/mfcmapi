#pragma once
#include <core/smartview/SmartViewParser.h>
#include <core/smartview/TimeZone.h>
#include <core/smartview/block/blockStringW.h>
#include <core/smartview/block/blockBytes.h>
#include <core/smartview/block/blockT.h>

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
		std::shared_ptr<blockT<BYTE>> bMajorVersion = emptyT<BYTE>();
		std::shared_ptr<blockT<BYTE>> bMinorVersion = emptyT<BYTE>();
		std::shared_ptr<blockT<WORD>> wReserved = emptyT<WORD>();
		std::shared_ptr<blockT<WORD>> wTZRuleFlags = emptyT<WORD>();
		std::shared_ptr<blockT<WORD>> wYear = emptyT<WORD>();
		std::shared_ptr<blockBytes> X = emptyBB(); // 14 bytes
		std::shared_ptr<blockT<DWORD>> lBias = emptyT<DWORD>(); // offset from GMT
		std::shared_ptr<blockT<DWORD>> lStandardBias = emptyT<DWORD>(); // offset from bias during standard time
		std::shared_ptr<blockT<DWORD>> lDaylightBias = emptyT<DWORD>(); // offset from bias during daylight time
		SYSTEMTIMEBlock stStandardDate; // time to switch to standard time
		SYSTEMTIMEBlock stDaylightDate; // time to switch to daylight time

		TZRule(std::shared_ptr<binaryParser>& parser);
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

		std::shared_ptr<blockT<BYTE>> m_bMajorVersion = emptyT<BYTE>();
		std::shared_ptr<blockT<BYTE>> m_bMinorVersion = emptyT<BYTE>();
		std::shared_ptr<blockT<WORD>> m_cbHeader = emptyT<WORD>();
		std::shared_ptr<blockT<WORD>> m_wReserved = emptyT<WORD>();
		std::shared_ptr<blockT<WORD>> m_cchKeyName = emptyT<WORD>();
		std::shared_ptr<blockStringW> m_szKeyName = emptySW();
		std::shared_ptr<blockT<WORD>> m_cRules = emptyT<WORD>();
		std::vector<std::shared_ptr<TZRule>> m_lpTZRule;
	};
} // namespace smartview