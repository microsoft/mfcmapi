#pragma once
#include <Interpret/SmartView/SmartViewParser.h>

namespace smartview
{
	struct TombstoneRecord
	{
		DWORD StartTime;
		DWORD EndTime;
		DWORD GlobalObjectIdSize;
		std::vector<BYTE> lpGlobalObjectId;
		WORD UsernameSize;
		std::string szUsername;
	};

	class TombStone : public SmartViewParser
	{
	public:
		TombStone();

	private:
		void Parse() override;
		_Check_return_ std::wstring ToStringInternal() override;

		DWORD m_Identifier;
		DWORD m_HeaderSize;
		DWORD m_Version;
		DWORD m_RecordsCount;
		DWORD m_ActualRecordsCount; // computed based on state, not read value
		DWORD m_RecordsSize;
		std::vector<TombstoneRecord> m_lpRecords;
	};
}