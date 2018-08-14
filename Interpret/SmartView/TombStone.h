#pragma once
#include <Interpret/SmartView/SmartViewParser.h>

namespace smartview
{
	struct TombstoneRecord
	{
		blockT<DWORD> StartTime;
		blockT<DWORD> EndTime;
		blockT<DWORD> GlobalObjectIdSize;
		blockBytes lpGlobalObjectId;
		blockT<WORD> UsernameSize;
		blockStringA szUsername;
	};

	class TombStone : public SmartViewParser
	{
	public:
		TombStone();

	private:
		void Parse() override;
		_Check_return_ std::wstring ToStringInternal() override;

		blockT<DWORD> m_Identifier;
		blockT<DWORD> m_HeaderSize;
		blockT<DWORD> m_Version;
		blockT<DWORD> m_RecordsCount;
		DWORD m_ActualRecordsCount{}; // computed based on state, not read value
		blockT<DWORD> m_RecordsSize;
		std::vector<TombstoneRecord> m_lpRecords;
	};
}