#pragma once
#include "SmartViewParser.h"

// [MS-OXOMSG].pdf
struct ResponseLevel
{
	bool DeltaCode;
	DWORD TimeDelta;
	BYTE Random;
	BYTE Level;
};

class ConversationIndex : public SmartViewParser
{
public:
	ConversationIndex();

private:
	void Parse() override;
	_Check_return_ wstring ToStringInternal() override;

	BYTE m_UnnamedByte;
	FILETIME m_ftCurrent;
	GUID m_guid;
	ULONG m_ulResponseLevels;
	vector<ResponseLevel> m_lpResponseLevels;
};