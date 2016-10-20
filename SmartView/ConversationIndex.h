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
	~ConversationIndex();

private:
	void Parse();
	_Check_return_ wstring ToStringInternal();

	BYTE m_UnnamedByte;
	FILETIME m_ftCurrent;
	GUID m_guid;
	ULONG m_ulResponseLevels;
	ResponseLevel* m_lpResponseLevels;
};