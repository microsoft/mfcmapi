#pragma once
#include "SmartViewParser.h"

struct ExtendedFlag
{
	BYTE Id;
	BYTE Cb;
	union
	{
		DWORD ExtendedFlags;
		GUID SearchFolderID;
		DWORD SearchFolderTag;
		DWORD ToDoFolderVersion;
	} Data;
	BYTE* lpUnknownData;
};

class ExtendedFlags : public SmartViewParser
{
public:
	ExtendedFlags(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin);
	~ExtendedFlags();

private:
	void Parse();
	_Check_return_ wstring ToStringInternal();

	ULONG m_ulNumFlags;
	ExtendedFlag* m_pefExtendedFlags;
};