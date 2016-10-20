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
	ExtendedFlags();
	~ExtendedFlags();

private:
	void Parse() override;
	_Check_return_ wstring ToStringInternal() override;

	ULONG m_ulNumFlags;
	ExtendedFlag* m_pefExtendedFlags;
};