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
	vector<BYTE> lpUnknownData;
};

class ExtendedFlags : public SmartViewParser
{
public:
	ExtendedFlags();

private:
	void Parse() override;
	_Check_return_ std::wstring ToStringInternal() override;

	ULONG m_ulNumFlags;
	vector<ExtendedFlag> m_pefExtendedFlags;
};