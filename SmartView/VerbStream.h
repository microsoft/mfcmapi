#pragma once
#include "SmartViewParser.h"

struct VerbDataStruct
{
	DWORD VerbType;
	BYTE DisplayNameCount;
	LPSTR DisplayName;
	BYTE MsgClsNameCount;
	LPSTR MsgClsName;
	BYTE Internal1StringCount;
	LPSTR Internal1String;
	BYTE DisplayNameCountRepeat;
	LPSTR DisplayNameRepeat;
	DWORD Internal2;
	BYTE Internal3;
	DWORD fUseUSHeaders;
	DWORD Internal4;
	DWORD SendBehavior;
	DWORD Internal5;
	DWORD ID;
	DWORD Internal6;
};

struct VerbExtraDataStruct
{
	BYTE DisplayNameCount;
	LPWSTR DisplayName;
	BYTE DisplayNameCountRepeat;
	LPWSTR DisplayNameRepeat;
};

class VerbStream : public SmartViewParser
{
public:
	VerbStream(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin);
	~VerbStream();

private:
	void Parse();
	_Check_return_ wstring ToStringInternal();

	WORD m_Version;
	DWORD m_Count;
	VerbDataStruct* m_lpVerbData;
	WORD m_Version2;
	VerbExtraDataStruct* m_lpVerbExtraData;
};