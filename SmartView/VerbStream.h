#pragma once
#include "SmartViewParser.h"

struct VerbData
{
	DWORD VerbType;
	BYTE DisplayNameCount;
	string DisplayName;
	BYTE MsgClsNameCount;
	string MsgClsName;
	BYTE Internal1StringCount;
	string Internal1String;
	BYTE DisplayNameCountRepeat;
	string DisplayNameRepeat;
	DWORD Internal2;
	BYTE Internal3;
	DWORD fUseUSHeaders;
	DWORD Internal4;
	DWORD SendBehavior;
	DWORD Internal5;
	DWORD ID;
	DWORD Internal6;
};

struct VerbExtraData
{
	BYTE DisplayNameCount;
	wstring DisplayName;
	BYTE DisplayNameCountRepeat;
	wstring DisplayNameRepeat;
};

class VerbStream : public SmartViewParser
{
public:
	VerbStream();

private:
	void Parse() override;
	_Check_return_ wstring ToStringInternal() override;

	WORD m_Version;
	DWORD m_Count;
	vector<VerbData> m_lpVerbData;
	WORD m_Version2;
	vector<VerbExtraData> m_lpVerbExtraData;
};