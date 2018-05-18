#pragma once
#include "SmartViewParser.h"

struct VerbData
{
	DWORD VerbType;
	BYTE DisplayNameCount;
	std::string DisplayName;
	BYTE MsgClsNameCount;
	std::string MsgClsName;
	BYTE Internal1StringCount;
	std::string Internal1String;
	BYTE DisplayNameCountRepeat;
	std::string DisplayNameRepeat;
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
	std::wstring DisplayName;
	BYTE DisplayNameCountRepeat;
	std::wstring DisplayNameRepeat;
};

class VerbStream : public SmartViewParser
{
public:
	VerbStream();

private:
	void Parse() override;
	_Check_return_ std::wstring ToStringInternal() override;

	WORD m_Version;
	DWORD m_Count;
	std::vector<VerbData> m_lpVerbData;
	WORD m_Version2;
	std::vector<VerbExtraData> m_lpVerbExtraData;
};