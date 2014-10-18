#include "stdafx.h"
#include "..\stdafx.h"
#include "VerbStream.h"
#include "..\String.h"
#include "SmartView.h"
#include "..\ExtraPropTags.h"

VerbStream::VerbStream(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin) : SmartViewParser(cbBin, lpBin)
{
	m_Version = 0;
	m_Count = 0;
	m_lpVerbData = 0;
	m_Version2 = 0;
	m_lpVerbExtraData = 0;
}

VerbStream::~VerbStream()
{
	if (m_Count && m_lpVerbData)
	{
		ULONG i = 0;

		for (i = 0; i < m_Count; i++)
		{
			delete[] m_lpVerbData[i].DisplayName;
			delete[] m_lpVerbData[i].MsgClsName;
			delete[] m_lpVerbData[i].Internal1String;
			delete[] m_lpVerbData[i].DisplayNameRepeat;
		}

		delete[] m_lpVerbData;
	}

	if (m_Count && m_lpVerbExtraData)
	{
		ULONG i = 0;

		for (i = 0; i < m_Count; i++)
		{
			delete[] m_lpVerbExtraData[i].DisplayName;
			delete[] m_lpVerbExtraData[i].DisplayNameRepeat;
		}

		delete[] m_lpVerbExtraData;
	}
}

void VerbStream::Parse()
{
	if (!m_lpBin) return;

	m_Parser.GetWORD(&m_Version);
	m_Parser.GetDWORD(&m_Count);

	if (m_Count && m_Count < _MaxEntriesSmall)
		m_lpVerbData = new VerbDataStruct[m_Count];

	if (m_lpVerbData)
	{
		memset(m_lpVerbData, 0, sizeof(VerbDataStruct)*m_Count);
		ULONG i = 0;

		for (i = 0; i < m_Count; i++)
		{
			m_Parser.GetDWORD(&m_lpVerbData[i].VerbType);
			m_Parser.GetBYTE(&m_lpVerbData[i].DisplayNameCount);
			m_Parser.GetStringA(m_lpVerbData[i].DisplayNameCount, &m_lpVerbData[i].DisplayName);
			m_Parser.GetBYTE(&m_lpVerbData[i].MsgClsNameCount);
			m_Parser.GetStringA(m_lpVerbData[i].MsgClsNameCount, &m_lpVerbData[i].MsgClsName);
			m_Parser.GetBYTE(&m_lpVerbData[i].Internal1StringCount);
			m_Parser.GetStringA(m_lpVerbData[i].Internal1StringCount, &m_lpVerbData[i].Internal1String);
			m_Parser.GetBYTE(&m_lpVerbData[i].DisplayNameCountRepeat);
			m_Parser.GetStringA(m_lpVerbData[i].DisplayNameCountRepeat, &m_lpVerbData[i].DisplayNameRepeat);
			m_Parser.GetDWORD(&m_lpVerbData[i].Internal2);
			m_Parser.GetBYTE(&m_lpVerbData[i].Internal3);
			m_Parser.GetDWORD(&m_lpVerbData[i].fUseUSHeaders);
			m_Parser.GetDWORD(&m_lpVerbData[i].Internal4);
			m_Parser.GetDWORD(&m_lpVerbData[i].SendBehavior);
			m_Parser.GetDWORD(&m_lpVerbData[i].Internal5);
			m_Parser.GetDWORD(&m_lpVerbData[i].ID);
			m_Parser.GetDWORD(&m_lpVerbData[i].Internal6);
		}
	}

	m_Parser.GetWORD(&m_Version2);

	if (m_Count && m_Count < _MaxEntriesSmall)
		m_lpVerbExtraData = new VerbExtraDataStruct[m_Count];
	if (m_lpVerbExtraData)
	{
		memset(m_lpVerbExtraData, 0, sizeof(VerbExtraDataStruct)*m_Count);
		ULONG i = 0;

		for (i = 0; i < m_Count; i++)
		{
			m_Parser.GetBYTE(&m_lpVerbExtraData[i].DisplayNameCount);
			m_Parser.GetStringW(m_lpVerbExtraData[i].DisplayNameCount, &m_lpVerbExtraData[i].DisplayName);
			m_Parser.GetBYTE(&m_lpVerbExtraData[i].DisplayNameCountRepeat);
			m_Parser.GetStringW(m_lpVerbExtraData[i].DisplayNameCountRepeat, &m_lpVerbExtraData[i].DisplayNameRepeat);
		}
	}
}

_Check_return_ LPWSTR VerbStream::ToString()
{
	Parse();

	wstring szVerbString;
	wstring szTmp;

	szVerbString = formatmessage(IDS_VERBHEADER, m_Version, m_Count);

	if (m_Count && m_lpVerbData)
	{
		ULONG i = 0;
		for (i = 0; i < m_Count; i++)
		{
			LPWSTR szVerb = NULL;
			InterpretNumberAsStringProp(m_lpVerbData[i].ID, PR_LAST_VERB_EXECUTED, &szVerb);
			szTmp = formatmessage(IDS_VERBDATA,
				i,
				m_lpVerbData[i].VerbType,
				m_lpVerbData[i].DisplayNameCount,
				m_lpVerbData[i].DisplayName,
				m_lpVerbData[i].MsgClsNameCount,
				m_lpVerbData[i].MsgClsName,
				m_lpVerbData[i].Internal1StringCount,
				m_lpVerbData[i].Internal1String,
				m_lpVerbData[i].DisplayNameCountRepeat,
				m_lpVerbData[i].DisplayNameRepeat,
				m_lpVerbData[i].Internal2,
				m_lpVerbData[i].Internal3,
				m_lpVerbData[i].fUseUSHeaders,
				m_lpVerbData[i].Internal4,
				m_lpVerbData[i].SendBehavior,
				m_lpVerbData[i].Internal5,
				m_lpVerbData[i].ID,
				szVerb,
				m_lpVerbData[i].Internal6);
			delete[] szVerb;
			szVerbString += szTmp;
		}
	}

	szTmp = formatmessage(IDS_VERBVERSION2, m_Version2);
	szVerbString += szTmp;

	if (m_Count && m_lpVerbData)
	{
		ULONG i = 0;
		for (i = 0; i < m_Count; i++)
		{
			szTmp = formatmessage(IDS_VERBEXTRADATA,
				i,
				m_lpVerbExtraData[i].DisplayNameCount,
				m_lpVerbExtraData[i].DisplayName,
				m_lpVerbExtraData[i].DisplayNameCountRepeat,
				m_lpVerbExtraData[i].DisplayNameRepeat);
			szVerbString += szTmp;
		}
	}

	szVerbString += JunkDataToString();

	return wstringToLPWSTR(szVerbString);
}