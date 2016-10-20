#include "stdafx.h"
#include "ReportTag.h"
#include "String.h"
#include "InterpretProp2.h"
#include "ExtraPropTags.h"

ReportTag::ReportTag()
{
	memset(m_Cookie, 0, sizeof m_Cookie);
	m_Version = 0;
	m_cbStoreEntryID = 0;
	m_lpStoreEntryID = nullptr;
	m_cbFolderEntryID = 0;
	m_lpFolderEntryID = nullptr;
	m_cbMessageEntryID = 0;
	m_lpMessageEntryID = nullptr;
	m_cbSearchFolderEntryID = 0;
	m_lpSearchFolderEntryID = nullptr;
	m_cbMessageSearchKey = 0;
	m_lpMessageSearchKey = nullptr;
	m_cchAnsiText = 0;
	m_lpszAnsiText = nullptr;
}

ReportTag::~ReportTag()
{
	delete[] m_lpStoreEntryID;
	delete[] m_lpFolderEntryID;
	delete[] m_lpMessageEntryID;
	delete[] m_lpSearchFolderEntryID;
	delete[] m_lpMessageSearchKey;
	delete[] m_lpszAnsiText;
}

void ReportTag::Parse()
{
	m_Parser.GetBYTESNoAlloc(sizeof m_Cookie, sizeof m_Cookie, reinterpret_cast<LPBYTE>(m_Cookie));

	// Version is big endian, so we have to read individual bytes
	WORD hiWord = NULL;
	WORD loWord = NULL;
	m_Parser.GetWORD(&hiWord);
	m_Parser.GetWORD(&loWord);
	m_Version = hiWord << 16 | loWord;

	m_Parser.GetDWORD(&m_cbStoreEntryID);
	if (m_cbStoreEntryID)
	{
		m_Parser.GetBYTES(m_cbStoreEntryID, _MaxEID, &m_lpStoreEntryID);
	}

	m_Parser.GetDWORD(&m_cbFolderEntryID);
	if (m_cbFolderEntryID)
	{
		m_Parser.GetBYTES(m_cbFolderEntryID, _MaxEID, &m_lpFolderEntryID);
	}

	m_Parser.GetDWORD(&m_cbMessageEntryID);
	if (m_cbMessageEntryID)
	{
		m_Parser.GetBYTES(m_cbMessageEntryID, _MaxEID, &m_lpMessageEntryID);
	}

	m_Parser.GetDWORD(&m_cbSearchFolderEntryID);
	if (m_cbSearchFolderEntryID)
	{
		m_Parser.GetBYTES(m_cbSearchFolderEntryID, _MaxEID, &m_lpSearchFolderEntryID);
	}

	m_Parser.GetDWORD(&m_cbMessageSearchKey);
	if (m_cbMessageSearchKey)
	{
		m_Parser.GetBYTES(m_cbMessageSearchKey, _MaxEID, &m_lpMessageSearchKey);
	}

	m_Parser.GetDWORD(&m_cchAnsiText);
	if (m_cchAnsiText)
	{
		m_Parser.GetStringA(m_cchAnsiText, &m_lpszAnsiText);
	}
}

_Check_return_ wstring ReportTag::ToStringInternal()
{
	wstring szReportTag;

	szReportTag = formatmessage(IDS_REPORTTAGHEADER);

	SBinary sBin = { 0 };
	sBin.cb = sizeof m_Cookie;
	sBin.lpb = reinterpret_cast<LPBYTE>(m_Cookie);
	szReportTag += BinToHexString(&sBin, true);

	auto szFlags = InterpretFlags(flagReportTagVersion, m_Version);
	szReportTag += formatmessage(IDS_REPORTTAGVERSION,
		m_Version,
		szFlags.c_str());

	if (m_cbStoreEntryID)
	{
		szReportTag += formatmessage(IDS_REPORTTAGSTOREEID);
		sBin.cb = m_cbStoreEntryID;
		sBin.lpb = m_lpStoreEntryID;
		szReportTag += BinToHexString(&sBin, true);
	}

	if (m_cbFolderEntryID)
	{
		szReportTag += formatmessage(IDS_REPORTTAGFOLDEREID);
		sBin.cb = m_cbFolderEntryID;
		sBin.lpb = m_lpFolderEntryID;
		szReportTag += BinToHexString(&sBin, true);
	}

	if (m_cbMessageEntryID)
	{
		szReportTag += formatmessage(IDS_REPORTTAGMESSAGEEID);
		sBin.cb = m_cbMessageEntryID;
		sBin.lpb = m_lpMessageEntryID;
		szReportTag += BinToHexString(&sBin, true);
	}

	if (m_cbSearchFolderEntryID)
	{
		szReportTag += formatmessage(IDS_REPORTTAGSFEID);
		sBin.cb = m_cbSearchFolderEntryID;
		sBin.lpb = m_lpSearchFolderEntryID;
		szReportTag += BinToHexString(&sBin, true);
	}

	if (m_cbMessageSearchKey)
	{
		szReportTag += formatmessage(IDS_REPORTTAGMESSAGEKEY);
		sBin.cb = m_cbMessageSearchKey;
		sBin.lpb = m_lpMessageSearchKey;
		szReportTag += BinToHexString(&sBin, true);
	}

	if (m_cchAnsiText)
	{
		szReportTag += formatmessage(IDS_REPORTTAGANSITEXT,
			m_cchAnsiText,
			m_lpszAnsiText ? m_lpszAnsiText : ""); // STRING_OK
	}

	return szReportTag;
}