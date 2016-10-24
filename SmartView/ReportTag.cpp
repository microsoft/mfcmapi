#include "stdafx.h"
#include "ReportTag.h"
#include "InterpretProp2.h"
#include "ExtraPropTags.h"

ReportTag::ReportTag()
{
	m_Version = 0;
	m_cbStoreEntryID = 0;
	m_cbFolderEntryID = 0;
	m_cbMessageEntryID = 0;
	m_cbSearchFolderEntryID = 0;
	m_cbMessageSearchKey = 0;
	m_cchAnsiText = 0;
}

void ReportTag::Parse()
{
	m_Cookie = m_Parser.GetBYTES(9);

	// Version is big endian, so we have to read individual bytes
	WORD hiWord = NULL;
	WORD loWord = NULL;
	m_Parser.GetWORD(&hiWord);
	m_Parser.GetWORD(&loWord);
	m_Version = hiWord << 16 | loWord;

	m_Parser.GetDWORD(&m_cbStoreEntryID);
	if (m_cbStoreEntryID)
	{
		m_lpStoreEntryID = m_Parser.GetBYTES(m_cbStoreEntryID, _MaxEID);
	}

	m_Parser.GetDWORD(&m_cbFolderEntryID);
	if (m_cbFolderEntryID)
	{
		m_lpFolderEntryID = m_Parser.GetBYTES(m_cbFolderEntryID, _MaxEID);
	}

	m_Parser.GetDWORD(&m_cbMessageEntryID);
	if (m_cbMessageEntryID)
	{
		m_lpMessageEntryID = m_Parser.GetBYTES(m_cbMessageEntryID, _MaxEID);
	}

	m_Parser.GetDWORD(&m_cbSearchFolderEntryID);
	if (m_cbSearchFolderEntryID)
	{
		m_lpSearchFolderEntryID = m_Parser.GetBYTES(m_cbSearchFolderEntryID, _MaxEID);
	}

	m_Parser.GetDWORD(&m_cbMessageSearchKey);
	if (m_cbMessageSearchKey)
	{
		m_lpMessageSearchKey = m_Parser.GetBYTES(m_cbMessageSearchKey, _MaxEID);
	}

	m_Parser.GetDWORD(&m_cchAnsiText);
	if (m_cchAnsiText)
	{
		m_lpszAnsiText = m_Parser.GetStringA(m_cchAnsiText);
	}
}

_Check_return_ wstring ReportTag::ToStringInternal()
{
	wstring szReportTag;

	szReportTag = formatmessage(IDS_REPORTTAGHEADER);

	szReportTag += BinToHexString(m_Cookie, true);

	auto szFlags = InterpretFlags(flagReportTagVersion, m_Version);
	szReportTag += formatmessage(IDS_REPORTTAGVERSION,
		m_Version,
		szFlags.c_str());

	if (m_cbStoreEntryID)
	{
		szReportTag += formatmessage(IDS_REPORTTAGSTOREEID);
		szReportTag += BinToHexString(m_lpStoreEntryID, true);
	}

	if (m_cbFolderEntryID)
	{
		szReportTag += formatmessage(IDS_REPORTTAGFOLDEREID);
		szReportTag += BinToHexString(m_lpFolderEntryID, true);
	}

	if (m_cbMessageEntryID)
	{
		szReportTag += formatmessage(IDS_REPORTTAGMESSAGEEID);
		szReportTag += BinToHexString(m_lpMessageEntryID, true);
	}

	if (m_cbSearchFolderEntryID)
	{
		szReportTag += formatmessage(IDS_REPORTTAGSFEID);
		szReportTag += BinToHexString(m_lpSearchFolderEntryID, true);
	}

	if (m_cbMessageSearchKey)
	{
		szReportTag += formatmessage(IDS_REPORTTAGMESSAGEKEY);
		szReportTag += BinToHexString(m_lpMessageSearchKey, true);
	}

	if (m_cchAnsiText)
	{
		szReportTag += formatmessage(IDS_REPORTTAGANSITEXT,
			m_cchAnsiText,
			m_lpszAnsiText.c_str()); // STRING_OK
	}

	return szReportTag;
}