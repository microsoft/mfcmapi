#include "stdafx.h"
#include "ReportTag.h"
#include <Interpret/InterpretProp2.h>
#include <Interpret/ExtraPropTags.h>

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
	auto hiWord = m_Parser.Get<WORD>();
	auto loWord = m_Parser.Get<WORD>();
	m_Version = hiWord << 16 | loWord;

	m_cbStoreEntryID = m_Parser.Get<DWORD>();
	if (m_cbStoreEntryID)
	{
		m_lpStoreEntryID = m_Parser.GetBYTES(m_cbStoreEntryID, _MaxEID);
	}

	m_cbFolderEntryID = m_Parser.Get<DWORD>();
	if (m_cbFolderEntryID)
	{
		m_lpFolderEntryID = m_Parser.GetBYTES(m_cbFolderEntryID, _MaxEID);
	}

	m_cbMessageEntryID = m_Parser.Get<DWORD>();
	if (m_cbMessageEntryID)
	{
		m_lpMessageEntryID = m_Parser.GetBYTES(m_cbMessageEntryID, _MaxEID);
	}

	m_cbSearchFolderEntryID = m_Parser.Get<DWORD>();
	if (m_cbSearchFolderEntryID)
	{
		m_lpSearchFolderEntryID = m_Parser.GetBYTES(m_cbSearchFolderEntryID, _MaxEID);
	}

	m_cbMessageSearchKey = m_Parser.Get<DWORD>();
	if (m_cbMessageSearchKey)
	{
		m_lpMessageSearchKey = m_Parser.GetBYTES(m_cbMessageSearchKey, _MaxEID);
	}

	m_cchAnsiText = m_Parser.Get<DWORD>();
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