#include <StdAfx.h>
#include <Interpret/SmartView/ReportTag.h>
#include <Interpret/InterpretProp.h>
#include <Interpret/ExtraPropTags.h>

namespace smartview
{
	ReportTag::ReportTag() {}

	void ReportTag::Parse()
	{
		m_Cookie = m_Parser.GetBlockBYTES(9);

		// Version is big endian, so we have to read individual bytes
		const auto hiWord = m_Parser.GetBlock<WORD>();
		const auto loWord = m_Parser.GetBlock<WORD>();
		m_Version.setOffset(hiWord.getOffset());
		m_Version.setSize(hiWord.getSize() + loWord.getSize());
		m_Version.setData(hiWord.getData() << 16 | loWord.getData());

		m_cbStoreEntryID = m_Parser.GetBlock<DWORD>();
		if (m_cbStoreEntryID.getData())
		{
			m_lpStoreEntryID = m_Parser.GetBlockBYTES(m_cbStoreEntryID.getData(), _MaxEID);
		}

		m_cbFolderEntryID = m_Parser.GetBlock<DWORD>();
		if (m_cbFolderEntryID.getData())
		{
			m_lpFolderEntryID = m_Parser.GetBlockBYTES(m_cbFolderEntryID.getData(), _MaxEID);
		}

		m_cbMessageEntryID = m_Parser.GetBlock<DWORD>();
		if (m_cbMessageEntryID.getData())
		{
			m_lpMessageEntryID = m_Parser.GetBlockBYTES(m_cbMessageEntryID.getData(), _MaxEID);
		}

		m_cbSearchFolderEntryID = m_Parser.GetBlock<DWORD>();
		if (m_cbSearchFolderEntryID.getData())
		{
			m_lpSearchFolderEntryID = m_Parser.GetBlockBYTES(m_cbSearchFolderEntryID.getData(), _MaxEID);
		}

		m_cbMessageSearchKey = m_Parser.GetBlock<DWORD>();
		if (m_cbMessageSearchKey.getData())
		{
			m_lpMessageSearchKey = m_Parser.GetBlockBYTES(m_cbMessageSearchKey.getData(), _MaxEID);
		}

		m_cchAnsiText = m_Parser.GetBlock<DWORD>();
		if (m_cchAnsiText.getData())
		{
			m_lpszAnsiText = m_Parser.GetBlockStringA(m_cchAnsiText.getData());
		}
	}

	_Check_return_ std::wstring ReportTag::ToStringInternal()
	{
		auto szReportTag = strings::formatmessage(IDS_REPORTTAGHEADER);

		szReportTag += strings::BinToHexString(m_Cookie.getData(), true);

		auto szFlags = interpretprop::InterpretFlags(flagReportTagVersion, m_Version.getData());
		szReportTag += strings::formatmessage(IDS_REPORTTAGVERSION, m_Version.getData(), szFlags.c_str());

		if (m_cbStoreEntryID.getData())
		{
			szReportTag += strings::formatmessage(IDS_REPORTTAGSTOREEID);
			szReportTag += strings::BinToHexString(m_lpStoreEntryID.getData(), true);
		}

		if (m_cbFolderEntryID.getData())
		{
			szReportTag += strings::formatmessage(IDS_REPORTTAGFOLDEREID);
			szReportTag += strings::BinToHexString(m_lpFolderEntryID.getData(), true);
		}

		if (m_cbMessageEntryID.getData())
		{
			szReportTag += strings::formatmessage(IDS_REPORTTAGMESSAGEEID);
			szReportTag += strings::BinToHexString(m_lpMessageEntryID.getData(), true);
		}

		if (m_cbSearchFolderEntryID.getData())
		{
			szReportTag += strings::formatmessage(IDS_REPORTTAGSFEID);
			szReportTag += strings::BinToHexString(m_lpSearchFolderEntryID.getData(), true);
		}

		if (m_cbMessageSearchKey.getData())
		{
			szReportTag += strings::formatmessage(IDS_REPORTTAGMESSAGEKEY);
			szReportTag += strings::BinToHexString(m_lpMessageSearchKey.getData(), true);
		}

		if (m_cchAnsiText.getData())
		{
			szReportTag += strings::formatmessage(
				IDS_REPORTTAGANSITEXT,
				m_cchAnsiText.getData(),
				m_lpszAnsiText.getData().c_str()); // STRING_OK
		}

		return szReportTag;
	}
}