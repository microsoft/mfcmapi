#include <StdAfx.h>
#include <Interpret/SmartView/ReportTag.h>
#include <Interpret/InterpretProp.h>
#include <Interpret/ExtraPropTags.h>

namespace smartview
{
	void ReportTag::Parse()
	{
		m_Cookie = m_Parser.GetBYTES(9);

		// Version is big endian, so we have to read individual bytes
		const auto hiWord = m_Parser.Get<WORD>();
		const auto loWord = m_Parser.Get<WORD>();
		m_Version.setOffset(hiWord.getOffset());
		m_Version.setSize(hiWord.getSize() + loWord.getSize());
		m_Version.setData(hiWord << 16 | loWord);

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

	void ReportTag::ParseBlocks()
	{
		addHeader(L"Report Tag: \r\n");
		addHeader(L"Cookie = ");
		addBlock(m_Cookie);

		addLine();
		auto szFlags = interpretprop::InterpretFlags(flagReportTagVersion, m_Version);
		addBlock(m_Version, L"Version = 0x%1!08X! = %2!ws!", m_Version.getData(), szFlags.c_str());

		if (m_cbStoreEntryID)
		{
			addHeader(L"\r\nStoreEntryID = ");
			addBlock(m_lpStoreEntryID);
		}

		if (m_cbFolderEntryID)
		{
			addHeader(L"\r\nFolderEntryID = ");
			addBlock(m_lpFolderEntryID);
		}

		if (m_cbMessageEntryID)
		{
			addHeader(L"\r\nMessageEntryID = ");
			addBlock(m_lpMessageEntryID);
		}

		if (m_cbSearchFolderEntryID)
		{
			addHeader(L"\r\nSearchFolderEntryID = ");
			addBlock(m_lpSearchFolderEntryID);
		}

		if (m_cbMessageSearchKey)
		{
			addHeader(L"\r\nMessageSearchKey = ");
			addBlock(m_lpMessageSearchKey);
		}

		if (m_cchAnsiText)
		{
			addLine();
			addBlock(m_cchAnsiText, L"cchAnsiText = 0x%1!08X!\r\n", m_cchAnsiText.getData());
			addBlock(m_lpszAnsiText, L"AnsiText = %1!hs!", m_lpszAnsiText.c_str());
		}
	}
}