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

	void ReportTag::ParseBlocks()
	{
		addHeader(L"Report Tag: \r\n");
		addHeader(L"Cookie = ");
		addBlockBytes(m_Cookie);

		addHeader(L"\r\n");
		auto szFlags = interpretprop::InterpretFlags(flagReportTagVersion, m_Version.getData());
		addBlock(m_Version, L"Version = 0x%1!08X! = %2!ws!", m_Version.getData(), szFlags.c_str());

		if (m_cbStoreEntryID.getData())
		{
			addHeader(L"\r\nStoreEntryID = ");
			addBlockBytes(m_lpStoreEntryID);
		}

		if (m_cbFolderEntryID.getData())
		{
			addHeader(L"\r\nFolderEntryID = ");
			addBlockBytes(m_lpFolderEntryID);
		}

		if (m_cbMessageEntryID.getData())
		{
			addHeader(L"\r\nMessageEntryID = ");
			addBlockBytes(m_lpMessageEntryID);
		}

		if (m_cbSearchFolderEntryID.getData())
		{
			addHeader(L"\r\nSearchFolderEntryID = ");
			addBlockBytes(m_lpSearchFolderEntryID);
		}

		if (m_cbMessageSearchKey.getData())
		{
			addHeader(L"\r\nMessageSearchKey = ");
			addBlockBytes(m_lpMessageSearchKey);
		}

		if (m_cchAnsiText.getData())
		{
			addHeader(L"\r\n");
			addBlock(m_cchAnsiText, L"cchAnsiText = 0x%1!08X!\r\n", m_cchAnsiText.getData());
			addBlock(m_lpszAnsiText, L"AnsiText = %1!hs!", m_lpszAnsiText.getData().c_str());
		}
	}
}