#include <core/stdafx.h>
#include <core/smartview/ReportTag.h>
#include <core/interpret/flags.h>
#include <core/mapi/extraPropTags.h>

namespace smartview
{
	void ReportTag::Parse()
	{
		m_Cookie = m_Parser->GetBYTES(9);

		// Version is big endian, so we have to read individual bytes
		const auto hiWord = m_Parser->Get<WORD>();
		const auto loWord = m_Parser->Get<WORD>();
		m_Version.setOffset(hiWord.getOffset());
		m_Version.setSize(hiWord.getSize() + loWord.getSize());
		m_Version.setData(hiWord << 16 | loWord);

		m_cbStoreEntryID = m_Parser->Get<DWORD>();
		m_lpStoreEntryID = m_Parser->GetBYTES(m_cbStoreEntryID, _MaxEID);

		m_cbFolderEntryID = m_Parser->Get<DWORD>();
		m_lpFolderEntryID = m_Parser->GetBYTES(m_cbFolderEntryID, _MaxEID);

		m_cbMessageEntryID = m_Parser->Get<DWORD>();
		m_lpMessageEntryID = m_Parser->GetBYTES(m_cbMessageEntryID, _MaxEID);

		m_cbSearchFolderEntryID = m_Parser->Get<DWORD>();
		m_lpSearchFolderEntryID = m_Parser->GetBYTES(m_cbSearchFolderEntryID, _MaxEID);

		m_cbMessageSearchKey = m_Parser->Get<DWORD>();
		m_lpMessageSearchKey = m_Parser->GetBYTES(m_cbMessageSearchKey, _MaxEID);

		m_cchAnsiText = m_Parser->Get<DWORD>();
		m_lpszAnsiText.init(m_Parser, m_cchAnsiText);
	}

	void ReportTag::addEID(const std::wstring& label, const blockT<ULONG>& cb, const blockBytes& eid)
	{
		if (cb)
		{
			terminateBlock();
			addHeader(label);
			addBlock(eid);
		}
	}

	void ReportTag::ParseBlocks()
	{
		setRoot(L"Report Tag: \r\n");
		addHeader(L"Cookie = ");
		addBlock(m_Cookie);

		terminateBlock();
		auto szFlags = flags::InterpretFlags(flagReportTagVersion, m_Version);
		addBlock(m_Version, L"Version = 0x%1!08X! = %2!ws!", m_Version.getData(), szFlags.c_str());

		addEID(L"StoreEntryID = ", m_cbStoreEntryID, m_lpStoreEntryID);
		addEID(L"FolderEntryID = ", m_cbFolderEntryID, m_lpFolderEntryID);
		addEID(L"MessageEntryID = ", m_cbMessageEntryID, m_lpMessageEntryID);
		addEID(L"SearchFolderEntryID = ", m_cbSearchFolderEntryID, m_lpSearchFolderEntryID);
		addEID(L"MessageSearchKey = ", m_cbMessageSearchKey, m_lpMessageSearchKey);

		if (m_cchAnsiText)
		{
			terminateBlock();
			addBlock(m_cchAnsiText, L"cchAnsiText = 0x%1!08X!\r\n", m_cchAnsiText.getData());
			addBlock(m_lpszAnsiText, L"AnsiText = %1!hs!", m_lpszAnsiText.c_str());
		}
	}
} // namespace smartview