#include <core/stdafx.h>
#include <core/smartview/ReportTag.h>
#include <core/interpret/flags.h>
#include <core/mapi/extraPropTags.h>

namespace smartview
{
	void ReportTag::Parse()
	{
		m_Cookie.parse(m_Parser, 9);

		// Version is big endian, so we have to read individual bytes
		const auto& hiWord = blockT<WORD>(m_Parser);
		const auto& loWord = blockT<WORD>(m_Parser);
		m_Version.setOffset(hiWord.getOffset());
		m_Version.setSize(hiWord.getSize() + loWord.getSize());
		m_Version.setData(hiWord << 16 | loWord);

		m_cbStoreEntryID.parse<DWORD>(m_Parser);
		m_lpStoreEntryID.parse(m_Parser, m_cbStoreEntryID, _MaxEID);

		m_cbFolderEntryID.parse<DWORD>(m_Parser);
		m_lpFolderEntryID.parse(m_Parser, m_cbFolderEntryID, _MaxEID);

		m_cbMessageEntryID.parse<DWORD>(m_Parser);
		m_lpMessageEntryID.parse(m_Parser, m_cbMessageEntryID, _MaxEID);

		m_cbSearchFolderEntryID.parse<DWORD>(m_Parser);
		m_lpSearchFolderEntryID.parse(m_Parser, m_cbSearchFolderEntryID, _MaxEID);

		m_cbMessageSearchKey.parse<DWORD>(m_Parser);
		m_lpMessageSearchKey.parse(m_Parser, m_cbMessageSearchKey, _MaxEID);

		m_cchAnsiText.parse<DWORD>(m_Parser);
		m_lpszAnsiText.parse(m_Parser, m_cchAnsiText);
	}

	void ReportTag::addEID(const std::wstring& label, const blockT<ULONG>& cb, blockBytes& eid)
	{
		if (cb)
		{
			terminateBlock();
			addHeader(label);
			addChild(eid);
		}
	}

	void ReportTag::ParseBlocks()
	{
		setRoot(L"Report Tag: \r\n");
		addHeader(L"Cookie = ");
		addChild(m_Cookie);

		terminateBlock();
		auto szFlags = flags::InterpretFlags(flagReportTagVersion, m_Version);
		addChild(m_Version, L"Version = 0x%1!08X! = %2!ws!", m_Version.getData(), szFlags.c_str());

		addEID(L"StoreEntryID = ", m_cbStoreEntryID, m_lpStoreEntryID);
		addEID(L"FolderEntryID = ", m_cbFolderEntryID, m_lpFolderEntryID);
		addEID(L"MessageEntryID = ", m_cbMessageEntryID, m_lpMessageEntryID);
		addEID(L"SearchFolderEntryID = ", m_cbSearchFolderEntryID, m_lpSearchFolderEntryID);
		addEID(L"MessageSearchKey = ", m_cbMessageSearchKey, m_lpMessageSearchKey);

		if (m_cchAnsiText)
		{
			terminateBlock();
			addChild(m_cchAnsiText, L"cchAnsiText = 0x%1!08X!\r\n", m_cchAnsiText.getData());
			addChild(m_lpszAnsiText, L"AnsiText = %1!hs!", m_lpszAnsiText.c_str());
		}
	}
} // namespace smartview