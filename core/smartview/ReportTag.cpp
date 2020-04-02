#include <core/stdafx.h>
#include <core/smartview/ReportTag.h>
#include <core/interpret/flags.h>
#include <core/mapi/extraPropTags.h>

namespace smartview
{
	void ReportTag::parse()
	{
		m_Cookie = blockBytes::parse(m_Parser, 9);

		// Version is big endian, so we have to read individual bytes
		const auto hiWord = blockT<WORD>::parse(m_Parser);
		const auto loWord = blockT<WORD>::parse(m_Parser);
		m_Version =
			blockT<DWORD>::create(*hiWord << 16 | *loWord, hiWord->getSize() + loWord->getSize(), hiWord->getOffset());

		m_cbStoreEntryID = blockT<DWORD>::parse(m_Parser);
		m_lpStoreEntryID = blockBytes::parse(m_Parser, *m_cbStoreEntryID, _MaxEID);

		m_cbFolderEntryID = blockT<DWORD>::parse(m_Parser);
		m_lpFolderEntryID = blockBytes::parse(m_Parser, *m_cbFolderEntryID, _MaxEID);

		m_cbMessageEntryID = blockT<DWORD>::parse(m_Parser);
		m_lpMessageEntryID = blockBytes::parse(m_Parser, *m_cbMessageEntryID, _MaxEID);

		m_cbSearchFolderEntryID = blockT<DWORD>::parse(m_Parser);
		m_lpSearchFolderEntryID = blockBytes::parse(m_Parser, *m_cbSearchFolderEntryID, _MaxEID);

		m_cbMessageSearchKey = blockT<DWORD>::parse(m_Parser);
		m_lpMessageSearchKey = blockBytes::parse(m_Parser, *m_cbMessageSearchKey, _MaxEID);

		m_cchAnsiText = blockT<DWORD>::parse(m_Parser);
		m_lpszAnsiText = blockStringA::parse(m_Parser, *m_cchAnsiText);
	}

	void ReportTag::addEID(
		const std::wstring& label,
		const std::shared_ptr<blockT<ULONG>>& cb,
		const std::shared_ptr<blockBytes>& eid)
	{
		if (*cb)
		{
			terminateBlock();
			addLabeledChild(label, eid);
		}
	}

	void ReportTag::parseBlocks()
	{
		setRoot(L"Report Tag: \r\n");
		addLabeledChild(L"Cookie = ", m_Cookie);

		terminateBlock();
		auto szFlags = flags::InterpretFlags(flagReportTagVersion, *m_Version);
		addChild(m_Version, L"Version = 0x%1!08X! = %2!ws!", m_Version->getData(), szFlags.c_str());

		addEID(L"StoreEntryID = ", m_cbStoreEntryID, m_lpStoreEntryID);
		addEID(L"FolderEntryID = ", m_cbFolderEntryID, m_lpFolderEntryID);
		addEID(L"MessageEntryID = ", m_cbMessageEntryID, m_lpMessageEntryID);
		addEID(L"SearchFolderEntryID = ", m_cbSearchFolderEntryID, m_lpSearchFolderEntryID);
		addEID(L"MessageSearchKey = ", m_cbMessageSearchKey, m_lpMessageSearchKey);

		if (m_cchAnsiText)
		{
			terminateBlock();
			addChild(m_cchAnsiText, L"cchAnsiText = 0x%1!08X!\r\n", m_cchAnsiText->getData());
			addChild(m_lpszAnsiText, L"AnsiText = %1!hs!", m_lpszAnsiText->c_str());
		}
	}
} // namespace smartview