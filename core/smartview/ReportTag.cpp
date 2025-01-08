#include <core/stdafx.h>
#include <core/smartview/ReportTag.h>
#include <core/interpret/flags.h>
#include <core/mapi/extraPropTags.h>

namespace smartview
{
	void ReportTag::parse()
	{
		m_Cookie = blockBytes::parse(parser, 9);

		// Version is big endian, so we have to read individual bytes
		const auto hiWord = blockT<WORD>::parse(parser);
		const auto loWord = blockT<WORD>::parse(parser);
		m_Version =
			blockT<DWORD>::create(*hiWord << 16 | *loWord, hiWord->getSize() + loWord->getSize(), hiWord->getOffset());

		m_cbStoreEntryID = blockT<DWORD>::parse(parser);
		m_lpStoreEntryID = blockBytes::parse(parser, *m_cbStoreEntryID, _MaxEID);

		m_cbFolderEntryID = blockT<DWORD>::parse(parser);
		m_lpFolderEntryID = blockBytes::parse(parser, *m_cbFolderEntryID, _MaxEID);

		m_cbMessageEntryID = blockT<DWORD>::parse(parser);
		m_lpMessageEntryID = blockBytes::parse(parser, *m_cbMessageEntryID, _MaxEID);

		m_cbSearchFolderEntryID = blockT<DWORD>::parse(parser);
		m_lpSearchFolderEntryID = blockBytes::parse(parser, *m_cbSearchFolderEntryID, _MaxEID);

		m_cbMessageSearchKey = blockT<DWORD>::parse(parser);
		m_lpMessageSearchKey = blockBytes::parse(parser, *m_cbMessageSearchKey, _MaxEID);

		m_cchAnsiText = blockT<DWORD>::parse(parser);
		m_lpszAnsiText = blockStringA::parse(parser, *m_cchAnsiText);
	}

	void ReportTag::addEID(
		const std::wstring& label,
		const std::shared_ptr<blockT<ULONG>>& _cb,
		const std::shared_ptr<blockBytes>& eid)
	{
		if (*_cb)
		{
			addLabeledChild(label, eid);
		}
	}

	void ReportTag::parseBlocks()
	{
		setText(L"Report Tag");
		addLabeledChild(L"Cookie", m_Cookie);

		auto szFlags = flags::InterpretFlags(flagReportTagVersion, *m_Version);
		addChild(m_Version, L"Version = 0x%1!08X! = %2!ws!", m_Version->getData(), szFlags.c_str());

		addEID(L"StoreEntryID", m_cbStoreEntryID, m_lpStoreEntryID);
		addEID(L"FolderEntryID", m_cbFolderEntryID, m_lpFolderEntryID);
		addEID(L"MessageEntryID", m_cbMessageEntryID, m_lpMessageEntryID);
		addEID(L"SearchFolderEntryID", m_cbSearchFolderEntryID, m_lpSearchFolderEntryID);
		addEID(L"MessageSearchKey", m_cbMessageSearchKey, m_lpMessageSearchKey);

		if (m_cchAnsiText)
		{
			addChild(m_cchAnsiText, L"cchAnsiText = 0x%1!08X!", m_cchAnsiText->getData());
			addChild(m_lpszAnsiText, L"AnsiText = \"%1!hs!\"", m_lpszAnsiText->c_str());
		}
	}
} // namespace smartview