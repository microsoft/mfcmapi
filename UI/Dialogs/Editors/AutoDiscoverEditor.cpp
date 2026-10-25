#include <StdAfx.h>
#include <UI/Dialogs/Editors/AutoDiscoverEditor.h>
#include <UI/file/FileDialogEx.h>
#include <core/utility/error.h>
#include <core/utility/strings.h>
#include <core/addin/mfcmapi.h>

namespace dialog::editor
{
	CAutoDiscoverEditor::CAutoDiscoverEditor(_In_ CWnd* parent, const std::vector<BYTE>& xml, const std::wstring& text)
		: CEditor(
			  parent,
			  IDS_AUTODISCOVER,
			  IDS_AUTODISCOVERRESULT,
			  CEDITOR_BUTTON_OK | CEDITOR_BUTTON_ACTION1,
			  IDS_AUTODISCOVERSAVE,
			  0,
			  0),
		  m_xml(xml)
	{
		AddPane(viewpane::TextPane::CreateMultiLinePane(0, IDS_AUTODISCOVERXML, text, true));
	}

	void CAutoDiscoverEditor::OnEditAction1()
	{
		const auto path = file::CFileDialogExW::SaveAs(
			L"xml", // STRING_OK
			L"Autodiscover.xml", // STRING_OK
			OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT,
			strings::loadstring(IDS_XMLFILES),
			this);
		if (path.empty()) return;

		if (m_xml.size() > MAXDWORD)
		{
			EC_H_S(MAPI_E_TOO_BIG);
			return;
		}

		const auto handle =
			CreateFileW(path.c_str(), GENERIC_WRITE, 0, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
		auto hRes = EC_B(handle != INVALID_HANDLE_VALUE);
		if (FAILED(hRes)) return;

		DWORD written{};
		hRes = EC_B(WriteFile(handle, m_xml.data(), static_cast<DWORD>(m_xml.size()), &written, nullptr));
		if (SUCCEEDED(hRes) && written != m_xml.size()) EC_H_S(STG_E_WRITEFAULT);

		EC_B_S(CloseHandle(handle));
	}
} // namespace dialog::editor
