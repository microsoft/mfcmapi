#pragma once
#include <UI/Dialogs/Editors/Editor.h>

namespace dialog::editor
{
	class CAutoDiscoverEditor : public CEditor
	{
	public:
		CAutoDiscoverEditor(_In_ CWnd* parent, const std::vector<BYTE>& xml, const std::wstring& text);

	private:
		void OnEditAction1() override;

		const std::vector<BYTE>& m_xml;
	};
} // namespace dialog::editor
