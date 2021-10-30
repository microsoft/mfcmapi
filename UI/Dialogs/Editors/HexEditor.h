#pragma once

#include <UI/Dialogs/Editors/Editor.h>
#include <UI/ParentWnd.h>
#include <core/mapi/cache/mapiObjects.h>

namespace dialog::editor
{
	class CHexEditor : public CEditor
	{
	public:
		CHexEditor(_In_ std::shared_ptr<cache::CMapiObjects> lpMapiObjects);

	private:
		_Check_return_ ULONG HandleChange(UINT nID) override;
		void OnEditAction1() override;
		void OnEditAction2() override;
		void OnEditAction3() override;
		void UpdateParser() const;

		void OnOK() override;
		void OnCancel() override;
		void SetHex(_In_opt_count_(cb) LPBYTE lpb, size_t cb) const;

		std::shared_ptr<cache::CMapiObjects> m_lpMapiObjects{};
	};
} // namespace dialog::editor