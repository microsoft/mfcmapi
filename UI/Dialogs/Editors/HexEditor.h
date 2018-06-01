#pragma once

#include <UI/Dialogs/Editors/Editor.h>
#include <UI/ParentWnd.h>
#include <MAPI/Cache/MapiObjects.h>

namespace dialog
{
	namespace editor
	{
		class CHexEditor : public CEditor
		{
		public:
			CHexEditor(
				_In_ ui::CParentWnd* pParentWnd,
				_In_ cache::CMapiObjects* lpMapiObjects);
			virtual ~CHexEditor();

		private:
			_Check_return_ ULONG HandleChange(UINT nID) override;
			void OnEditAction1() override;
			void OnEditAction2() override;
			void OnEditAction3() override;
			void UpdateParser() const;

			void OnOK() override;
			void OnCancel() override;

			cache::CMapiObjects* m_lpMapiObjects;
		};
	}
}