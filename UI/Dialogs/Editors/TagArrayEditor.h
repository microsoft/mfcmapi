#pragma once
#include <UI/Dialogs/Editors/Editor.h>

namespace dialog
{
	namespace editor
	{
		class CTagArrayEditor : public CEditor
		{
		public:
			CTagArrayEditor(
				_In_opt_ CWnd* pParentWnd,
				UINT uidTitle,
				UINT uidPrompt,
				_In_opt_ LPMAPITABLE lpContentsTable,
				_In_opt_ LPSPropTagArray lpTagArray,
				bool bIsAB,
				_In_opt_ LPMAPIPROP lpMAPIProp);
			virtual ~CTagArrayEditor();

			_Check_return_ LPSPropTagArray DetachModifiedTagArray();
			_Check_return_ bool DoListEdit(ULONG ulListNum, int iItem, _In_ controls::sortlistdata::SortListData* lpData) override;

		private:
			BOOL OnInitDialog() override;
			void OnOK() override;
			void OnEditAction1() override;
			void OnEditAction2() override;

			void ReadTagArrayToList(ULONG ulListNum, LPSPropTagArray lpTagArray) const;
			void WriteListToTagArray(ULONG ulListNum);

			// source variables
			LPMAPITABLE m_lpContentsTable;
			LPSPropTagArray m_lpSourceTagArray; // Source tag array - not our memory - do not modify or free
			LPSPropTagArray m_lpOutputTagArray;
			bool m_bIsAB; // whether the tag is from the AB or not
			LPMAPIPROP m_lpMAPIProp;
			ULONG m_ulSetColumnsFlags;
		};
	}
}