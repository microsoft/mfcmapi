#include <StdAfx.h>
#include <UI/Dialogs/Editors/PropertySelector.h>
#include <core/utility/strings.h>
#include <UI/Controls/SortList/PropListData.h>
#include <Interpret/InterpretProp.h>
#include <AddIns.h>
#include <IO/MFCOutput.h>

namespace dialog
{
	namespace editor
	{
		static std::wstring CLASS = L"CPropertySelector";

		// Property selection dialog
		// Displays a list of known property tags - no add or delete
		CPropertySelector::CPropertySelector(bool bIncludeABProps, _In_ LPMAPIPROP lpMAPIProp, _In_ CWnd* pParentWnd)
			: CEditor(pParentWnd, IDS_PROPSELECTOR, 0, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL)
		{
			TRACE_CONSTRUCTOR(CLASS);
			m_ulPropTag = PR_NULL;
			m_bIncludeABProps = bIncludeABProps;
			m_lpMAPIProp = lpMAPIProp;

			if (m_lpMAPIProp) m_lpMAPIProp->AddRef();

			AddPane(viewpane::ListPane::Create(0, IDS_KNOWNPROPTAGS, true, true, ListEditCallBack(this)));
			SetListID(0);
		}

		CPropertySelector::~CPropertySelector()
		{
			TRACE_DESTRUCTOR(CLASS);
			if (m_lpMAPIProp) m_lpMAPIProp->Release();
		}

		BOOL CPropertySelector::OnInitDialog()
		{
			const auto bRet = CEditor::OnInitDialog();

			InsertColumn(0, 1, IDS_PROPERTYNAMES);
			InsertColumn(0, 2, IDS_TAG);
			InsertColumn(0, 3, IDS_TYPE);

			ULONG ulCurRow = 0;
			for (size_t i = 0; i < PropTagArray.size(); i++)
			{
				if (!m_bIncludeABProps && PropTagArray[i].ulValue & 0x80000000) continue;
				auto lpData = InsertListRow(0, ulCurRow, PropTagArray[i].lpszName);

				if (lpData)
				{
					lpData->InitializePropList(PropTagArray[i].ulValue);
				}

				SetListString(0, ulCurRow, 1, strings::format(L"0x%08X", PropTagArray[i].ulValue)); // STRING_OK
				SetListString(0, ulCurRow, 2, interpretprop::TypeToString(PropTagArray[i].ulValue));
				ulCurRow++;
			}

			// Initial sort is by property tag
			ResizeList(0, true);

			return bRet;
		}

		void CPropertySelector::OnOK()
		{
			const auto lpListData = GetSelectedListRowData(0);
			if (lpListData && lpListData->Prop())
			{
				m_ulPropTag = lpListData->Prop()->m_ulPropTag;
			}

			CEditor::OnOK();
		}

		// We're not actually editing the list here - just overriding this to allow double-click
		// So it's OK to return false
		_Check_return_ bool CPropertySelector::DoListEdit(
			ULONG /*ulListNum*/,
			int /*iItem*/,
			_In_ controls::sortlistdata::SortListData* /*lpData*/)
		{
			OnOK();
			return false;
		}

		_Check_return_ ULONG CPropertySelector::GetPropertyTag() const { return m_ulPropTag; }

		_Check_return_ controls::sortlistdata::SortListData* CPropertySelector::GetSelectedListRowData(ULONG id) const
		{
			const auto lpPane = static_cast<viewpane::ListPane*>(GetPane(id));
			if (lpPane)
			{
				return lpPane->GetSelectedListRowData();
			}

			return nullptr;
		}
	} // namespace editor
} // namespace dialog