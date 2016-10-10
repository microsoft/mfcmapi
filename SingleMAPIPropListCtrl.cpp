#include "stdafx.h"
#include "SingleMAPIPropListCtrl.h"
#include "MAPIFunctions.h"
#include "ColumnTags.h"
#include "MFCUtilityFunctions.h"
#include "UIFunctions.h"
#include "MySecInfo.h"
#include "InterpretProp2.h"
#include "FileDialogEx.h"
#include "ImportProcs.h"
#include "MAPIProgress.h"
#include "NamedPropCache.h"
#include "SmartView/SmartView.h"
#include "PropertyBag/PropertyBag.h"
#include "PropertyBag/MAPIPropPropertyBag.h"
#include "PropertyBag/RowPropertyBag.h"
#include "String.h"
#include "SortList/PropListData.h"
#include "GlobalCache.h"
#include "Dialogs/Editors/Editor.h"
#include "Dialogs/Editors/RestrictEditor.h"
#include "Dialogs/Editors/StreamEditor.h"
#include "Dialogs/Editors/TagArrayEditor.h"
#include "Dialogs/Editors/PropertyEditor.h"
#include "Dialogs/Editors/PropertyTagEditor.h"

static wstring CLASS = L"CSingleMAPIPropListCtrl";

// 26 columns should be enough for anybody
#define MAX_SORT_COLS 26

CSingleMAPIPropListCtrl::CSingleMAPIPropListCtrl(
	_In_ CWnd* pCreateParent,
	_In_ CBaseDialog *lpHostDlg,
	_In_ CMapiObjects* lpMapiObjects,
	bool bIsAB)
	:CSortListCtrl()
{
	TRACE_CONSTRUCTOR(CLASS);
	auto hRes = S_OK;

	EC_H(Create(pCreateParent, LVS_SINGLESEL, IDC_LIST_CTRL, true));

	m_bHaveEverDisplayedSomething = false;

	// We borrow our parent's Mapi objects
	m_lpMapiObjects = lpMapiObjects;
	if (m_lpMapiObjects) m_lpMapiObjects->AddRef();

	m_lpHostDlg = lpHostDlg;
	if (m_lpHostDlg) m_lpHostDlg->AddRef();

	for (ULONG i = 0; i < PropColumns.size(); i++)
	{
		auto szHeaderName = loadstring(PropColumns[i].uidName);
		InsertColumn(i, wstringTotstring(szHeaderName).c_str());
	}

	auto lpMyHeader = GetHeaderCtrl();

	// Column orders are stored as lowercase letters
	// bacdefghi would mean the first two columns are swapped
	if (lpMyHeader && !RegKeys[regkeyPROP_COLUMN_ORDER].szCurSTRING.empty())
	{
		auto bSetCols = false;
		auto nColumnCount = lpMyHeader->GetItemCount();
		auto cchOrder = RegKeys[regkeyPROP_COLUMN_ORDER].szCurSTRING.length();
		if (SUCCEEDED(hRes) && nColumnCount == static_cast<int>(cchOrder))
		{
			auto pnOrder = new int[nColumnCount];

			if (pnOrder)
			{
				for (auto i = 0; i < nColumnCount; i++)
				{
					pnOrder[i] = RegKeys[regkeyPROP_COLUMN_ORDER].szCurSTRING[i] - L'a';
				}

				if (SetColumnOrderArray(nColumnCount, pnOrder))
				{
					bSetCols = true;
				}
			}

			delete[] pnOrder;
		}

		// If we didn't like the reg key, clear it so we don't see it again
		if (!bSetCols) RegKeys[regkeyPROP_COLUMN_ORDER].szCurSTRING.clear();
	}

	AutoSizeColumns(false);

	m_sptExtraProps = nullptr;

	m_bIsAB = bIsAB;
	m_lpPropBag = nullptr;
}

CSingleMAPIPropListCtrl::~CSingleMAPIPropListCtrl()
{
	TRACE_DESTRUCTOR(CLASS);
	if (m_sptExtraProps) MAPIFreeBuffer(m_sptExtraProps);
	if (m_lpPropBag) delete m_lpPropBag;
	if (m_lpMapiObjects) m_lpMapiObjects->Release();
	if (m_lpHostDlg) m_lpHostDlg->Release();
}

BEGIN_MESSAGE_MAP(CSingleMAPIPropListCtrl, CSortListCtrl)
	ON_NOTIFY_REFLECT(NM_DBLCLK, OnDblclk)
	ON_WM_KEYDOWN()
	ON_WM_CONTEXTMENU()
	ON_MESSAGE(WM_MFCMAPI_SAVECOLUMNORDERLIST, msgOnSaveColumnOrder)
END_MESSAGE_MAP()

LRESULT CSingleMAPIPropListCtrl::WindowProc(UINT message, WPARAM wParam, LPARAM lParam)
{
	switch (message)
	{
	case WM_ERASEBKGND:
		if (!m_lpPropBag)
		{
			return true;
		}

		break;
	case WM_PAINT:
		if (!m_lpPropBag)
		{
			DrawHelpText(m_hWnd, IDS_HELPTEXTNOPROPS);
			return true;
		}

		break;
	}

	return CSortListCtrl::WindowProc(message, wParam, lParam);
}

// WM_MFCMAPI_SAVECOLUMNORDERLIST
_Check_return_ LRESULT CSingleMAPIPropListCtrl::msgOnSaveColumnOrder(WPARAM /*wParam*/, LPARAM /*lParam*/)
{
	auto hRes = S_OK;
	auto lpMyHeader = GetHeaderCtrl();

	if (lpMyHeader)
	{
		ULONG nColumnCount = lpMyHeader->GetItemCount();
		if (nColumnCount && nColumnCount <= MAX_SORT_COLS)
		{
			auto pnOrder = new int[nColumnCount];

			if (pnOrder)
			{
				RegKeys[regkeyPROP_COLUMN_ORDER].szCurSTRING.clear();
				EC_B(GetColumnOrderArray(pnOrder, nColumnCount));
				for (ULONG i = 0; i < nColumnCount; i++)
				{
					RegKeys[regkeyPROP_COLUMN_ORDER].szCurSTRING.push_back(static_cast<wchar_t>(L'a' + pnOrder[i]));
				}
			}

			delete[] pnOrder;
		}
	}
	return S_OK;
}

void CSingleMAPIPropListCtrl::InitMenu(_In_ CMenu* pMenu) const
{
	if (pMenu)
	{
		ULONG ulPropTag = NULL;
		auto bHasSource = m_lpPropBag != nullptr;

		GetSelectedPropTag(&ulPropTag);
		auto bPropSelected = NULL != ulPropTag;

		auto ulStatus = CGlobalCache::getInstance().GetBufferStatus();
		auto lpEIDsToCopy = CGlobalCache::getInstance().GetMessagesToCopy();
		pMenu->EnableMenuItem(ID_PASTE_PROPERTY, DIM(bHasSource && (ulStatus & BUFFER_PROPTAG) && (ulStatus & BUFFER_SOURCEPROPOBJ)));
		pMenu->EnableMenuItem(ID_COPYTO, DIM(bHasSource && (ulStatus & BUFFER_SOURCEPROPOBJ)));
		pMenu->EnableMenuItem(ID_PASTE_NAMEDPROPS, DIM(bHasSource && (ulStatus & BUFFER_MESSAGES) && lpEIDsToCopy && 1 == lpEIDsToCopy->cValues));

		pMenu->EnableMenuItem(ID_COPY_PROPERTY, DIM(bHasSource));

		pMenu->EnableMenuItem(ID_DELETEPROPERTY, DIM(bHasSource && bPropSelected));
		pMenu->EnableMenuItem(ID_DISPLAYPROPERTYASSECURITYDESCRIPTORPROPSHEET, DIM(bHasSource && bPropSelected && pfnEditSecurity));
		pMenu->EnableMenuItem(ID_EDITPROPASBINARYSTREAM, DIM(bHasSource && bPropSelected));
		pMenu->EnableMenuItem(ID_EDITPROPERTY, DIM(bPropSelected));
		pMenu->EnableMenuItem(ID_EDITPROPERTYASASCIISTREAM, DIM(bHasSource && bPropSelected));
		pMenu->EnableMenuItem(ID_EDITPROPERTYASUNICODESTREAM, DIM(bHasSource && bPropSelected));
		pMenu->EnableMenuItem(ID_EDITPROPERTYASPRRTFCOMPRESSEDSTREAM, DIM(bHasSource && bPropSelected));
		pMenu->EnableMenuItem(ID_OPEN_PROPERTY, DIM(bPropSelected));

		pMenu->EnableMenuItem(ID_SAVEPROPERTIES, DIM(bHasSource));
		pMenu->EnableMenuItem(ID_EDITGIVENPROPERTY, DIM(bHasSource));
		pMenu->EnableMenuItem(ID_OPENPROPERTYASTABLE, DIM(bHasSource));
		pMenu->EnableMenuItem(ID_FINDALLNAMEDPROPS, DIM(bHasSource));
		pMenu->EnableMenuItem(ID_COUNTNAMEDPROPS, DIM(bHasSource));

		if (m_lpHostDlg)
		{
			for (ULONG ulMenu = ID_ADDINPROPERTYMENU;; ulMenu++)
			{
				auto lpAddInMenu = GetAddinMenuItem(m_lpHostDlg->m_hWnd, ulMenu);
				if (!lpAddInMenu) break;

				UINT uiEnable = DIM(bPropSelected);

				pMenu->EnableMenuItem(ulMenu, uiEnable);
			}
		}

		if (!m_lpPropBag || m_lpPropBag->GetType() == pbRow)
		{
			pMenu->EnableMenuItem(ID_DELETEPROPERTY, DIM(false));
		}
	}
}

_Check_return_ bool CSingleMAPIPropListCtrl::HandleMenu(WORD wMenuSelect)
{
	DebugPrint(DBGMenu, L"CSingleMAPIPropListCtrl::HandleMenu wMenuSelect = 0x%X = %u\n", wMenuSelect, wMenuSelect);
	switch (wMenuSelect)
	{
	case ID_COPY_PROPERTY: OnCopyProperty(); return true;
	case ID_COPYTO: OnCopyTo(); return true;
	case ID_DELETEPROPERTY: OnDeleteProperty(); return true;
	case ID_DISPLAYPROPERTYASSECURITYDESCRIPTORPROPSHEET: OnDisplayPropertyAsSecurityDescriptorPropSheet(); return true;
	case ID_EDITGIVENPROPERTY: OnEditGivenProperty(); return true;
	case ID_EDITPROPERTY: OnEditProp(); return true;
	case ID_EDITPROPASBINARYSTREAM: OnEditPropAsStream(PT_BINARY, false); return true;
	case ID_EDITPROPERTYASASCIISTREAM: OnEditPropAsStream(PT_STRING8, false); return true;
	case ID_EDITPROPERTYASUNICODESTREAM: OnEditPropAsStream(PT_UNICODE, false); return true;
	case ID_EDITPROPERTYASPRRTFCOMPRESSEDSTREAM: OnEditPropAsStream(PT_BINARY, true); return true;
	case ID_FINDALLNAMEDPROPS: FindAllNamedProps(); return true;
	case ID_COUNTNAMEDPROPS: CountNamedProps(); return true;
	case ID_MODIFYEXTRAPROPS: OnModifyExtraProps(); return true;
	case ID_OPEN_PROPERTY: OnOpenProperty(); return true;
	case ID_OPENPROPERTYASTABLE: OnOpenPropertyAsTable(); return true;
	case ID_PASTE_NAMEDPROPS: OnPasteNamedProps(); return true;
	case ID_PASTE_PROPERTY: OnPasteProperty(); return true;
	case ID_SAVEPROPERTIES: SavePropsToXML(); return true;
	}

	return HandleAddInMenu(wMenuSelect);
}

void CSingleMAPIPropListCtrl::GetSelectedPropTag(_Out_ ULONG* lpPropTag) const
{
	if (lpPropTag) *lpPropTag = NULL;

	auto iItem = GetNextItem(-1, LVNI_FOCUSED | LVNI_SELECTED);

	if (-1 != iItem)
	{
		auto lpListData = reinterpret_cast<SortListData*>(GetItemData(iItem));
		if (lpListData && lpListData->Prop() && lpPropTag)
		{
			*lpPropTag = lpListData->Prop()->m_ulPropTag;
		}
	}

	if (lpPropTag)
	{
		DebugPrintEx(DBGGeneric, CLASS, L"GetSelectedPropTag", L"returning lpPropTag = 0x%X\n", *lpPropTag);
	}
}

bool IsABPropSet(ULONG ulProps, LPSPropValue lpProps)
{
	auto lpObjectType = PpropFindProp(lpProps, ulProps, PR_OBJECT_TYPE);

	if (lpObjectType && PR_OBJECT_TYPE == lpObjectType->ulPropTag)
	{
		switch (lpObjectType->Value.l)
		{
		case MAPI_ADDRBOOK:
		case MAPI_ABCONT:
		case MAPI_MAILUSER:
		case MAPI_DISTLIST:
			return true;
		}
	}
	return false;
}

// Call GetProps with NULL to get a list of (almost) all properties.
// Parse this list and render them in the control.
// Add any extra props we've asked for through the UI
_Check_return_ HRESULT CSingleMAPIPropListCtrl::LoadMAPIPropList()
{
	auto hRes = S_OK;
	ULONG ulCurListBoxRow = 0;
	CWaitCursor Wait; // Change the mouse to an hourglass while we work.
	ULONG ulProps = 0;
	LPSPropValue lpPropsToAdd = nullptr;
	LPSPropValue lpMappingSig = nullptr;

	if (!m_lpPropBag) return MAPI_E_INVALID_PARAMETER;

	if (!RegKeys[regkeyONLY_ADDITIONAL_PROPERTIES].ulCurDWORD)
	{
		WC_H(m_lpPropBag->GetAllProps(
			&ulProps,
			&lpPropsToAdd));

		// If this is an AB object, make sure we interpret it as such
		if (IsABPropSet(ulProps, lpPropsToAdd))
		{
			m_bIsAB = true;
		}

		if (m_lpHostDlg)
		{
			// This flag may be set by a GetProps call, so we make this check AFTER we get our props
			if (pbBackedByGetProps == m_lpPropBag->GetFlags())
			{
				m_lpHostDlg->UpdateStatusBarText(STATUSINFOTEXT, IDS_PROPSFROMGETPROPS);
			}
			else
			{
				m_lpHostDlg->UpdateStatusBarText(STATUSINFOTEXT, IDS_PROPSFROMROW);
			}
		}

		// If we got some properties and PR_MAPPING_SIGNATURE is among them, grab it now
		lpMappingSig = PpropFindProp(lpPropsToAdd, ulProps, PR_MAPPING_SIGNATURE);

		// Add our props to the view
		if (lpPropsToAdd)
		{
			ULONG ulPropNames = 0;
			LPMAPINAMEID* lppPropNames = nullptr;
			ULONG ulCurTag = 0;
			if (!m_bIsAB && RegKeys[regkeyPARSED_NAMED_PROPS].ulCurDWORD)
			{
				// If we don't pass named property information to AddPropToListBox, it will look it up for us
				// But this costs a GetNamesFromIDs call for each property we add
				// As a speed up, we put together a single GetNamesFromIDs call here and pass its results to AddPropToListBox
				// The speed up in non-cached mode is enormous
				if (m_lpPropBag->GetMAPIProp())
				{
					ULONG ulNamedProps = 0;

					// First, count how many props to look up
					if (RegKeys[regkeyGETPROPNAMES_ON_ALL_PROPS].ulCurDWORD)
					{
						ulNamedProps = ulProps;
					}
					else
					{
						for (ULONG ulCurPropRow = 0; ulCurPropRow < ulProps; ulCurPropRow++)
						{
							if (PROP_ID(lpPropsToAdd[ulCurPropRow].ulPropTag) >= 0x8000) ulNamedProps++;
						}
					}

					// Allocate our tag array
					LPSPropTagArray lpTag = nullptr;
					if (ulNamedProps)
					{
						EC_H(MAPIAllocateBuffer(
							CbNewSPropTagArray(ulNamedProps),
							reinterpret_cast<LPVOID*>(&lpTag)));
						if (lpTag)
						{
							// Populate the array
							lpTag->cValues = ulNamedProps;
							if (RegKeys[regkeyGETPROPNAMES_ON_ALL_PROPS].ulCurDWORD)
							{
								for (ULONG ulCurPropRow = 0; ulCurPropRow < ulProps; ulCurPropRow++)
								{
									lpTag->aulPropTag[ulCurPropRow] = lpPropsToAdd[ulCurPropRow].ulPropTag;
								}
							}
							else
							{
								for (ULONG ulCurPropRow = 0; ulCurPropRow < ulProps; ulCurPropRow++)
								{
									if (PROP_ID(lpPropsToAdd[ulCurPropRow].ulPropTag) >= 0x8000)
									{
										lpTag->aulPropTag[ulCurTag] = lpPropsToAdd[ulCurPropRow].ulPropTag;
										ulCurTag++;
									}
								}
							}

							// Get the names
							WC_H_GETPROPS(GetNamesFromIDs(m_lpPropBag->GetMAPIProp(),
								lpMappingSig ? &lpMappingSig->Value.bin : NULL,
								&lpTag,
								NULL,
								NULL,
								&ulPropNames,
								&lppPropNames));

							MAPIFreeBuffer(lpTag);
						}
					}
				}
			}

			// Is this worth it?
			// Set the item count to speed up the addition of items
			auto ulTotalRowCount = ulProps;
			if (m_lpPropBag && m_sptExtraProps) ulTotalRowCount += m_sptExtraProps->cValues;
			SetItemCount(ulTotalRowCount);

			// get each property in turn and add it to the list
			ulCurTag = 0;
			for (ULONG ulCurPropRow = 0; ulCurPropRow < ulProps; ulCurPropRow++)
			{
				LPMAPINAMEID lpNameIDInfo = nullptr;
				// We shouldn't need to check ulCurTag < ulPropNames, but I fear bad GetNamesFromIDs implementations
				if (lppPropNames && ulCurTag < ulPropNames)
				{
					if (RegKeys[regkeyGETPROPNAMES_ON_ALL_PROPS].ulCurDWORD ||
						PROP_ID(lpPropsToAdd[ulCurPropRow].ulPropTag) >= 0x8000)
					{
						lpNameIDInfo = lppPropNames[ulCurTag];
						ulCurTag++;
					}
				}

				AddPropToListBox(
					ulCurListBoxRow,
					lpPropsToAdd[ulCurPropRow].ulPropTag,
					lpNameIDInfo,
					lpMappingSig ? &lpMappingSig->Value.bin : NULL,
					&lpPropsToAdd[ulCurPropRow]);

				ulCurListBoxRow++;
			}
			MAPIFreeBuffer(lppPropNames);
		}
	}

	// Now check if the user has given us any other properties to add and get them one at a time
	if (m_sptExtraProps)
	{
		// Let's get each extra property one at a time
		ULONG cExtraProps = 0;
		LPSPropValue pExtraProps = nullptr;
		SPropValue ExtraPropForList;
		SPropTagArray pNewTag;
		pNewTag.cValues = 1;

		for (ULONG iCurExtraProp = 0; iCurExtraProp < m_sptExtraProps->cValues; iCurExtraProp++)
		{
			hRes = S_OK; // clear the error flag before each run
			pNewTag.aulPropTag[0] = m_sptExtraProps->aulPropTag[iCurExtraProp];

			// Let's add some extra properties
			// Don't need to report since we're gonna put show the error in the UI
			WC_H(m_lpPropBag->GetProps(
				&pNewTag,
				fMapiUnicode,
				&cExtraProps,
				&pExtraProps));

			if (pExtraProps)
			{
				ExtraPropForList.dwAlignPad = pExtraProps[0].dwAlignPad;

				if (PROP_TYPE(pNewTag.aulPropTag[0]) == NULL)
				{
					// In this case, we started with a NULL tag, but we got a property back - let's 'fix' our tag for the UI
					pNewTag.aulPropTag[0] =
						CHANGE_PROP_TYPE(pNewTag.aulPropTag[0], PROP_TYPE(pExtraProps[0].ulPropTag));
				}

				// We want to give our parser the tag that came back from GetProps
				ExtraPropForList.ulPropTag = pExtraProps[0].ulPropTag;

				ExtraPropForList.Value = pExtraProps[0].Value;
			}
			else
			{
				ExtraPropForList.dwAlignPad = NULL;
				ExtraPropForList.ulPropTag = CHANGE_PROP_TYPE(pNewTag.aulPropTag[0], PT_ERROR);
				ExtraPropForList.Value.err = hRes;
			}

			// Add the property to the list
			AddPropToListBox(
				ulCurListBoxRow,
				pNewTag.aulPropTag[0], // Tag to use in the UI
				nullptr, // Let AddPropToListBox look up any named prop information it needs
				lpMappingSig ? &lpMappingSig->Value.bin : NULL,
				&ExtraPropForList); // Tag + Value to parse - may differ in case of errors or NULL type.

			ulCurListBoxRow++;

			MAPIFreeBuffer(pExtraProps);
			pExtraProps = nullptr;
		}
	}

	// lpMappingSig might come from lpPropsToAdd, so don't free this until here
	m_lpPropBag->FreeBuffer(lpPropsToAdd);

	DebugPrintEx(DBGGeneric, CLASS, L"LoadMAPIPropList", L"added %u properties\n", ulCurListBoxRow);

	SortClickedColumn();

	// Don't report any errors from here - don't care at this point
	return S_OK;
}

_Check_return_ HRESULT CSingleMAPIPropListCtrl::RefreshMAPIPropList()
{
	auto hRes = S_OK;
	int iSelectedItem;
	DebugPrintEx(DBGGeneric, CLASS, L"RefreshMAPIPropList", L"\n");

	// Turn off redraw while we work on the window
	MySetRedraw(false);
	auto MyPos = GetFirstSelectedItemPosition();

	iSelectedItem = GetNextSelectedItem(MyPos);

	EC_B(DeleteAllItems());

	if (m_lpPropBag)
		EC_H(LoadMAPIPropList());

	SetItemState(iSelectedItem, LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED);

	EnsureVisible(iSelectedItem, false);

	// Turn redraw back on to update our view
	MySetRedraw(true);

	if (m_lpHostDlg)
		m_lpHostDlg->UpdateStatusBarText(STATUSDATA2, IDS_STATUSTEXTNUMPROPS, GetItemCount());

	return hRes;
}

_Check_return_ HRESULT CSingleMAPIPropListCtrl::AddPropToExtraProps(ULONG ulPropTag, bool bRefresh)
{
	auto hRes = S_OK;
	SPropTagArray sptSingleProp;

	DebugPrintEx(DBGGeneric, CLASS, L"AddPropToExtraProps", L"adding proptag 0x%X\n", ulPropTag);

	// Cache this proptag so we continue to request it in this view
	// We've got code to refresh any props cached in m_sptExtraProps...let's add to that.

	sptSingleProp.cValues = 1;
	sptSingleProp.aulPropTag[0] = ulPropTag;

	EC_H(AddPropsToExtraProps(&sptSingleProp, bRefresh));

	return hRes;
}

_Check_return_ HRESULT CSingleMAPIPropListCtrl::AddPropsToExtraProps(_In_ LPSPropTagArray lpPropsToAdd, bool bRefresh)
{
	auto hRes = S_OK;
	LPSPropTagArray lpNewExtraProps = nullptr;

	DebugPrintEx(DBGGeneric, CLASS, L"AddPropsToExtraProps", L"adding prop array %p\n", lpPropsToAdd);

	EC_H(ConcatSPropTagArrays(
		m_sptExtraProps,
		lpPropsToAdd,
		&lpNewExtraProps));

	MAPIFreeBuffer(m_sptExtraProps);
	m_sptExtraProps = lpNewExtraProps;

	if (bRefresh)
	{
		WC_H(RefreshMAPIPropList());
	}

	return hRes;
}

#define NUMPROPTYPES 31
static ULONG _PropTypeIcons[NUMPROPTYPES][2] =
{
 { PT_UNSPECIFIED, slIconUNSPECIFIED },
 { PT_NULL, slIconNULL },
 { PT_I2, slIconI2 },
 { PT_LONG, slIconLONG },
 { PT_R4, slIconR4 },
 { PT_DOUBLE, slIconDOUBLE },
 { PT_CURRENCY, slIconCURRENCY },
 { PT_APPTIME, slIconAPPTIME },
 { PT_ERROR, slIconERROR },
 { PT_BOOLEAN, slIconBOOLEAN },
 { PT_OBJECT, slIconOBJECT },
 { PT_I8, slIconI8 },
 { PT_STRING8, slIconSTRING8 },
 { PT_UNICODE, slIconUNICODE },
 { PT_SYSTIME, slIconSYSTIME },
 { PT_CLSID, slIconCLSID },
 { PT_BINARY, slIconBINARY },
 { PT_MV_I2, slIconMV_I2 },
 { PT_MV_LONG, slIconMV_LONG },
 { PT_MV_R4, slIconMV_R4 },
 { PT_MV_DOUBLE, slIconMV_DOUBLE },
 { PT_MV_CURRENCY, slIconMV_CURRENCY },
 { PT_MV_APPTIME, slIconMV_APPTIME },
 { PT_MV_SYSTIME, slIconMV_SYSTIME },
 { PT_MV_STRING8, slIconMV_STRING8 },
 { PT_MV_BINARY, slIconMV_BINARY },
 { PT_MV_UNICODE, slIconMV_UNICODE },
 { PT_MV_CLSID, slIconMV_CLSID },
 { PT_MV_I8, slIconMV_I8 },
 { PT_SRESTRICTION, slIconSRESTRICTION },
 { PT_ACTIONS, slIconACTIONS },
};

// Crack open the given SPropValue and render it to the given row in the list.
void CSingleMAPIPropListCtrl::AddPropToListBox(
	int iRow,
	ULONG ulPropTag,
	_In_opt_ LPMAPINAMEID lpNameID,
	_In_opt_ LPSBinary lpMappingSignature, // optional mapping signature for object to speed named prop lookups
	_In_ LPSPropValue lpsPropToAdd)
{
	ULONG ulImage = slIconDefault;
	if (lpsPropToAdd)
	{
		for (auto i = 0; i < NUMPROPTYPES; i++)
		{
			if (_PropTypeIcons[i][0] == PROP_TYPE(lpsPropToAdd->ulPropTag))
			{
				ulImage = _PropTypeIcons[i][1];
				break;
			}
		}
	}

	auto lpData = InsertRow(iRow, L"", 0, ulImage);
	// Data used to refer to specific property tags. See GetSelectedPropTag.
	if (lpData)
	{
		lpData->InitializePropList(ulPropTag);
	}

	auto PropTag = format(L"0x%08X", ulPropTag);
	wstring PropString;
	wstring AltPropString;
	wstring szExactMatches;
	wstring szPartialMatches;
	wstring szNamedPropName;
	wstring szNamedPropGUID;
	wstring szNamedPropDASL;

	// If we have lpNameID, we don't need to ask InterpretProp to look up named property information

	NameIDToStrings(
		ulPropTag,
		m_lpPropBag->GetMAPIProp(),
		lpNameID,
		lpMappingSignature,
		m_bIsAB,
		szNamedPropName, // Built from lpProp & lpMAPIProp
		szNamedPropGUID, // Built from lpProp & lpMAPIProp
		szNamedPropDASL);

	PropTagToPropName(ulPropTag, m_bIsAB, szExactMatches, szPartialMatches);
	if (!szExactMatches.empty())
	{
		SetItemText(iRow, pcPROPEXACTNAMES, szExactMatches);
	}
	else
	{
		SetItemText(iRow, pcPROPEXACTNAMES, PropTag);
	}

	SetItemText(iRow, pcPROPPARTIALNAMES, szPartialMatches);

	SetItemText(iRow, pcPROPTAG, PropTag);
	SetItemText(iRow, pcPROPTYPE, TypeToString(ulPropTag));

	InterpretProp(lpsPropToAdd, &PropString, &AltPropString);
	SetItemText(iRow, pcPROPVAL, PropString);
	SetItemText(iRow, pcPROPVALALT, AltPropString);

	auto szSmartView = InterpretPropSmartView(
		lpsPropToAdd,
		m_lpPropBag->GetMAPIProp(),
		lpNameID,
		lpMappingSignature,
		m_bIsAB,
		false); // Built from lpProp & lpMAPIProp
	if (!szSmartView.empty()) SetItemText(iRow, pcPROPSMARTVIEW, szSmartView);
	if (!szNamedPropName.empty()) SetItemText(iRow, pcPROPNAMEDNAME, szNamedPropName);
	if (!szNamedPropGUID.empty()) SetItemText(iRow, pcPROPNAMEDIID, szNamedPropGUID);
}

_Check_return_ HRESULT CSingleMAPIPropListCtrl::GetDisplayedProps(ULONG FAR* lpcValues, LPSPropValue FAR* lppPropArray) const
{
	if (!m_lpPropBag) return MAPI_E_INVALID_PARAMETER;

	return m_lpPropBag->GetAllProps(lpcValues, lppPropArray);
}

_Check_return_ bool CSingleMAPIPropListCtrl::IsModifiedPropVals() const
{
	if (pbModified == (m_lpPropBag->GetFlags() & pbModified)) return true;
	return false;
}

_Check_return_ HRESULT CSingleMAPIPropListCtrl::SetDataSource(_In_opt_ LPMAPIPROP lpMAPIProp, _In_opt_ SortListData* lpListData, bool bIsAB)
{
	DebugPrintEx(DBGGeneric, CLASS, L"SetDataSource", L"setting new data source\n");

	LPMAPIPROPERTYBAG lpNewPropBag = nullptr;
	if (lpMAPIProp)
	{
		lpNewPropBag = new MAPIPropPropertyBag(lpMAPIProp, lpListData);
	}
	else if (lpListData)
	{
		lpNewPropBag = new RowPropertyBag(lpListData);
	}

	return SetDataSource(lpNewPropBag, bIsAB);
}
// Clear the current property list from the control.
// Load a new list from the IMAPIProp or lpSourceProps object passed in
// Most calls to this will come through CBaseDialog::OnUpdateSingleMAPIPropListCtrl, which will preserve the current bIsAB
// Exceptions will be where we need to set a specific bIsAB
_Check_return_ HRESULT CSingleMAPIPropListCtrl::SetDataSource(_In_opt_ LPMAPIPROPERTYBAG lpPropBag, bool bIsAB)
{
	auto hRes = S_OK;
	DebugPrintEx(DBGGeneric, CLASS, L"SetDataSource", L"setting new data source\n");

	// if nothing to do...do nothing
	if (lpPropBag && lpPropBag->IsEqual(m_lpPropBag))
	{
		delete lpPropBag;
		return S_OK;
	}

	if (m_lpPropBag) delete m_lpPropBag;
	m_lpPropBag = lpPropBag;
	m_bIsAB = bIsAB;

	// Turn off redraw while we work on the window
	MySetRedraw(false);

	WC_H(RefreshMAPIPropList());

	// Reset our header widths if weren't showing anything before and are now
	if (S_OK == hRes && !m_bHaveEverDisplayedSomething && m_lpPropBag && GetItemCount())
	{
		m_bHaveEverDisplayedSomething = true;

		auto lpMyHeader = GetHeaderCtrl();

		if (lpMyHeader)
		{
			// This fixes a ton of flashing problems
			lpMyHeader->SetRedraw(true);
			for (auto  iCurCol = 0; iCurCol < int(PropColumns.size()); iCurCol++)
			{
				SetColumnWidth(iCurCol, LVSCW_AUTOSIZE_USEHEADER);
				if (GetColumnWidth(iCurCol) > 200) SetColumnWidth(iCurCol, 200);
			}

			lpMyHeader->SetRedraw(false);
		}
	}

	// Turn redraw back on to update our view
	MySetRedraw(true);
	return hRes;
}

void CSingleMAPIPropListCtrl::SavePropsToXML()
{
	auto szFileName = CFileDialogExW::SaveAs(
		L"xml", // STRING_OK
		L"props.xml", // STRING_OK
		OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT,
		loadstring(IDS_XMLFILES),
		this);
	if (!szFileName.empty())
	{
		auto fProps = MyOpenFile(szFileName, true);
		if (fProps)
		{
			DebugPrintEx(DBGGeneric, CLASS, L"SavePropsToXML", L"saving to %ws\n", szFileName.c_str());

			// Force a sort on the tag column to make output consistent
			FakeClickColumn(pcPROPTAG, false);

			int iRow;
			auto iItemCount = GetItemCount();

			OutputToFile(fProps, g_szXMLHeader);
			OutputToFile(fProps, L"<propertypane>\n");
			for (iRow = 0; iRow < iItemCount; iRow++)
			{
				auto szTemp1 = GetItemText(iRow, pcPROPTAG);
				auto szTemp2 = GetItemText(iRow, pcPROPTYPE);
				OutputToFilef(fProps, L"\t<property tag = \"%ws\" type = \"%ws\">\n", szTemp1.c_str(), szTemp2.c_str());

				szTemp1 = GetItemText(iRow, pcPROPEXACTNAMES);
				OutputXMLValueToFile(fProps, PropXMLNames[pcPROPEXACTNAMES].uidName, szTemp1, false, 2);

				szTemp1 = GetItemText(iRow, pcPROPPARTIALNAMES);
				OutputXMLValueToFile(fProps, PropXMLNames[pcPROPPARTIALNAMES].uidName, szTemp1, false, 2);

				szTemp1 = GetItemText(iRow, pcPROPNAMEDIID);
				OutputXMLValueToFile(fProps, PropXMLNames[pcPROPNAMEDIID].uidName, szTemp1, false, 2);

				szTemp1 = GetItemText(iRow, pcPROPNAMEDNAME);
				OutputXMLValueToFile(fProps, PropXMLNames[pcPROPNAMEDNAME].uidName, szTemp1, false, 2);

				auto lpListData = reinterpret_cast<SortListData*>(GetItemData(iRow));
				ULONG ulPropType = PT_NULL;
				if (lpListData && lpListData->Prop())
				{
					ulPropType = PROP_TYPE(lpListData->Prop()->m_ulPropTag);
				}

				szTemp1 = GetItemText(iRow, pcPROPVAL);
				szTemp2 = GetItemText(iRow, pcPROPVALALT);
				switch (ulPropType)
				{
				case PT_STRING8:
				case PT_UNICODE:
					OutputXMLValueToFile(fProps, PropXMLNames[pcPROPVAL].uidName, szTemp1, true, 2);
					OutputXMLValueToFile(fProps, PropXMLNames[pcPROPVALALT].uidName, szTemp2, false, 2);
					break;
				case PT_BINARY:
					OutputXMLValueToFile(fProps, PropXMLNames[pcPROPVAL].uidName, szTemp1, false, 2);
					OutputXMLValueToFile(fProps, PropXMLNames[pcPROPVALALT].uidName, szTemp2, true, 2);
					break;
				default:
					OutputXMLValueToFile(fProps, PropXMLNames[pcPROPVAL].uidName, szTemp1, false, 2);
					OutputXMLValueToFile(fProps, PropXMLNames[pcPROPVALALT].uidName, szTemp2, false, 2);
					break;
				}

				szTemp1 = GetItemText(iRow, pcPROPSMARTVIEW);
				OutputXMLValueToFile(fProps, PropXMLNames[pcPROPSMARTVIEW].uidName, szTemp1, true, 2);

				OutputToFile(fProps, L"\t</property>\n");
			}

			OutputToFile(fProps, L"</propertypane>");
			CloseFile(fProps);
		}
	}
}

void CSingleMAPIPropListCtrl::OnDblclk(_In_ NMHDR* /*pNMHDR*/, _In_ LRESULT* pResult)
{
	DebugPrintEx(DBGGeneric, CLASS, L"OnDblclk", L"calling OnEditProp\n");
	OnEditProp();
	*pResult = 0;
}

void CSingleMAPIPropListCtrl::OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags)
{
	DebugPrintEx(DBGMenu, CLASS, L"OnKeyDown", L"0x%X\n", nChar);

	auto hRes = S_OK;
	auto bCtrlPressed = GetKeyState(VK_CONTROL) < 0;
	auto bShiftPressed = GetKeyState(VK_SHIFT) < 0;
	auto bMenuPressed = GetKeyState(VK_MENU) < 0;

	if (!bMenuPressed)
	{
		if ('X' == nChar && bCtrlPressed && !bShiftPressed)
		{
			OnDeleteProperty();
		}
		else if (VK_DELETE == nChar)
		{
			OnDeleteProperty();
		}
		else if ('S' == nChar && bCtrlPressed)
		{
			SavePropsToXML();
		}
		else if ('E' == nChar && bCtrlPressed)
		{
			OnEditProp();
		}
		else if ('C' == nChar && bCtrlPressed && !bShiftPressed)
		{
			OnCopyProperty();
		}
		else if ('V' == nChar && bCtrlPressed && !bShiftPressed)
		{
			OnPasteProperty();
		}
		else if (VK_F5 == nChar)
		{
			WC_H(RefreshMAPIPropList());
		}
		else if (VK_RETURN == nChar)
		{
			if (!bCtrlPressed)
			{
				DebugPrintEx(DBGMenu, CLASS, L"OnKeyDown", L"calling OnEditProp\n");
				OnEditProp();
			}
			else
			{
				DebugPrintEx(DBGMenu, CLASS, L"OnKeyDown", L"calling OnOpenProperty\n");
				OnOpenProperty();
			}
		}
		else if (!m_lpHostDlg || !m_lpHostDlg->HandleKeyDown(nChar, bShiftPressed, bCtrlPressed, bMenuPressed))
		{
			CSortListCtrl::OnKeyDown(nChar, nRepCnt, nFlags);
		}
	}
}

void CSingleMAPIPropListCtrl::OnContextMenu(_In_ CWnd* pWnd, CPoint pos)
{
	if (pWnd && -1 == pos.x && -1 == pos.y)
	{
		POINT point = { 0 };
		auto iItem = GetNextItem(
			-1,
			LVNI_SELECTED);
		GetItemPosition(iItem, &point);
		::ClientToScreen(pWnd->m_hWnd, &point);
		pos = point;
	}
	DisplayContextMenu(IDR_MENU_PROPERTY_POPUP, NULL, m_lpHostDlg->m_hWnd, pos.x, pos.y);
}

void CSingleMAPIPropListCtrl::FindAllNamedProps()
{
	auto hRes = S_OK;
	if (!m_lpPropBag) return;

	// Exchange can return MAPI_E_NOT_ENOUGH_MEMORY when I call this - give it a try - PSTs support it
	DebugPrintEx(DBGNamedProp, CLASS, L"FindAllNamedProps", L"Calling GetIDsFromNames with a NULL\n");
	LPSPropTagArray lptag = nullptr;
	WC_H(GetIDsFromNames(m_lpPropBag->GetMAPIProp(),
		NULL,
		NULL,
		NULL,
		&lptag));
	if (S_OK == hRes && lptag && lptag->cValues)
	{
		// Now we have an array of tags - add them in:
		EC_H(AddPropsToExtraProps(lptag, false));
		MAPIFreeBuffer(lptag);
		lptag = nullptr;
	}
	else
	{
		hRes = S_OK;
		DebugPrintEx(DBGNamedProp, CLASS, L"FindAllNamedProps", L"Exchange didn't support GetIDsFromNames(NULL).\n");

#define __LOWERBOUND 0x8000
#define __UPPERBOUNDDEFAULT 0x8FFF
#define __UPPERBOUND 0xFFFF

		CEditor MyData(
			this,
			IDS_FINDNAMEPROPSLIMIT,
			IDS_FINDNAMEPROPSLIMITPROMPT,
			CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
		MyData.InitPane(0, TextPane::CreateSingleLinePane(IDS_LOWERBOUND, false));
		MyData.SetHex(0, __LOWERBOUND);
		MyData.InitPane(1, TextPane::CreateSingleLinePane(IDS_UPPERBOUND, false));
		MyData.SetHex(1, __UPPERBOUNDDEFAULT);

		WC_H(MyData.DisplayDialog());

		if (S_OK == hRes)
		{
			auto ulLowerBound = MyData.GetHex(0);
			auto ulUpperBound = MyData.GetHex(1);

			DebugPrintEx(DBGNamedProp, CLASS, L"FindAllNamedProps", L"Walking through all IDs from 0x%X to 0x%X, looking for mappings to names\n", ulLowerBound, ulUpperBound);
			if (ulLowerBound < __LOWERBOUND)
			{
				ErrDialog(__FILE__, __LINE__, IDS_EDLOWERBOUNDTOOLOW, ulLowerBound, __LOWERBOUND);
				hRes = MAPI_E_INVALID_PARAMETER;
			}
			else if (ulUpperBound > __UPPERBOUND)
			{
				ErrDialog(__FILE__, __LINE__, IDS_EDUPPERBOUNDTOOHIGH, ulUpperBound, __UPPERBOUND);
				hRes = MAPI_E_INVALID_PARAMETER;
			}
			else if (ulLowerBound > ulUpperBound)
			{
				ErrDialog(__FILE__, __LINE__, IDS_EDLOWEROVERUPPER, ulLowerBound, ulUpperBound);
				hRes = MAPI_E_INVALID_PARAMETER;
			}
			else
			{
				SPropTagArray tag = { 0 };
				tag.cValues = 1;
				lptag = &tag;
				for (auto iTag = ulLowerBound; iTag <= ulUpperBound; iTag++)
				{
					LPMAPINAMEID* lppPropNames = nullptr;
					ULONG ulPropNames = 0;
					hRes = S_OK;
					tag.aulPropTag[0] = PROP_TAG(NULL, iTag);

					WC_H(GetNamesFromIDs(m_lpPropBag->GetMAPIProp(),
						&lptag,
						NULL,
						NULL,
						&ulPropNames,
						&lppPropNames));
					if (S_OK == hRes && ulPropNames == 1 && lppPropNames && *lppPropNames)
					{
						DebugPrintEx(DBGNamedProp, CLASS, L"FindAllNamedProps", L"Found an ID with a name (0x%X). Adding to extra prop list.\n", iTag);
						EC_H(AddPropToExtraProps(PROP_TAG(NULL, iTag), false));
					}
					MAPIFreeBuffer(lppPropNames);
					lppPropNames = nullptr;
				}
			}
		}
	}

	// Refresh the display
	WC_H(RefreshMAPIPropList());
}

void CSingleMAPIPropListCtrl::CountNamedProps()
{
	if (!m_lpPropBag) return;

	DebugPrintEx(DBGNamedProp, CLASS, L"CountNamedProps", L"Searching for the highest named prop mapping\n");

	auto hRes = S_OK;
	ULONG ulLower = 0x8000;
	ULONG ulUpper = 0xFFFF;
	ULONG ulHighestKnown = 0;
	auto ulCurrent = (ulUpper + ulLower) / 2;

	SPropTagArray tag = { 0 };
	LPMAPINAMEID* lppPropNames = nullptr;
	ULONG ulPropNames = 0;
	auto lptag = &tag;
	tag.cValues = 1;

	while (ulUpper - ulLower > 1)
	{
		hRes = S_OK;
		tag.aulPropTag[0] = PROP_TAG(NULL, ulCurrent);

		WC_H(GetNamesFromIDs(m_lpPropBag->GetMAPIProp(),
			&lptag,
			NULL,
			NULL,
			&ulPropNames,
			&lppPropNames));
		if (S_OK == hRes && ulPropNames == 1 && lppPropNames && *lppPropNames)
		{
			// Found a named property, reset lower bound

			// Avoid NameIDToStrings call if we're not debug printing
			if (fIsSet(DBGNamedProp))
			{
				DebugPrintEx(DBGNamedProp, CLASS, L"CountNamedProps", L"Found a named property at 0x%04X.\n", ulCurrent);
				wstring szNamedPropName;
				wstring szNamedPropGUID;
				wstring szNamedPropDASL;
				NameIDToStrings(
					tag.aulPropTag[0],
					nullptr,
					lppPropNames[0],
					nullptr,
					false,
					szNamedPropName,
					szNamedPropGUID,
					szNamedPropDASL);
				DebugPrintEx(DBGNamedProp, CLASS, L"CountNamedProps", L"Name = %ws, GUID = %ws\n", szNamedPropName.c_str(), szNamedPropGUID.c_str());
			}

			ulHighestKnown = ulCurrent;
			ulLower = ulCurrent;
		}
		else
		{
			// Did not find a named property, reset upper bound
			ulUpper = ulCurrent;
		}

		MAPIFreeBuffer(lppPropNames);
		lppPropNames = nullptr;

		ulCurrent = (ulUpper + ulLower) / 2;
	}

	CEditor MyResult(
		this,
		IDS_COUNTNAMEDPROPS,
		IDS_COUNTNAMEDPROPSPROMPT,
		CEDITOR_BUTTON_OK);
	if (ulHighestKnown)
	{
		tag.aulPropTag[0] = PROP_TAG(NULL, ulHighestKnown);

		hRes = S_OK;
		WC_H(GetNamesFromIDs(m_lpPropBag->GetMAPIProp(),
			&lptag,
			NULL,
			NULL,
			&ulPropNames,
			&lppPropNames));
		if (S_OK == hRes && ulPropNames == 1 && lppPropNames && *lppPropNames)
		{
			DebugPrintEx(DBGNamedProp, CLASS, L"CountNamedProps", L"Found a named property at 0x%04X.\n", ulCurrent);
			ulHighestKnown = ulCurrent;
		}

		MyResult.InitPane(0, TextPane::CreateSingleLinePane(IDS_HIGHESTNAMEDPROPTOTAL, true));
		MyResult.SetDecimal(0, ulHighestKnown - 0x8000);

		MyResult.InitPane(1, TextPane::CreateMultiLinePane(IDS_HIGHESTNAMEDPROPNUM, true));

		if (S_OK == hRes && ulPropNames == 1 && lppPropNames && *lppPropNames)
		{
			wstring szNamedPropName;
			wstring szNamedPropGUID;
			wstring szNamedPropDASL;
			NameIDToStrings(
				tag.aulPropTag[0],
				nullptr,
				lppPropNames[0],
				nullptr,
				false,
				szNamedPropName,
				szNamedPropGUID,
				szNamedPropDASL);
			MyResult.SetStringW(1, formatmessage(IDS_HIGHESTNAMEDPROPNAME, ulHighestKnown, szNamedPropName.c_str(), szNamedPropGUID.c_str()));

			MAPIFreeBuffer(lppPropNames);
			lppPropNames = nullptr;
		}
	}
	else
	{
		MyResult.InitPane(0, TextPane::CreateSingleLinePane(IDS_HIGHESTNAMEDPROPTOTAL, true));
		MyResult.LoadString(0, IDS_HIGHESTNAMEDPROPNOTFOUND);
	}

	hRes = S_OK;
	WC_H(MyResult.DisplayDialog());
}

// Delete the selected property
void CSingleMAPIPropListCtrl::OnDeleteProperty()
{
	auto hRes = S_OK;
	ULONG ulPropTag = NULL;

	if (!m_lpPropBag || m_lpPropBag->GetType() == pbRow) return;

	GetSelectedPropTag(&ulPropTag);
	if (!ulPropTag) return;

	CEditor Query(
		this,
		IDS_DELETEPROPERTY,
		IDS_DELETEPROPERTYPROMPT,
		CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
	WC_H(Query.DisplayDialog());
	if (S_OK == hRes)
	{
		DebugPrintEx(DBGGeneric, CLASS, L"OnDeleteProperty", L"deleting property 0x%08X\n", ulPropTag);

		EC_H(m_lpPropBag->DeleteProp(ulPropTag));

		// Refresh the display
		WC_H(RefreshMAPIPropList());
	}
}

// Display the selected property as a security dscriptor using a property sheet
void CSingleMAPIPropListCtrl::OnDisplayPropertyAsSecurityDescriptorPropSheet() const
{
	auto hRes = S_OK;
	ULONG ulPropTag = NULL;

	if (!m_lpPropBag || !pfnEditSecurity) return;

	GetSelectedPropTag(&ulPropTag);
	if (!ulPropTag) return;

	DebugPrintEx(DBGGeneric, CLASS, L"OnDisplayPropertyAsSecurityDescriptorPropSheet", L"interpreting 0x%X as Security Descriptor\n", ulPropTag);

	auto MySecInfo = new CMySecInfo(m_lpPropBag->GetMAPIProp(), ulPropTag);

	if (MySecInfo)
	{
		EC_B(pfnEditSecurity(m_hWnd, MySecInfo));

		MySecInfo->Release();
	}
}

void CSingleMAPIPropListCtrl::OnEditProp()
{
	ULONG ulPropTag = NULL;

	if (!m_lpPropBag) return;

	GetSelectedPropTag(&ulPropTag);
	if (!ulPropTag) return;

	OnEditGivenProp(ulPropTag);
}

void CSingleMAPIPropListCtrl::OnEditPropAsRestriction(ULONG ulPropTag)
{
	auto hRes = S_OK;

	if (!m_lpPropBag || !ulPropTag || PT_SRESTRICTION != PROP_TYPE(ulPropTag)) return;

	LPSPropValue lpEditProp = nullptr;
	WC_H(m_lpPropBag->GetProp(ulPropTag, &lpEditProp));
	hRes = S_OK;

	LPSRestriction lpResIn = nullptr;
	if (lpEditProp)
	{
		lpResIn = reinterpret_cast<LPSRestriction>(lpEditProp->Value.lpszA);
	}

	DebugPrint(DBGGeneric, L"Source restriction before editing:\n");
	DebugPrintRestriction(DBGGeneric, lpResIn, m_lpPropBag->GetMAPIProp());
	CRestrictEditor MyResEditor(
		this,
		nullptr, // No alloc parent - we must MAPIFreeBuffer the result
		lpResIn);
	WC_H(MyResEditor.DisplayDialog());

	if (S_OK == hRes)
	{
		auto lpModRes = MyResEditor.DetachModifiedSRestriction();
		if (lpModRes)
		{
			DebugPrint(DBGGeneric, L"Modified restriction:\n");
			DebugPrintRestriction(DBGGeneric, lpModRes, m_lpPropBag->GetMAPIProp());

			// need to merge the data we got back from the CRestrictEditor with our current prop set
			// so that we can free lpModRes
			SPropValue ResProp = { 0 };
			ResProp.ulPropTag = lpEditProp->ulPropTag;
			ResProp.Value.lpszA = reinterpret_cast<LPSTR>(lpModRes);

			EC_H(m_lpPropBag->SetProp(&ResProp));

			// Remember, we had no alloc parent - this is safe to free
			MAPIFreeBuffer(lpModRes);

			// refresh
			WC_H(RefreshMAPIPropList());
		}
	}

	m_lpPropBag->FreeBuffer(lpEditProp);
}

void CSingleMAPIPropListCtrl::OnEditGivenProp(ULONG ulPropTag)
{
	auto hRes = S_OK;
	LPSPropValue lpEditProp = nullptr;

	if (!m_lpPropBag) return;

	// Explicit check since TagToString is expensive
	if (fIsSet(DBGGeneric))
	{
		DebugPrintEx(DBGGeneric, CLASS, L"OnEditGivenProp", L"editing property 0x%X (= %ws)\n",
			ulPropTag,
			TagToString(ulPropTag, m_lpPropBag->GetMAPIProp(),
				m_bIsAB,
				true).c_str());
	}

	ulPropTag = PT_ERROR == PROP_TYPE(ulPropTag) ? CHANGE_PROP_TYPE(ulPropTag, PT_UNSPECIFIED) : ulPropTag;

	if (PT_SRESTRICTION == PROP_TYPE(ulPropTag))
	{
		OnEditPropAsRestriction(ulPropTag);
		return;
	}

	if (PT_OBJECT == PROP_TYPE(ulPropTag))
	{
		EC_H(DisplayTable(
			m_lpPropBag->GetMAPIProp(),
			ulPropTag,
			otDefault,
			m_lpHostDlg));
		return;
	}

	auto lpSourceObj = m_lpPropBag->GetMAPIProp();

	auto bUseStream = false;

	if (PROP_ID(PR_RTF_COMPRESSED) == PROP_ID(ulPropTag))
	{
		bUseStream = true;
	}
	else
	{
		WC_H(m_lpPropBag->GetProp(ulPropTag, &lpEditProp));
	}

	if (MAPI_E_NOT_ENOUGH_MEMORY == hRes) bUseStream = true;

	if (bUseStream)
	{
		hRes = S_OK;
		CStreamEditor MyEditor(
			this,
			IDS_PROPEDITOR,
			IDS_STREAMEDITORPROMPT,
			lpSourceObj,
			ulPropTag,
			true, // Guess the type of stream to use
			m_bIsAB,
			false,
			false,
			NULL,
			NULL,
			NULL);

		WC_H(MyEditor.DisplayDialog());
	}
	else
	{
		if (lpEditProp) ulPropTag = lpEditProp->ulPropTag;

		hRes = S_OK;
		LPSPropValue lpModProp = nullptr;
		WC_H(DisplayPropertyEditor(
			this,
			IDS_PROPEDITOR,
			NULL,
			m_bIsAB,
			NULL,
			lpSourceObj,
			ulPropTag,
			false,
			lpEditProp,
			lpSourceObj ? NULL : &lpModProp));

		// If we had didn't have a source object, we need to shove our results back in to the property bag
		if (S_OK == hRes && !lpSourceObj && lpModProp)
		{
			EC_H(m_lpPropBag->SetProp(lpModProp));
			// At this point, we're done with lpModProp - it was allocated off of lpSourceArray
			// and freed when a new source array was allocated. Nothing to free here. Move along.
		}

	}
	WC_H(RefreshMAPIPropList());

	m_lpPropBag->FreeBuffer(lpEditProp);
}

// Display the selected property as a stream using CStreamEditor
void CSingleMAPIPropListCtrl::OnEditPropAsStream(ULONG ulType, bool bEditAsRTF)
{
	auto hRes = S_OK;
	ULONG ulPropTag = NULL;

	if (!m_lpPropBag) return;

	GetSelectedPropTag(&ulPropTag);
	if (!ulPropTag) return;

	// Explicit check since TagToString is expensive
	if (fIsSet(DBGGeneric))
	{
		DebugPrintEx(DBGGeneric, CLASS, L"OnEditPropAsStream", L"editing property 0x%X (= %ws) as stream, ulType = 0x%08X, bEditAsRTF = 0x%X\n",
			ulPropTag,
			TagToString(ulPropTag, m_lpPropBag->GetMAPIProp(), m_bIsAB, true).c_str(),
			ulType,
			bEditAsRTF);
	}

	ulPropTag = CHANGE_PROP_TYPE(ulPropTag, ulType);

	auto bUseWrapEx = false;
	ULONG ulRTFFlags = NULL;
	ULONG ulInCodePage = NULL;
	ULONG ulOutCodePage = CP_ACP; // Default to ANSI - check if this is valid for UNICODE builds

	if (bEditAsRTF)
	{
		CEditor MyPrompt(
			this,
			IDS_USEWRAPEX,
			IDS_USEWRAPEXPROMPT,
			CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
		MyPrompt.InitPane(0, CheckPane::Create(IDS_USEWRAPEX, true, false));

		WC_H(MyPrompt.DisplayDialog());
		if (S_OK == hRes)
		{
			if (MyPrompt.GetCheck(0))
			{
				bUseWrapEx = true;
				LPSPropValue lpProp = nullptr;

				WC_H(m_lpPropBag->GetProp(
					PR_INTERNET_CPID,
					&lpProp));
				if (lpProp && PT_LONG == PROP_TYPE(lpProp[0].ulPropTag))
				{
					ulInCodePage = lpProp[0].Value.l;
				}

				m_lpPropBag->FreeBuffer(lpProp);

				CEditor MyPrompt2(
					this,
					IDS_WRAPEXFLAGS,
					IDS_WRAPEXFLAGSPROMPT,
					CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
				MyPrompt2.InitPane(0, TextPane::CreateSingleLinePane(IDS_WRAPEXFLAGS, false));
				MyPrompt2.SetHex(0, MAPI_NATIVE_BODY);
				MyPrompt2.InitPane(1, TextPane::CreateSingleLinePane(IDS_ULINCODEPAGE, false));
				MyPrompt2.SetDecimal(1, ulInCodePage);
				MyPrompt2.InitPane(2, TextPane::CreateSingleLinePane(IDS_ULOUTCODEPAGE, false));
				MyPrompt2.SetDecimal(2, 0);

				WC_H(MyPrompt2.DisplayDialog());

				if (S_OK == hRes)
				{
					ulRTFFlags = MyPrompt2.GetHex(0);
					ulInCodePage = MyPrompt2.GetDecimal(1);
					ulOutCodePage = MyPrompt2.GetDecimal(2);
				}
			}
		}
	}

	if (S_OK == hRes)
	{
		CStreamEditor MyEditor(
			this,
			IDS_PROPEDITOR,
			IDS_STREAMEDITORPROMPT,
			m_lpPropBag->GetMAPIProp(),
			ulPropTag,
			false, // No stream guessing
			m_bIsAB,
			bEditAsRTF,
			bUseWrapEx,
			ulRTFFlags,
			ulInCodePage,
			ulOutCodePage);

		WC_H(MyEditor.DisplayDialog());

		WC_H(RefreshMAPIPropList());
	}
}

void CSingleMAPIPropListCtrl::OnCopyProperty() const
{
	// for now, we only copy from objects - copying from rows would be difficult to generalize
	if (!m_lpPropBag) return;

	ULONG ulPropTag = NULL;
	GetSelectedPropTag(&ulPropTag);

	CGlobalCache::getInstance().SetPropertyToCopy(ulPropTag, m_lpPropBag->GetMAPIProp());
}

void CSingleMAPIPropListCtrl::OnPasteProperty()
{
	// For now, we only paste to objects - copying to rows would be difficult to generalize
	// TODO: Now that we have property bags, figure out how to generalize this
	if (!m_lpHostDlg || !m_lpPropBag) return;

	auto ulSourcePropTag = CGlobalCache::getInstance().GetPropertyToCopy();
	auto lpSourcePropObj = CGlobalCache::getInstance().GetSourcePropObject();
	if (!lpSourcePropObj) return;

	auto hRes = S_OK;
	LPSPropProblemArray lpProblems = nullptr;
	SPropTagArray TagArray = { 0 };
	TagArray.cValues = 1;
	TagArray.aulPropTag[0] = ulSourcePropTag;

	CEditor MyData(
		this,
		IDS_PASTEPROP,
		IDS_PASTEPROPPROMPT,
		CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);

	UINT uidDropDown[] = {
	IDS_DDCOPYPROPS,
	IDS_DDGETSETPROPS,
	IDS_DDCOPYSTREAM
	};
	MyData.InitPane(0, DropDownPane::Create(IDS_COPYSTYLE, _countof(uidDropDown), uidDropDown, true));
	MyData.InitPane(1, TextPane::CreateSingleLinePane(IDS_SOURCEPROP, false));
	MyData.SetHex(1, ulSourcePropTag);
	MyData.InitPane(2, TextPane::CreateSingleLinePane(IDS_TARGETPROP, false));
	MyData.SetHex(2, ulSourcePropTag);

	WC_H(MyData.DisplayDialog());
	if (S_OK == hRes)
	{
		auto ulSourceTag = MyData.GetHex(1);
		auto ulTargetTag = MyData.GetHex(2);
		TagArray.aulPropTag[0] = ulSourceTag;

		if (PROP_TYPE(ulTargetTag) != PROP_TYPE(ulSourceTag))
			ulTargetTag = CHANGE_PROP_TYPE(ulTargetTag, PROP_TYPE(ulSourceTag));

		switch (MyData.GetDropDown(0))
		{
		case 0:
		{
			CEditor MyCopyData(
				this,
				IDS_PASTEPROP,
				IDS_COPYPASTEPROMPT,
				CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);

			auto szGuid = GUIDToStringAndName(&IID_IMAPIProp);
			MyCopyData.InitPane(0, TextPane::CreateSingleLinePane(IDS_INTERFACE, szGuid, false));
			MyCopyData.InitPane(1, TextPane::CreateSingleLinePane(IDS_FLAGS, false));
			MyCopyData.SetHex(1, MAPI_DIALOG);

			WC_H(MyCopyData.DisplayDialog());
			if (S_OK == hRes)
			{
				auto MyGUID = StringToGUID(MyCopyData.GetStringW(0));
				LPMAPIPROGRESS lpProgress = GetMAPIProgress(L"IMAPIProp::CopyProps", m_lpHostDlg->m_hWnd); // STRING_OK
				auto ulCopyFlags = MyCopyData.GetHex(1);

				if (lpProgress)
					ulCopyFlags |= MAPI_DIALOG;

				EC_MAPI(lpSourcePropObj->CopyProps(
					&TagArray,
					lpProgress ? reinterpret_cast<ULONG_PTR>(m_lpHostDlg->m_hWnd) : NULL, // ui param
					lpProgress, // progress
					&MyGUID,
					m_lpPropBag->GetMAPIProp(),
					ulCopyFlags,
					&lpProblems));

				if (lpProgress)
					lpProgress->Release();
			}
		}
		break;
		case 1:
		{
			ULONG ulValues = NULL;
			LPSPropValue lpSourceProp = nullptr;
			EC_MAPI(lpSourcePropObj->GetProps(
				&TagArray,
				fMapiUnicode,
				&ulValues,
				&lpSourceProp));
			if (!FAILED(hRes) && ulValues && lpSourceProp && PT_ERROR != lpSourceProp->ulPropTag)
			{
				lpSourceProp->ulPropTag = ulTargetTag;
				EC_H(m_lpPropBag->SetProps(
					ulValues,
					lpSourceProp));
			}
		}
		break;
		case 2:
			EC_H(CopyPropertyAsStream(lpSourcePropObj, m_lpPropBag->GetMAPIProp(), ulSourceTag, ulTargetTag));
			break;
		}
	}

	EC_PROBLEMARRAY(lpProblems);
	MAPIFreeBuffer(lpProblems);

	if (!FAILED(hRes))
	{
		EC_H(m_lpPropBag->Commit());

		// refresh
		WC_H(RefreshMAPIPropList());
	}

	lpSourcePropObj->Release();
}

void CSingleMAPIPropListCtrl::OnCopyTo()
{
	// for now, we only copy from objects - copying from rows would be difficult to generalize
	if (!m_lpHostDlg || !m_lpPropBag) return;

	auto lpSourcePropObj = CGlobalCache::getInstance().GetSourcePropObject();
	if (!lpSourcePropObj) return;

	auto hRes = S_OK;

	EC_H(CopyTo(
		m_lpHostDlg->m_hWnd,
		lpSourcePropObj,
		m_lpPropBag->GetMAPIProp(),
		&IID_IMAPIProp,
		NULL,
		m_bIsAB,
		true));
	if (!FAILED(hRes))
	{
		EC_H(m_lpPropBag->Commit());

		// refresh
		WC_H(RefreshMAPIPropList());
	}

	lpSourcePropObj->Release();
}

void CSingleMAPIPropListCtrl::OnOpenProperty() const
{
	auto hRes = S_OK;
	ULONG ulPropTag = NULL;

	if (!m_lpHostDlg) return;

	GetSelectedPropTag(&ulPropTag);
	if (!ulPropTag) return;

	DebugPrintEx(DBGGeneric, CLASS, L"OnOpenProperty", L"asked to open 0x%X\n", ulPropTag);
	LPSPropValue lpProp = nullptr;
	if (m_lpPropBag)
	{
		EC_H(m_lpPropBag->GetProp(ulPropTag, &lpProp));
	}

	if (!FAILED(hRes) && lpProp)
	{
		if (m_lpPropBag && PT_OBJECT == PROP_TYPE(lpProp->ulPropTag))
		{
			EC_H(DisplayTable(
				m_lpPropBag->GetMAPIProp(),
				lpProp->ulPropTag,
				otDefault,
				m_lpHostDlg));
		}
		else if (PT_BINARY == PROP_TYPE(lpProp->ulPropTag) || PT_MV_BINARY == PROP_TYPE(lpProp->ulPropTag))
		{
			switch (PROP_TYPE(lpProp->ulPropTag))
			{
			case PT_BINARY:
				DebugPrintEx(DBGGeneric, CLASS, L"OnOpenProperty", L"property is PT_BINARY\n");
				m_lpHostDlg->OnOpenEntryID(&lpProp->Value.bin);
				break;
			case PT_MV_BINARY:
				DebugPrintEx(DBGGeneric, CLASS, L"OnOpenProperty", L"property is PT_MV_BINARY\n");
				if (S_OK == hRes && lpProp && PT_MV_BINARY == PROP_TYPE(lpProp->ulPropTag))
				{
					DebugPrintEx(DBGGeneric, CLASS, L"OnOpenProperty", L"opened MV structure. There are 0x%X binaries in it.\n", lpProp->Value.MVbin.cValues);
					for (ULONG i = 0; i < lpProp->Value.MVbin.cValues; i++)
					{
						m_lpHostDlg->OnOpenEntryID(&lpProp->Value.MVbin.lpbin[i]);
					}
				}
				break;
			}
		}
	}

	if (m_lpPropBag)
	{
		m_lpPropBag->FreeBuffer(lpProp);
	}
}

void CSingleMAPIPropListCtrl::OnModifyExtraProps()
{
	CGlobalCache::getInstance().MAPIInitialize(NULL);

	auto hRes = S_OK;

	CTagArrayEditor MyTagArrayEditor(
		this,
		IDS_EXTRAPROPS,
		NULL,
		nullptr,
		m_sptExtraProps,
		m_bIsAB,
		m_lpPropBag ? m_lpPropBag->GetMAPIProp() : NULL);

	WC_H(MyTagArrayEditor.DisplayDialog());
	if (S_OK != hRes) return;

	auto lpNewTagArray = MyTagArrayEditor.DetachModifiedTagArray();
	if (lpNewTagArray)
	{
		MAPIFreeBuffer(m_sptExtraProps);
		m_sptExtraProps = lpNewTagArray;
	}

	WC_H(RefreshMAPIPropList());
}

void CSingleMAPIPropListCtrl::OnEditGivenProperty()
{
	if (!m_lpPropBag) return;

	// Display a dialog to get a property number.
	auto hRes = S_OK;

	CPropertyTagEditor MyPropertyTag(
		IDS_EDITGIVENPROP,
		NULL, // prompt
		NULL,
		m_bIsAB,
		m_lpPropBag->GetMAPIProp(),
		this);

	WC_H(MyPropertyTag.DisplayDialog());
	if (S_OK == hRes)
	{
		OnEditGivenProp(MyPropertyTag.GetPropertyTag());
	}
}

void CSingleMAPIPropListCtrl::OnOpenPropertyAsTable()
{
	if (!m_lpPropBag) return;

	// Display a dialog to get a property number.
	auto hRes = S_OK;
	CPropertyTagEditor MyPropertyTag(
		IDS_OPENPROPASTABLE,
		NULL, // prompt
		NULL,
		m_bIsAB,
		m_lpPropBag->GetMAPIProp(),
		this);

	WC_H(MyPropertyTag.DisplayDialog());
	if (S_OK == hRes)
	{
		CEditor MyData(
			this,
			IDS_OPENPROPASTABLE,
			IDS_OPENPROPASTABLEPROMPT,
			CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);

		MyData.InitPane(0, CheckPane::Create(IDS_OPENASEXTABLE, false, false));

		WC_H(MyData.DisplayDialog());
		if (S_OK == hRes)
		{
			if (MyData.GetCheck(0))
			{
				EC_H(DisplayExchangeTable(
					m_lpPropBag->GetMAPIProp(),
					CHANGE_PROP_TYPE(MyPropertyTag.GetPropertyTag(), PT_OBJECT),
					otDefault,
					m_lpHostDlg));
			}
			else
			{
				EC_H(DisplayTable(
					m_lpPropBag->GetMAPIProp(),
					CHANGE_PROP_TYPE(MyPropertyTag.GetPropertyTag(), PT_OBJECT),
					otDefault,
					m_lpHostDlg));
			}
		}
	}
}

void CSingleMAPIPropListCtrl::OnPasteNamedProps()
{
	if (!m_lpPropBag) return;

	auto hRes = S_OK;
	auto lpSourceMsgEID = CGlobalCache::getInstance().GetMessagesToCopy();

	if (CGlobalCache::getInstance().GetBufferStatus() & BUFFER_MESSAGES
		&& lpSourceMsgEID
		&& 1 == lpSourceMsgEID->cValues)
	{
		CEditor MyData(
			this,
			IDS_PASTENAMEDPROPS,
			IDS_PASTENAMEDPROPSPROMPT,
			CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);

		auto szGuid = GUIDToStringAndName(&PS_PUBLIC_STRINGS);
		MyData.InitPane(0, TextPane::CreateSingleLinePane(IDS_GUID, szGuid, false));
		MyData.InitPane(1, CheckPane::Create(IDS_MAPIMOVE, false, false));
		MyData.InitPane(2, CheckPane::Create(IDS_MAPINOREPLACE, false, false));

		WC_H(MyData.DisplayDialog());

		if (S_OK == hRes)
		{
			ULONG ulObjType = 0;
			LPMAPIPROP lpSource = nullptr;
			auto propSetGUID = StringToGUID(MyData.GetStringW(0));

			if (S_OK == hRes)
			{
				EC_H(CallOpenEntry(
					NULL,
					NULL,
					CGlobalCache::getInstance().GetSourceParentFolder(),
					NULL,
					lpSourceMsgEID->lpbin,
					NULL,
					MAPI_BEST_ACCESS,
					&ulObjType,
					reinterpret_cast<LPUNKNOWN*>(&lpSource)));

				if (S_OK == hRes && MAPI_MESSAGE == ulObjType && lpSource)
				{
					EC_H(CopyNamedProps(
						lpSource,
						&propSetGUID,
						MyData.GetCheck(1),
						MyData.GetCheck(2),
						m_lpPropBag->GetMAPIProp(),
						m_lpHostDlg->m_hWnd));

					EC_H(m_lpPropBag->Commit());

					WC_H(RefreshMAPIPropList());

					lpSource->Release();
				}
			}
		}
	}
}

_Check_return_ bool CSingleMAPIPropListCtrl::HandleAddInMenu(WORD wMenuSelect) const
{
	if (wMenuSelect < ID_ADDINPROPERTYMENU) return false;
	CWaitCursor Wait; // Change the mouse to an hourglass while we work.

	auto lpAddInMenu = GetAddinMenuItem(m_lpHostDlg->m_hWnd, wMenuSelect);
	if (!lpAddInMenu) return false;

	_AddInMenuParams MyAddInMenuParams = { nullptr };
	MyAddInMenuParams.lpAddInMenu = lpAddInMenu;
	MyAddInMenuParams.ulAddInContext = MENU_CONTEXT_PROPERTY;
	MyAddInMenuParams.hWndParent = m_hWnd;
	MyAddInMenuParams.lpMAPIProp = m_lpPropBag->GetMAPIProp();
	if (m_lpMapiObjects)
	{
		MyAddInMenuParams.lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
		MyAddInMenuParams.lpMDB = m_lpMapiObjects->GetMDB(); // do not release
		MyAddInMenuParams.lpAdrBook = m_lpMapiObjects->GetAddrBook(false); // do not release
	}

	if (m_lpPropBag && pbRow == m_lpPropBag->GetType())
	{
		SRow MyRow = { 0 };
		(void)m_lpPropBag->GetAllProps(&MyRow.cValues, &MyRow.lpProps);
		MyAddInMenuParams.lpRow = &MyRow;
		MyAddInMenuParams.ulCurrentFlags |= MENU_FLAGS_ROW;
	}

	GetSelectedPropTag(&MyAddInMenuParams.ulPropTag);

	InvokeAddInMenu(&MyAddInMenuParams);
	return true;
}
