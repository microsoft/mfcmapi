// Displays the contents of a form container
#include "stdafx.h"
#include "FormContainerDlg.h"
#include <UI/Controls/ContentsTableListCtrl.h>
#include <MAPI/MapiObjects.h>
#include <UI/Controls/SingleMAPIPropListCtrl.h>
#include <MAPI/ColumnTags.h>
#include <UI/MFCUtilityFunctions.h>
#include <UI/Dialogs/Editors/Editor.h>
#include <MAPI/MAPIFunctions.h>
#include <UI/FileDialogEx.h>
#include <Interpret/InterpretProp.h>

static std::wstring CLASS = L"CFormContainerDlg";

CFormContainerDlg::CFormContainerDlg(
	_In_ CParentWnd* pParentWnd,
	_In_ CMapiObjects* lpMapiObjects,
	_In_ LPMAPIFORMCONTAINER lpFormContainer
) :
	CContentsTableDlg(
		pParentWnd,
		lpMapiObjects,
		IDS_FORMCONTAINER,
		mfcmapiDO_NOT_CALL_CREATE_DIALOG,
		nullptr,
		LPSPropTagArray(&sptDEFCols),
		DEFColumns,
		IDR_MENU_FORM_CONTAINER_POPUP,
		MENU_CONTEXT_FORM_CONTAINER)
{
	TRACE_CONSTRUCTOR(CLASS);

	m_lpFormContainer = lpFormContainer;
	if (m_lpFormContainer)
	{
		m_lpFormContainer->AddRef();
		LPTSTR lpszDisplayName = nullptr;
		(void)m_lpFormContainer->GetDisplay(fMapiUnicode, &lpszDisplayName);
		if (lpszDisplayName)
		{
			m_szTitle = strings::LPCTSTRToWstring(lpszDisplayName);
			MAPIFreeBuffer(lpszDisplayName);
		}
	}

	CContentsTableDlg::CreateDialogAndMenu(IDR_MENU_FORM_CONTAINER);
}

CFormContainerDlg::~CFormContainerDlg()
{
	TRACE_DESTRUCTOR(CLASS);
	if (m_lpFormContainer) m_lpFormContainer->Release();
}

BEGIN_MESSAGE_MAP(CFormContainerDlg, CContentsTableDlg)
	ON_COMMAND(ID_DELETESELECTEDITEM, OnDeleteSelectedItem)
	ON_COMMAND(ID_INSTALLFORM, OnInstallForm)
	ON_COMMAND(ID_REMOVEFORM, OnRemoveForm)
	ON_COMMAND(ID_RESOLVEMESSAGECLASS, OnResolveMessageClass)
	ON_COMMAND(ID_RESOLVEMULTIPLEMESSAGECLASSES, OnResolveMultipleMessageClasses)
	ON_COMMAND(ID_CALCFORMPOPSET, OnCalcFormPropSet)
	ON_COMMAND(ID_GETDISPLAY, OnGetDisplay)
END_MESSAGE_MAP()

void CFormContainerDlg::OnInitMenu(_In_ CMenu* pMenu)
{
	if (pMenu && m_lpContentsTableListCtrl)
	{
		int iNumSel = m_lpContentsTableListCtrl->GetSelectedCount();
		pMenu->EnableMenuItem(ID_DELETESELECTEDITEM, DIMMSOK(iNumSel));
	}
	CContentsTableDlg::OnInitMenu(pMenu);
}

BOOL CFormContainerDlg::OnInitDialog()
{
	auto bRet = CContentsTableDlg::OnInitDialog();

	if (m_lpContentsTableListCtrl)
	{
		OnRefreshView();
	}

	return bRet;
}

// Clear the current list and get a new one
void CFormContainerDlg::OnRefreshView()
{
	auto hRes = S_OK;

	if (!m_lpContentsTableListCtrl) return;

	if (m_lpContentsTableListCtrl->IsLoading()) m_lpContentsTableListCtrl->OnCancelTableLoad();
	DebugPrintEx(DBGForms, CLASS, L"OnRefreshView", L"\n");

	EC_B(m_lpContentsTableListCtrl->DeleteAllItems());
	if (m_lpFormContainer)
	{
		LPSMAPIFORMINFOARRAY lpMAPIFormInfoArray = nullptr;
		WC_MAPI(m_lpFormContainer->ResolveMultipleMessageClasses(nullptr, NULL, &lpMAPIFormInfoArray));
		if (lpMAPIFormInfoArray)
		{
			for (ULONG i = 0; i < lpMAPIFormInfoArray->cForms; i++)
			{
				if (0 == i)
				{
					LPSPropTagArray lpTagArray = nullptr;
					EC_MAPI(lpMAPIFormInfoArray->aFormInfo[i]->GetPropList(fMapiUnicode, &lpTagArray));
					EC_H(m_lpContentsTableListCtrl->SetUIColumns(lpTagArray));
					MAPIFreeBuffer(lpTagArray);
				}
				if (lpMAPIFormInfoArray->aFormInfo[i])
				{
					ULONG ulPropVals = NULL;
					LPSPropValue lpPropVals = nullptr;
					EC_H_GETPROPS(GetPropsNULL(lpMAPIFormInfoArray->aFormInfo[i], fMapiUnicode, &ulPropVals, &lpPropVals));
					if (lpPropVals)
					{
						SRow sRow = { 0 };
						sRow.cValues = ulPropVals;
						sRow.lpProps = lpPropVals;
						(void)::SendMessage(
							m_lpContentsTableListCtrl->m_hWnd,
							WM_MFCMAPI_THREADADDITEM,
							i,
							reinterpret_cast<LPARAM>(&sRow));
					}
					lpMAPIFormInfoArray->aFormInfo[i]->Release();
				}
			}
			MAPIFreeBuffer(lpMAPIFormInfoArray);
		}
	}
	m_lpContentsTableListCtrl->AutoSizeColumns(false);
}

_Check_return_ HRESULT CFormContainerDlg::OpenItemProp(int iSelectedItem, __mfcmapiModifyEnum /*bModify*/, _Deref_out_opt_ LPMAPIPROP* lppMAPIProp)
{
	auto hRes = S_OK;

	DebugPrintEx(DBGOpenItemProp, CLASS, L"OpenItemProp", L"iSelectedItem = 0x%X\n", iSelectedItem);

	if (!lppMAPIProp || !m_lpContentsTableListCtrl || !m_lpFormContainer) return MAPI_E_INVALID_PARAMETER;

	*lppMAPIProp = nullptr;

	auto lpListData = m_lpContentsTableListCtrl->GetSortListData(iSelectedItem);
	if (lpListData)
	{
		auto lpProp = PpropFindProp(
			lpListData->lpSourceProps,
			lpListData->cSourceProps,
			PR_MESSAGE_CLASS_A); // ResolveMessageClass requires an ANSI string
		if (CheckStringProp(lpProp, PT_STRING8))
		{
			LPMAPIFORMINFO lpFormInfoProp = nullptr;
			EC_MAPI(m_lpFormContainer->ResolveMessageClass(
				lpProp->Value.lpszA,
				MAPIFORM_EXACTMATCH,
				&lpFormInfoProp));
			if (SUCCEEDED(hRes))
			{
				*lppMAPIProp = lpFormInfoProp;
			}
			else if (lpFormInfoProp) lpFormInfoProp->Release();
		}
	}
	return hRes;
}

void CFormContainerDlg::OnDeleteSelectedItem()
{
	if (!m_lpFormContainer || !m_lpContentsTableListCtrl) return;

	auto items = m_lpContentsTableListCtrl->GetSelectedItemData();
	for (const auto& lpListData : items)
	{
		auto hRes = S_OK;
		// Find the highlighted item AttachNum
		if (!lpListData) break;

		auto lpProp = PpropFindProp(
			lpListData->lpSourceProps,
			lpListData->cSourceProps,
			PR_MESSAGE_CLASS_A); // RemoveForm requires an ANSI string
		if (CheckStringProp(lpProp, PT_STRING8))
		{
			DebugPrintEx(
				DBGDeleteSelectedItem,
				CLASS,
				L"OnDeleteSelectedItem", // STRING_OK
				L"Removing form \"%hs\"\n", // STRING_OK
				lpProp->Value.lpszA);
			EC_MAPI(m_lpFormContainer->RemoveForm(
				lpProp->Value.lpszA));
		}
	}

	OnRefreshView(); // Update the view since we don't have notifications here.
}

void CFormContainerDlg::OnInstallForm()
{
	auto hRes = S_OK;
	if (!m_lpFormContainer) return;

	DebugPrintEx(DBGForms, CLASS, L"OnInstallForm", L"installing form\n");
	CEditor MyFlags(
		this,
		IDS_INSTALLFORM,
		IDS_INSTALLFORMPROMPT,
		CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
	MyFlags.InitPane(0, TextPane::CreateSingleLinePane(IDS_FLAGS, false));
	MyFlags.SetHex(0, MAPIFORM_INSTALL_DIALOG);

	WC_H(MyFlags.DisplayDialog());
	if (S_OK == hRes)
	{
		auto files = CFileDialogExW::OpenFiles(
			L"cfg", // STRING_OK
			strings::emptystring,
			OFN_FILEMUSTEXIST,
			strings::loadstring(IDS_CFGFILES),
			this);
		if (!files.empty())
		{
			auto ulFlags = MyFlags.GetHex(0);
			auto hwnd = ulFlags & MAPIFORM_INSTALL_DIALOG ? m_hWnd : 0;
			for (auto& lpszPath : files)
			{
				hRes = S_OK;
				DebugPrintEx(DBGForms, CLASS, L"OnInstallForm",
					L"Calling InstallForm(%p,0x%08X,\"%ws\")\n", hwnd, ulFlags, lpszPath.c_str()); // STRING_OK
				WC_MAPI(m_lpFormContainer->InstallForm(reinterpret_cast<ULONG_PTR>(hwnd), ulFlags, LPCTSTR(strings::wstringTostring(lpszPath).c_str())));
				if (MAPI_E_EXTENDED_ERROR == hRes)
				{
					LPMAPIERROR lpErr = nullptr;
					WC_MAPI(m_lpFormContainer->GetLastError(hRes, fMapiUnicode, &lpErr));
					if (lpErr)
					{
						EC_MAPIERR(fMapiUnicode, lpErr);
						MAPIFreeBuffer(lpErr);
					}
				}
				else CHECKHRES(hRes);

				if (bShouldCancel(this, hRes)) break;
			}

			OnRefreshView(); // Update the view since we don't have notifications here.
		}
	}
}

void CFormContainerDlg::OnRemoveForm()
{
	auto hRes = S_OK;
	if (!m_lpFormContainer) return;

	DebugPrintEx(DBGForms, CLASS, L"OnRemoveForm", L"removing form\n");
	CEditor MyClass(
		this,
		IDS_REMOVEFORM,
		IDS_REMOVEFORMPROMPT,
		CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
	MyClass.InitPane(0, TextPane::CreateSingleLinePane(IDS_CLASS, false));

	WC_H(MyClass.DisplayDialog());
	if (S_OK == hRes)
	{
		auto szClass = strings::wstringTostring(MyClass.GetStringW(0)); // RemoveForm requires an ANSI string
		if (!szClass.empty())
		{
			DebugPrintEx(DBGForms, CLASS, L"OnRemoveForm",
				L"Calling RemoveForm(\"%hs\")\n", szClass.c_str()); // STRING_OK
			EC_MAPI(m_lpFormContainer->RemoveForm(szClass.c_str()));
			OnRefreshView(); // Update the view since we don't have notifications here.
		}
	}
}

void CFormContainerDlg::OnResolveMessageClass()
{
	auto hRes = S_OK;
	if (!m_lpFormContainer) return;

	DebugPrintEx(DBGForms, CLASS, L"OnResolveMessageClass", L"resolving message class\n");
	CEditor MyData(
		this,
		IDS_RESOLVECLASS,
		IDS_RESOLVECLASSPROMPT,
		CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
	MyData.InitPane(0, TextPane::CreateSingleLinePane(IDS_CLASS, false));
	MyData.InitPane(1, TextPane::CreateSingleLinePane(IDS_FLAGS, false));

	WC_H(MyData.DisplayDialog());
	if (S_OK == hRes)
	{
		auto szClass = strings::wstringTostring(MyData.GetStringW(0)); // ResolveMessageClass requires an ANSI string
		auto ulFlags = MyData.GetHex(1);
		if (!szClass.empty())
		{
			LPMAPIFORMINFO lpMAPIFormInfo = nullptr;
			DebugPrintEx(DBGForms, CLASS, L"OnResolveMessageClass",
				L"Calling ResolveMessageClass(\"%hs\",0x%08X)\n", szClass.c_str(), ulFlags); // STRING_OK
			EC_MAPI(m_lpFormContainer->ResolveMessageClass(szClass.c_str(), ulFlags, &lpMAPIFormInfo));
			if (lpMAPIFormInfo)
			{
				OnUpdateSingleMAPIPropListCtrl(lpMAPIFormInfo, nullptr);
				DebugPrintFormInfo(DBGForms, lpMAPIFormInfo);
				lpMAPIFormInfo->Release();
			}
		}
	}
}

void CFormContainerDlg::OnResolveMultipleMessageClasses()
{
	auto hRes = S_OK;
	if (!m_lpFormContainer) return;

	DebugPrintEx(DBGForms, CLASS, L"OnResolveMultipleMessageClasses", L"resolving multiple message classes\n");
	CEditor MyData(
		this,
		IDS_RESOLVECLASSES,
		IDS_RESOLVECLASSESPROMPT,
		CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
	MyData.InitPane(0, TextPane::CreateSingleLinePane(IDS_NUMBER, false));
	MyData.SetDecimal(0, 1);
	MyData.InitPane(1, TextPane::CreateSingleLinePane(IDS_FLAGS, false));

	WC_H(MyData.DisplayDialog());
	if (S_OK == hRes)
	{
		auto bCancel = false;
		auto ulNumClasses = MyData.GetDecimal(0);
		auto ulFlags = MyData.GetHex(1);
		LPSMESSAGECLASSARRAY lpMSGClassArray = nullptr;
		if (ulNumClasses && ulNumClasses < MAXMessageClassArray)
		{
			EC_H(MAPIAllocateBuffer(CbMessageClassArray(ulNumClasses), reinterpret_cast<LPVOID*>(&lpMSGClassArray)));

			if (lpMSGClassArray)
			{
				lpMSGClassArray->cValues = ulNumClasses;
				for (ULONG i = 0; i < ulNumClasses; i++)
				{
					CEditor MyClass(
						this,
						IDS_ENTERMSGCLASS,
						IDS_ENTERMSGCLASSPROMPT,
						CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
					MyClass.InitPane(0, TextPane::CreateSingleLinePane(IDS_CLASS, false));

					WC_H(MyClass.DisplayDialog());
					if (S_OK == hRes)
					{
						auto szClass = strings::wstringTostring(MyClass.GetStringW(0)); // MSDN says always use ANSI strings here
						auto cbClass = szClass.length();

						if (cbClass)
						{
							cbClass++; // for the NULL terminator
							EC_H(MAPIAllocateMore(static_cast<ULONG>(cbClass), lpMSGClassArray, (LPVOID*)&lpMSGClassArray->aMessageClass[i]));
							EC_H(StringCbCopyA(const_cast<LPSTR>(lpMSGClassArray->aMessageClass[i]), cbClass, szClass.c_str()));
						}
						else bCancel = true;
					}
					else bCancel = true;
					if (bCancel) break;
				}
			}
		}

		if (!bCancel)
		{
			LPSMAPIFORMINFOARRAY lpMAPIFormInfoArray = nullptr;
			DebugPrintEx(DBGForms, CLASS, L"OnResolveMultipleMessageClasses",
				L"Calling ResolveMultipleMessageClasses(Num Classes = 0x%08X,0x%08X)\n", ulNumClasses, ulFlags); // STRING_OK
			EC_MAPI(m_lpFormContainer->ResolveMultipleMessageClasses(lpMSGClassArray, ulFlags, &lpMAPIFormInfoArray));
			if (lpMAPIFormInfoArray)
			{
				DebugPrintEx(DBGForms, CLASS, L"OnResolveMultipleMessageClasses", L"Got 0x%08X forms\n", lpMAPIFormInfoArray->cForms);
				for (ULONG i = 0; i < lpMAPIFormInfoArray->cForms; i++)
				{
					if (lpMAPIFormInfoArray->aFormInfo[i])
					{
						DebugPrintFormInfo(DBGForms, lpMAPIFormInfoArray->aFormInfo[i]);
						lpMAPIFormInfoArray->aFormInfo[i]->Release();
					}
				}

				MAPIFreeBuffer(lpMAPIFormInfoArray);
			}
		}

		MAPIFreeBuffer(lpMSGClassArray);
	}
}

void CFormContainerDlg::OnCalcFormPropSet()
{
	auto hRes = S_OK;
	if (!m_lpFormContainer) return;

	DebugPrintEx(DBGForms, CLASS, L"OnCalcFormPropSet", L"calculating form property set\n");
	CEditor MyData(
		this,
		IDS_CALCFORMPROPSET,
		IDS_CALCFORMPROPSETPROMPT,
		CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
	MyData.InitPane(0, TextPane::CreateSingleLinePane(IDS_FLAGS, false));
	MyData.SetHex(0, FORMPROPSET_UNION);

	WC_H(MyData.DisplayDialog());
	if (S_OK == hRes)
	{
		auto ulFlags = MyData.GetHex(0);

		LPMAPIFORMPROPARRAY lpFormPropArray = nullptr;
		DebugPrintEx(DBGForms, CLASS, L"OnCalcFormPropSet",
			L"Calling CalcFormPropSet(0x%08X)\n", ulFlags); // STRING_OK
		EC_MAPI(m_lpFormContainer->CalcFormPropSet(ulFlags, &lpFormPropArray));
		if (lpFormPropArray)
		{
			DebugPrintFormPropArray(DBGForms, lpFormPropArray);
			MAPIFreeBuffer(lpFormPropArray);
		}
	}
}

void CFormContainerDlg::OnGetDisplay()
{
	auto hRes = S_OK;
	if (!m_lpFormContainer) return;

	LPTSTR lpszDisplayName = nullptr;
	EC_MAPI(m_lpFormContainer->GetDisplay(fMapiUnicode, &lpszDisplayName));

	if (lpszDisplayName)
	{
		auto szDisplayName = strings::LPCTSTRToWstring(lpszDisplayName);
		DebugPrintEx(DBGForms, CLASS, L"OnGetDisplay", L"Got display name \"%ws\"\n", szDisplayName.c_str());
		CEditor MyOutput(
			this,
			IDS_GETDISPLAY,
			IDS_GETDISPLAYPROMPT,
			CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
		MyOutput.InitPane(0, TextPane::CreateSingleLinePane(IDS_GETDISPLAY, szDisplayName, true));
		WC_H(MyOutput.DisplayDialog());
		MAPIFreeBuffer(lpszDisplayName);
	}
}

void CFormContainerDlg::HandleAddInMenuSingle(
	_In_ LPADDINMENUPARAMS lpParams,
	_In_ LPMAPIPROP lpMAPIProp,
	_In_ LPMAPICONTAINER /*lpContainer*/)
{
	if (lpParams)
	{
		lpParams->lpFormContainer = m_lpFormContainer;
		lpParams->lpFormInfoProp = static_cast<LPMAPIFORMINFO>(lpMAPIProp); // OpenItemProp returns LPMAPIFORMINFO
	}

	InvokeAddInMenu(lpParams);
}