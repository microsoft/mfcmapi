// Displays the contents of a form container
#include <StdAfx.h>
#include <UI/Dialogs/ContentsTable/FormContainerDlg.h>
#include <UI/Controls/ContentsTableListCtrl.h>
#include <MAPI/Cache/MapiObjects.h>
#include <UI/Controls/SingleMAPIPropListCtrl.h>
#include <MAPI/ColumnTags.h>
#include <UI/Dialogs/MFCUtilityFunctions.h>
#include <UI/Dialogs/Editors/Editor.h>
#include <MAPI/MAPIFunctions.h>
#include <UI/FileDialogEx.h>
#include <Interpret/InterpretProp.h>
#include <strsafe.h>

namespace dialog
{
	static std::wstring CLASS = L"CFormContainerDlg";

	CFormContainerDlg::CFormContainerDlg(
		_In_ ui::CParentWnd* pParentWnd,
		_In_ cache::CMapiObjects* lpMapiObjects,
		_In_ LPMAPIFORMCONTAINER lpFormContainer)
		: CContentsTableDlg(
			  pParentWnd,
			  lpMapiObjects,
			  IDS_FORMCONTAINER,
			  mfcmapiDO_NOT_CALL_CREATE_DIALOG,
			  nullptr,
			  nullptr,
			  LPSPropTagArray(&columns::sptDEFCols),
			  columns::DEFColumns,
			  IDR_MENU_FORM_CONTAINER_POPUP,
			  MENU_CONTEXT_FORM_CONTAINER)
	{
		TRACE_CONSTRUCTOR(CLASS);

		m_lpFormContainer = lpFormContainer;
		if (m_lpFormContainer)
		{
			m_lpFormContainer->AddRef();
			LPTSTR lpszDisplayName = nullptr;
			(void) m_lpFormContainer->GetDisplay(fMapiUnicode, &lpszDisplayName);
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
			const int iNumSel = m_lpContentsTableListCtrl->GetSelectedCount();
			pMenu->EnableMenuItem(ID_DELETESELECTEDITEM, DIMMSOK(iNumSel));
		}
		CContentsTableDlg::OnInitMenu(pMenu);
	}

	BOOL CFormContainerDlg::OnInitDialog()
	{
		const auto bRet = CContentsTableDlg::OnInitDialog();

		if (m_lpContentsTableListCtrl)
		{
			OnRefreshView();
		}

		return bRet;
	}

	// Clear the current list and get a new one
	void CFormContainerDlg::OnRefreshView()
	{
		if (!m_lpContentsTableListCtrl) return;

		if (m_lpContentsTableListCtrl->IsLoading()) m_lpContentsTableListCtrl->OnCancelTableLoad();
		output::DebugPrintEx(DBGForms, CLASS, L"OnRefreshView", L"\n");

		EC_B_S(m_lpContentsTableListCtrl->DeleteAllItems());
		if (m_lpFormContainer)
		{
			LPSMAPIFORMINFOARRAY lpMAPIFormInfoArray = nullptr;
			WC_MAPI_S(m_lpFormContainer->ResolveMultipleMessageClasses(nullptr, NULL, &lpMAPIFormInfoArray));
			if (lpMAPIFormInfoArray)
			{
				for (ULONG i = 0; i < lpMAPIFormInfoArray->cForms; i++)
				{
					if (i == 0)
					{
						LPSPropTagArray lpTagArray = nullptr;
						EC_MAPI_S(lpMAPIFormInfoArray->aFormInfo[i]->GetPropList(fMapiUnicode, &lpTagArray));
						m_lpContentsTableListCtrl->SetUIColumns(lpTagArray);
						MAPIFreeBuffer(lpTagArray);
					}

					if (lpMAPIFormInfoArray->aFormInfo[i])
					{
						ULONG ulPropVals = NULL;
						LPSPropValue lpPropVals = nullptr;
						EC_H_GETPROPS_S(mapi::GetPropsNULL(
							lpMAPIFormInfoArray->aFormInfo[i], fMapiUnicode, &ulPropVals, &lpPropVals));
						if (lpPropVals)
						{
							SRow sRow = {0};
							sRow.cValues = ulPropVals;
							sRow.lpProps = lpPropVals;
							(void) ::SendMessage(
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

	_Check_return_ HRESULT CFormContainerDlg::OpenItemProp(
		int iSelectedItem,
		__mfcmapiModifyEnum /*bModify*/,
		_Deref_out_opt_ LPMAPIPROP* lppMAPIProp)
	{
		auto hRes = S_OK;

		output::DebugPrintEx(DBGOpenItemProp, CLASS, L"OpenItemProp", L"iSelectedItem = 0x%X\n", iSelectedItem);

		if (!lppMAPIProp || !m_lpContentsTableListCtrl || !m_lpFormContainer) return MAPI_E_INVALID_PARAMETER;

		*lppMAPIProp = nullptr;

		const auto lpListData = m_lpContentsTableListCtrl->GetSortListData(iSelectedItem);
		if (lpListData)
		{
			const auto lpProp = PpropFindProp(
				lpListData->lpSourceProps,
				lpListData->cSourceProps,
				PR_MESSAGE_CLASS_A); // ResolveMessageClass requires an ANSI string
			if (mapi::CheckStringProp(lpProp, PT_STRING8))
			{
				LPMAPIFORMINFO lpFormInfoProp = nullptr;
				hRes = EC_MAPI(
					m_lpFormContainer->ResolveMessageClass(lpProp->Value.lpszA, MAPIFORM_EXACTMATCH, &lpFormInfoProp));
				if (SUCCEEDED(hRes))
				{
					*lppMAPIProp = lpFormInfoProp;
				}
				else if (lpFormInfoProp)
					lpFormInfoProp->Release();
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
			// Find the highlighted item AttachNum
			if (!lpListData) break;

			const auto lpProp = PpropFindProp(
				lpListData->lpSourceProps,
				lpListData->cSourceProps,
				PR_MESSAGE_CLASS_A); // RemoveForm requires an ANSI string
			if (mapi::CheckStringProp(lpProp, PT_STRING8))
			{
				output::DebugPrintEx(
					DBGDeleteSelectedItem,
					CLASS,
					L"OnDeleteSelectedItem", // STRING_OK
					L"Removing form \"%hs\"\n", // STRING_OK
					lpProp->Value.lpszA);
				EC_MAPI_S(m_lpFormContainer->RemoveForm(lpProp->Value.lpszA));
			}
		}

		OnRefreshView(); // Update the view since we don't have notifications here.
	}

	void CFormContainerDlg::OnInstallForm()
	{
		auto hRes = S_OK;
		if (!m_lpFormContainer) return;

		output::DebugPrintEx(DBGForms, CLASS, L"OnInstallForm", L"installing form\n");
		editor::CEditor MyFlags(
			this, IDS_INSTALLFORM, IDS_INSTALLFORMPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
		MyFlags.InitPane(0, viewpane::TextPane::CreateSingleLinePane(IDS_FLAGS, false));
		MyFlags.SetHex(0, MAPIFORM_INSTALL_DIALOG);

		if (!MyFlags.DisplayDialog()) return;

		auto files = file::CFileDialogExW::OpenFiles(
			L"cfg", // STRING_OK
			strings::emptystring,
			OFN_FILEMUSTEXIST,
			strings::loadstring(IDS_CFGFILES),
			this);
		if (!files.empty())
		{
			const auto ulFlags = MyFlags.GetHex(0);
			auto hwnd = ulFlags & MAPIFORM_INSTALL_DIALOG ? m_hWnd : nullptr;
			for (auto& lpszPath : files)
			{
				output::DebugPrintEx(
					DBGForms,
					CLASS,
					L"OnInstallForm",
					L"Calling InstallForm(%p,0x%08X,\"%ws\")\n",
					hwnd,
					ulFlags,
					lpszPath.c_str()); // STRING_OK
				hRes = WC_MAPI(m_lpFormContainer->InstallForm(
					reinterpret_cast<ULONG_PTR>(hwnd), ulFlags, LPCTSTR(strings::wstringTostring(lpszPath).c_str())));
				if (hRes == MAPI_E_EXTENDED_ERROR)
				{
					LPMAPIERROR lpErr = nullptr;
					hRes = WC_MAPI(m_lpFormContainer->GetLastError(hRes, fMapiUnicode, &lpErr));
					if (lpErr)
					{
						EC_MAPIERR(fMapiUnicode, lpErr);
						MAPIFreeBuffer(lpErr);
					}
				}
				else
					CHECKHRES(hRes);

				if (bShouldCancel(this, hRes)) break;
			}

			OnRefreshView(); // Update the view since we don't have notifications here.
		}
	}

	void CFormContainerDlg::OnRemoveForm()
	{
		if (!m_lpFormContainer) return;

		output::DebugPrintEx(DBGForms, CLASS, L"OnRemoveForm", L"removing form\n");
		editor::CEditor MyClass(this, IDS_REMOVEFORM, IDS_REMOVEFORMPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
		MyClass.InitPane(0, viewpane::TextPane::CreateSingleLinePane(IDS_CLASS, false));

		if (!MyClass.DisplayDialog()) return;

		auto szClass = strings::wstringTostring(MyClass.GetStringW(0)); // RemoveForm requires an ANSI string
		if (!szClass.empty())
		{
			output::DebugPrintEx(
				DBGForms, CLASS, L"OnRemoveForm", L"Calling RemoveForm(\"%hs\")\n", szClass.c_str()); // STRING_OK
			EC_MAPI_S(m_lpFormContainer->RemoveForm(szClass.c_str()));
			OnRefreshView(); // Update the view since we don't have notifications here.
		}
	}

	void CFormContainerDlg::OnResolveMessageClass()
	{
		if (!m_lpFormContainer) return;

		output::DebugPrintEx(DBGForms, CLASS, L"OnResolveMessageClass", L"resolving message class\n");
		editor::CEditor MyData(
			this, IDS_RESOLVECLASS, IDS_RESOLVECLASSPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
		MyData.InitPane(0, viewpane::TextPane::CreateSingleLinePane(IDS_CLASS, false));
		MyData.InitPane(1, viewpane::TextPane::CreateSingleLinePane(IDS_FLAGS, false));

		if (!MyData.DisplayDialog()) return;
		auto szClass = strings::wstringTostring(MyData.GetStringW(0)); // ResolveMessageClass requires an ANSI string
		const auto ulFlags = MyData.GetHex(1);
		if (!szClass.empty())
		{
			LPMAPIFORMINFO lpMAPIFormInfo = nullptr;
			output::DebugPrintEx(
				DBGForms,
				CLASS,
				L"OnResolveMessageClass",
				L"Calling ResolveMessageClass(\"%hs\",0x%08X)\n",
				szClass.c_str(),
				ulFlags); // STRING_OK
			EC_MAPI_S(m_lpFormContainer->ResolveMessageClass(szClass.c_str(), ulFlags, &lpMAPIFormInfo));
			if (lpMAPIFormInfo)
			{
				OnUpdateSingleMAPIPropListCtrl(lpMAPIFormInfo, nullptr);
				output::DebugPrintFormInfo(DBGForms, lpMAPIFormInfo);
				lpMAPIFormInfo->Release();
			}
		}
	}

	void CFormContainerDlg::OnResolveMultipleMessageClasses()
	{
		if (!m_lpFormContainer) return;

		output::DebugPrintEx(
			DBGForms, CLASS, L"OnResolveMultipleMessageClasses", L"resolving multiple message classes\n");
		editor::CEditor MyData(
			this, IDS_RESOLVECLASSES, IDS_RESOLVECLASSESPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
		MyData.InitPane(0, viewpane::TextPane::CreateSingleLinePane(IDS_NUMBER, false));
		MyData.SetDecimal(0, 1);
		MyData.InitPane(1, viewpane::TextPane::CreateSingleLinePane(IDS_FLAGS, false));

		if (!MyData.DisplayDialog()) return;

		auto bCancel = false;
		const auto ulNumClasses = MyData.GetDecimal(0);
		const auto ulFlags = MyData.GetHex(1);
		LPSMESSAGECLASSARRAY lpMSGClassArray = nullptr;
		if (ulNumClasses && ulNumClasses < MAXMessageClassArray)
		{
			auto hRes = EC_H(
				MAPIAllocateBuffer(CbMessageClassArray(ulNumClasses), reinterpret_cast<LPVOID*>(&lpMSGClassArray)));

			if (lpMSGClassArray)
			{
				lpMSGClassArray->cValues = ulNumClasses;
				for (ULONG i = 0; i < ulNumClasses; i++)
				{
					editor::CEditor MyClass(
						this, IDS_ENTERMSGCLASS, IDS_ENTERMSGCLASSPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
					MyClass.InitPane(0, viewpane::TextPane::CreateSingleLinePane(IDS_CLASS, false));

					if (MyClass.DisplayDialog())
					{
						auto szClass =
							strings::wstringTostring(MyClass.GetStringW(0)); // MSDN says always use ANSI strings here
						auto cbClass = szClass.length();

						if (cbClass)
						{
							cbClass++; // for the NULL terminator
							hRes = EC_H(MAPIAllocateMore(
								static_cast<ULONG>(cbClass),
								lpMSGClassArray,
								reinterpret_cast<LPVOID*>(const_cast<LPSTR*>(&lpMSGClassArray->aMessageClass[i]))));

							if (SUCCEEDED(hRes))
							{
								hRes = EC_H(StringCbCopyA(
									const_cast<LPSTR>(lpMSGClassArray->aMessageClass[i]), cbClass, szClass.c_str()));
							}
						}
						else
							bCancel = true;
					}
					else
						bCancel = true;
					if (bCancel) break;
				}
			}
		}

		if (!bCancel)
		{
			LPSMAPIFORMINFOARRAY lpMAPIFormInfoArray = nullptr;
			output::DebugPrintEx(
				DBGForms,
				CLASS,
				L"OnResolveMultipleMessageClasses",
				L"Calling ResolveMultipleMessageClasses(Num Classes = 0x%08X,0x%08X)\n",
				ulNumClasses,
				ulFlags); // STRING_OK
			EC_MAPI_S(m_lpFormContainer->ResolveMultipleMessageClasses(lpMSGClassArray, ulFlags, &lpMAPIFormInfoArray));
			if (lpMAPIFormInfoArray)
			{
				output::DebugPrintEx(
					DBGForms,
					CLASS,
					L"OnResolveMultipleMessageClasses",
					L"Got 0x%08X forms\n",
					lpMAPIFormInfoArray->cForms);
				for (ULONG i = 0; i < lpMAPIFormInfoArray->cForms; i++)
				{
					if (lpMAPIFormInfoArray->aFormInfo[i])
					{
						output::DebugPrintFormInfo(DBGForms, lpMAPIFormInfoArray->aFormInfo[i]);
						lpMAPIFormInfoArray->aFormInfo[i]->Release();
					}
				}

				MAPIFreeBuffer(lpMAPIFormInfoArray);
			}
		}

		MAPIFreeBuffer(lpMSGClassArray);
	}

	void CFormContainerDlg::OnCalcFormPropSet()
	{
		if (!m_lpFormContainer) return;

		output::DebugPrintEx(DBGForms, CLASS, L"OnCalcFormPropSet", L"calculating form property set\n");
		editor::CEditor MyData(
			this, IDS_CALCFORMPROPSET, IDS_CALCFORMPROPSETPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
		MyData.InitPane(0, viewpane::TextPane::CreateSingleLinePane(IDS_FLAGS, false));
		MyData.SetHex(0, FORMPROPSET_UNION);

		if (!MyData.DisplayDialog()) return;

		const auto ulFlags = MyData.GetHex(0);

		LPMAPIFORMPROPARRAY lpFormPropArray = nullptr;
		output::DebugPrintEx(
			DBGForms, CLASS, L"OnCalcFormPropSet", L"Calling CalcFormPropSet(0x%08X)\n", ulFlags); // STRING_OK
		EC_MAPI_S(m_lpFormContainer->CalcFormPropSet(ulFlags, &lpFormPropArray));
		if (lpFormPropArray)
		{
			output::DebugPrintFormPropArray(DBGForms, lpFormPropArray);
			MAPIFreeBuffer(lpFormPropArray);
		}
	}

	void CFormContainerDlg::OnGetDisplay()
	{
		if (!m_lpFormContainer) return;

		LPTSTR lpszDisplayName = nullptr;
		EC_MAPI_S(m_lpFormContainer->GetDisplay(fMapiUnicode, &lpszDisplayName));

		if (lpszDisplayName)
		{
			auto szDisplayName = strings::LPCTSTRToWstring(lpszDisplayName);
			output::DebugPrintEx(
				DBGForms, CLASS, L"OnGetDisplay", L"Got display name \"%ws\"\n", szDisplayName.c_str());
			editor::CEditor MyOutput(
				this, IDS_GETDISPLAY, IDS_GETDISPLAYPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
			MyOutput.InitPane(0, viewpane::TextPane::CreateSingleLinePane(IDS_GETDISPLAY, szDisplayName, true));
			(void) MyOutput.DisplayDialog();
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
			lpParams->lpFormInfoProp = mapi::safe_cast<LPMAPIFORMINFO>(lpMAPIProp);
		}

		addin::InvokeAddInMenu(lpParams);

		if (lpParams && lpParams->lpFormInfoProp)
		{
			lpParams->lpFormInfoProp->Release();
			lpParams->lpFormInfoProp = nullptr;
		}
	}
}