// QSSpecialFolders.cpp : implementation file
//

#include "stdafx.h"
#include "QSSpecialFolders.h"
#include "QuickStart.h"
#include "MAPIFunctions.h"
#include "InterpretProp2.h"
#include "SmartView\SmartView.h"
#include "Editor.h"

static wstring SPECIALFOLDERCLASS = L"SpecialFolderEditor"; // STRING_OK
class SpecialFolderEditor : public CEditor
{
public:
	SpecialFolderEditor(
		_In_ CWnd* pParentWnd,
		_In_ LPMDB lpMDB);
	~SpecialFolderEditor();

private:
	BOOL OnInitDialog();
	void LoadFolders();
	_Check_return_ bool DoListEdit(ULONG ulListNum, int iItem, _In_ SortListData* lpData);

	LPMDB m_lpMDB;
};

SpecialFolderEditor::SpecialFolderEditor(
	_In_ CWnd* pParentWnd,
	_In_ LPMDB lpMDB) :
	CEditor(pParentWnd, IDS_QSSPECIALFOLDERS, NULL, 1, CEDITOR_BUTTON_OK, NULL, NULL, NULL)
{
	TRACE_CONSTRUCTOR(SPECIALFOLDERCLASS);

	m_lpMDB = lpMDB;
	if (m_lpMDB) m_lpMDB->AddRef();

	InitPane(0, CreateListPane(NULL, true, true, this));
}

SpecialFolderEditor::~SpecialFolderEditor()
{
	TRACE_DESTRUCTOR(SPECIALFOLDERCLASS);

	if (m_lpMDB) m_lpMDB->Release();
	m_lpMDB = NULL;
}

BOOL SpecialFolderEditor::OnInitDialog()
{
	BOOL bRet = CEditor::OnInitDialog();
	LoadFolders();
	return bRet;
}

struct sfCol
{
	ULONG ulID;
	ULONG ulType;
};

sfCol g_sfCol[] = {
	{ IDS_SHARP, PT_LONG },
	{ IDS_QSSFFOLDERNAME, PT_NULL },
	{ IDS_QSSFSTENTRYID, PT_NULL },
	{ IDS_QSSFLOCALNAME, PT_NULL },
	{ IDS_QSSFCONTAINERCLASS, PT_NULL },
	{ IDS_QSSFCOMMENT, PT_LONG },
	{ IDS_QSSFCONTENTCOUNT, PT_LONG },
	{ IDS_QSSFHIDDENCONTENTCOUNT, PT_LONG },
	{ IDS_QSSFUNREAD, PT_LONG },
	{ IDS_QSSFCHILDCOUNT, PT_LONG },
	{ IDS_QSSFFOLDERTYPE, PT_NULL },
	{ IDS_QSSFCREATIONTIME, PT_NULL },
	{ IDS_QSSFLASTMODIFICATIONTIME, PT_NULL },
	{ IDS_QSSFRIGHTS, PT_NULL },
	{ IDS_QSSFLTENTRYID, PT_NULL },
};
ULONG g_ulsfCol = _countof(g_sfCol);

void SpecialFolderEditor::LoadFolders()
{
	HRESULT hRes = S_OK;
	ULONG ulListNum = 0;
	if (!IsValidList(ulListNum)) return;

	static const SizedSPropTagArray(12, lptaFolderProps) =
	{
		12,
		PR_DISPLAY_NAME,
		PR_CONTAINER_CLASS,
		PR_COMMENT,
		PR_CONTENT_COUNT,
		PR_ASSOC_CONTENT_COUNT,
		PR_CONTENT_UNREAD,
		PR_FOLDER_CHILD_COUNT,
		PR_FOLDER_TYPE,
		PR_CREATION_TIME,
		PR_LAST_MODIFICATION_TIME,
		PR_RIGHTS,
		PR_ENTRYID,
	};

	ClearList(ulListNum);

	ULONG i = 0;
	for (i = 0; i < g_ulsfCol; i++)
	{
		InsertColumn(ulListNum, i, g_sfCol[i].ulID, g_sfCol[i].ulType);
	}

	CString szTmp;
	wstring szProp;

	// This will iterate over all the special folders we know how to get.
	for (i = DEFAULT_UNSPECIFIED + 1; i < NUM_DEFAULT_PROPS; i++)
	{
		hRes = S_OK;
		ULONG cb = NULL;
		LPENTRYID lpeid = NULL;

		szTmp.Format(_T("%u"), i); // STRING_OK
		SortListData* lpData = InsertListRow(ulListNum, i - 1, szTmp);
		if (lpData)
		{
			int iCol = 1;
			int iRow = i - 1;

			SetListStringA(ulListNum, iRow, iCol, FolderNames[i]);
			iCol++;

			WC_H(GetDefaultFolderEID(i, m_lpMDB, &cb, &lpeid));
			if (SUCCEEDED(hRes))
			{
				SPropValue eid = { 0 };
				eid.ulPropTag = PR_ENTRYID;
				eid.Value.bin.cb = cb;
				eid.Value.bin.lpb = (LPBYTE)lpeid;
				InterpretProp(&eid, &szProp, NULL);
				SetListStringW(ulListNum, iRow, iCol, szProp.c_str());
				iCol++;

				LPMAPIFOLDER lpFolder = NULL;
				WC_H(CallOpenEntry(m_lpMDB, NULL, NULL, NULL, cb, lpeid, NULL, NULL, NULL, (LPUNKNOWN*)&lpFolder));
				if (SUCCEEDED(hRes) && lpFolder)
				{
					ULONG ulProps = 0;
					LPSPropValue lpProps = NULL;
					WC_H_GETPROPS(lpFolder->GetProps((LPSPropTagArray)&lptaFolderProps, fMapiUnicode, &ulProps, &lpProps));

					ULONG ulPropNum = 0;
					for (ulPropNum = 0; ulPropNum < ulProps; ulPropNum++)
					{
						szTmp.Empty();
						if (PT_LONG == PROP_TYPE(lpProps[ulPropNum].ulPropTag))
						{
							wstring szSmartView = InterpretNumberAsString(
								lpProps[ulPropNum].Value,
								lpProps[ulPropNum].ulPropTag,
								NULL,
								NULL,
								NULL,
								false);
							szTmp = szSmartView.c_str();
						}

						if (szTmp.IsEmpty() && PT_ERROR != PROP_TYPE(lpProps[ulPropNum].ulPropTag))
						{
							InterpretProp(&lpProps[ulPropNum], &szProp, NULL);
							SetListStringW(ulListNum, iRow, iCol, szProp.c_str());
						}
						else
						{
							SetListString(ulListNum, iRow, iCol, szTmp);
						}
						iCol++;
					}
				}
				else
				{
					// We couldn't open the folder - log the error
					szTmp.FormatMessage(IDS_QSSFCANNOTOPEN, ErrorNameFromErrorCode(hRes), hRes);
					SetListString(ulListNum, iRow, iCol, szTmp);
				}

				if (lpFolder) lpFolder->Release();
				lpFolder = NULL;
			}
			else
			{
				// We couldn't locate the entry ID- log the error
				szTmp.FormatMessage(IDS_QSSFCANNOTLOCATE, ErrorNameFromErrorCode(hRes), hRes);
				SetListString(ulListNum, iRow, iCol, szTmp);
			}

			MAPIFreeBuffer(lpeid);
			lpData->bItemFullyLoaded = true;
		}
	}
	ResizeList(ulListNum, false);
}

_Check_return_ bool SpecialFolderEditor::DoListEdit(ULONG ulListNum, int iItem, _In_ SortListData* lpData)
{
	if (!IsValidList(ulListNum)) return false;
	if (!lpData) return false;

	HRESULT	hRes = S_OK;

	CEditor MyResults(
		this,
		IDS_QSSPECIALFOLDER,
		NULL,
		1,
		CEDITOR_BUTTON_OK);
	MyResults.InitPane(0, CreateMultiLinePane(NULL, NULL, true));

	CString szTemp1;
	ListPane* listPane = (ListPane*)GetControl(ulListNum);
	if (listPane)
	{
		CString szLabel;
		CString szData;
		ULONG i = 0;
		// We skip the first column, which is just the index
		for (i = 1; i < g_ulsfCol; i++)
		{
			(void)szLabel.LoadString(g_sfCol[i].ulID);
			szData = listPane->GetItemText(iItem, i);
			szTemp1 += szLabel + ": " + szData + "\n";
		}
	}
	if (szTemp1)
	{
		MyResults.SetString(0, szTemp1);
	}

	WC_H(MyResults.DisplayDialog());

	if (S_OK != hRes) return false;

	return false;
}

void OnQSCheckSpecialFolders(_In_ CMainDlg* lpHostDlg, _In_ HWND hwnd)
{
	HRESULT hRes = S_OK;

	lpHostDlg->UpdateStatusBarText(STATUSINFOTEXT, IDS_STATUSTEXTCHECKINGSPECIALFOLDERS);
	lpHostDlg->SendMessage(WM_PAINT, NULL, NULL); // force paint so we update the status now

	LPMDB lpMDB = NULL;
	WC_H(OpenStoreForQuickStart(lpHostDlg, hwnd, &lpMDB));

	if (lpMDB)
	{
		SpecialFolderEditor MyResults(lpHostDlg, lpMDB);

		WC_H(MyResults.DisplayDialog());

		lpMDB->Release();
	}
	lpHostDlg->UpdateStatusBarText(STATUSINFOTEXT, emptystring);
}