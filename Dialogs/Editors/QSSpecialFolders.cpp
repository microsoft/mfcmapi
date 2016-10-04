#include "stdafx.h"
#include "QSSpecialFolders.h"
#include "QuickStart.h"
#include "MAPIFunctions.h"
#include "SmartView/SmartView.h"
#include <Dialogs/Editors/Editor.h>
#include <Dialogs/ContentsTable/MainDlg.h>

static wstring SPECIALFOLDERCLASS = L"SpecialFolderEditor"; // STRING_OK
class SpecialFolderEditor : public CEditor
{
public:
	SpecialFolderEditor(
		_In_ CWnd* pParentWnd,
		_In_ LPMDB lpMDB);
	~SpecialFolderEditor();

private:
	BOOL OnInitDialog() override;
	void LoadFolders() const;
	_Check_return_ bool DoListEdit(ULONG ulListNum, int iItem, _In_ SortListData* lpData) override;

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

	InitPane(0, ListPane::Create(NULL, true, true, this));
}

SpecialFolderEditor::~SpecialFolderEditor()
{
	TRACE_DESTRUCTOR(SPECIALFOLDERCLASS);

	if (m_lpMDB) m_lpMDB->Release();
	m_lpMDB = nullptr;
}

BOOL SpecialFolderEditor::OnInitDialog()
{
	auto bRet = CEditor::OnInitDialog();
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

void SpecialFolderEditor::LoadFolders() const
{
	auto hRes = S_OK;
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

	for (ULONG i = 0; i < g_ulsfCol; i++)
	{
		InsertColumn(ulListNum, i, g_sfCol[i].ulID, g_sfCol[i].ulType);
	}

	wstring szTmp;
	wstring szProp;

	// This will iterate over all the special folders we know how to get.
	for (ULONG i = DEFAULT_UNSPECIFIED + 1; i < NUM_DEFAULT_PROPS; i++)
	{
		hRes = S_OK;
		ULONG cb = NULL;
		LPENTRYID lpeid = nullptr;

		auto lpData = InsertListRow(ulListNum, i - 1, format(L"%u", i)); // STRING_OK
		if (lpData)
		{
			auto iCol = 1;
			int iRow = i - 1;

			SetListString(ulListNum, iRow, iCol, FolderNames[i]);
			iCol++;

			WC_H(GetDefaultFolderEID(i, m_lpMDB, &cb, &lpeid));
			if (SUCCEEDED(hRes))
			{
				SPropValue eid = { 0 };
				eid.ulPropTag = PR_ENTRYID;
				eid.Value.bin.cb = cb;
				eid.Value.bin.lpb = reinterpret_cast<LPBYTE>(lpeid);
				InterpretProp(&eid, &szProp, nullptr);
				SetListString(ulListNum, iRow, iCol, szProp);
				iCol++;

				LPMAPIFOLDER lpFolder = nullptr;
				WC_H(CallOpenEntry(m_lpMDB, NULL, NULL, NULL, cb, lpeid, NULL, NULL, NULL, reinterpret_cast<LPUNKNOWN*>(&lpFolder)));
				if (SUCCEEDED(hRes) && lpFolder)
				{
					ULONG ulProps = 0;
					LPSPropValue lpProps = nullptr;
					WC_H_GETPROPS(lpFolder->GetProps(LPSPropTagArray(&lptaFolderProps), fMapiUnicode, &ulProps, &lpProps));

					for (ULONG ulPropNum = 0; ulPropNum < ulProps; ulPropNum++)
					{
						szTmp.clear();
						if (PT_LONG == PROP_TYPE(lpProps[ulPropNum].ulPropTag))
						{
							szTmp = InterpretNumberAsString(
								lpProps[ulPropNum].Value,
								lpProps[ulPropNum].ulPropTag,
								NULL,
								nullptr,
								nullptr,
								false);
						}

						if (szTmp.empty() && PT_ERROR != PROP_TYPE(lpProps[ulPropNum].ulPropTag))
						{
							InterpretProp(&lpProps[ulPropNum], &szProp, nullptr);
							SetListString(ulListNum, iRow, iCol, szProp);
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
					szTmp = formatmessage(IDS_QSSFCANNOTOPEN, ErrorNameFromErrorCode(hRes).c_str(), hRes);
					SetListString(ulListNum, iRow, iCol, szTmp);
				}

				if (lpFolder) lpFolder->Release();
				lpFolder = nullptr;
			}
			else
			{
				// We couldn't locate the entry ID- log the error
				szTmp = formatmessage(IDS_QSSFCANNOTLOCATE, ErrorNameFromErrorCode(hRes).c_str(), hRes);
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

	auto hRes = S_OK;

	CEditor MyResults(
		this,
		IDS_QSSPECIALFOLDER,
		NULL,
		1,
		CEDITOR_BUTTON_OK);
	MyResults.InitPane(0, TextPane::CreateMultiLinePane(NULL, true));

	wstring szTmp;
	auto listPane = static_cast<ListPane*>(GetControl(ulListNum));
	if (listPane)
	{
		wstring szLabel;
		wstring szData;
		// We skip the first column, which is just the index
		for (ULONG i = 1; i < g_ulsfCol; i++)
		{
			szLabel = loadstring(g_sfCol[i].ulID);
			szData = listPane->GetItemText(iItem, i);
			szTmp += szLabel + L": " + szData + L"\n";
		}
	}

	if (!szTmp.empty())
	{
		MyResults.SetStringW(0, szTmp);
	}

	WC_H(MyResults.DisplayDialog());

	if (S_OK != hRes) return false;

	return false;
}

void OnQSCheckSpecialFolders(_In_ CMainDlg* lpHostDlg, _In_ HWND hwnd)
{
	auto hRes = S_OK;

	lpHostDlg->UpdateStatusBarText(STATUSINFOTEXT, IDS_STATUSTEXTCHECKINGSPECIALFOLDERS);
	lpHostDlg->SendMessage(WM_PAINT, NULL, NULL); // force paint so we update the status now

	LPMDB lpMDB = nullptr;
	WC_H(OpenStoreForQuickStart(lpHostDlg, hwnd, &lpMDB));

	if (lpMDB)
	{
		SpecialFolderEditor MyResults(lpHostDlg, lpMDB);

		WC_H(MyResults.DisplayDialog());

		lpMDB->Release();
	}

	lpHostDlg->UpdateStatusBarText(STATUSINFOTEXT, emptystring);
}