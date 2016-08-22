// HexEditor.cpp : implementation file
//

#include "stdafx.h"
#include "HexEditor.h"
#include "SmartView\SmartView.h"
#include "FileDialogEx.h"
#include "ImportProcs.h"

static wstring CLASS = L"CHexEditor";

enum __HexEditorFields
{
	HEXED_ANSI,
	HEXED_UNICODE,
	HEXED_BASE64,
	HEXED_HEX,
	HEXED_SMARTVIEW
};

CHexEditor::CHexEditor(_In_ CParentWnd* pParentWnd, _In_ CMapiObjects* lpMapiObjects) :
	CEditor(pParentWnd, IDS_HEXEDITOR, NULL, 0, CEDITOR_BUTTON_ACTION1 | CEDITOR_BUTTON_ACTION2 | CEDITOR_BUTTON_ACTION3, IDS_IMPORT, IDS_EXPORT, IDS_CLOSE)
{
	TRACE_CONSTRUCTOR(CLASS);
	m_lpMapiObjects = lpMapiObjects;
	if (m_lpMapiObjects) m_lpMapiObjects->AddRef();

	CreateControls(5);
	InitPane(HEXED_ANSI, CreateCollapsibleTextPane(IDS_ANSISTRING, false));
	InitPane(HEXED_UNICODE, CreateCollapsibleTextPane(IDS_UNISTRING, false));
	InitPane(HEXED_BASE64, CreateCountedTextPane(IDS_BASE64STRING, false, IDS_CCH));
	InitPane(HEXED_HEX, CreateCountedTextPane(IDS_HEX, false, IDS_CB));
	InitPane(HEXED_SMARTVIEW, CreateSmartViewPane(IDS_SMARTVIEW));
	DisplayParentedDialog(pParentWnd, 1000);
} // CHexEditor::CHexEditor

CHexEditor::~CHexEditor()
{
	TRACE_DESTRUCTOR(CLASS);

	if (m_lpMapiObjects) m_lpMapiObjects->Release();
} // CHexEditor::~CHexEditor

void CHexEditor::OnOK()
{
	ShowWindow(SW_HIDE);
	delete this;
} // CHexEditor::OnOK

void CHexEditor::OnCancel()
{
	OnOK();
} // CHexEditor::OnCancel

void CleanString(_In_ CString* lpString)
{
	if (!lpString) return;

	// remove any whitespace
	lpString->Replace(_T("\r"), _T("")); // STRING_OK
	lpString->Replace(_T("\n"), _T("")); // STRING_OK
	lpString->Replace(_T("\t"), _T("")); // STRING_OK
	lpString->Replace(_T(" "), _T("")); // STRING_OK
}

_Check_return_ ULONG CHexEditor::HandleChange(UINT nID)
{
	HRESULT hRes = S_OK;
	ULONG i = CEditor::HandleChange(nID);

	if ((ULONG)-1 == i) return (ULONG)-1;

	CString szTmpString;
	LPBYTE lpb = NULL;
	size_t cb = 0;
	wstring szEncodeStr;
	size_t cchEncodeStr = 0;
	switch (i)
	{
	case HEXED_ANSI:
	{
		size_t cchStr = NULL;
		lpb = (LPBYTE)GetEditBoxTextA(HEXED_ANSI, &cchStr);

		SetStringA(HEXED_UNICODE, (LPCSTR)lpb, cchStr);

		// What we just read includes a NULL terminator, in both the string and count.
		// When we write binary/base64, we don't want to include this NULL
		if (cchStr) cchStr -= 1;
		cb = cchStr * sizeof(CHAR);

		szEncodeStr = Base64Encode(cb, lpb);
		SetStringW(HEXED_BASE64, szEncodeStr.c_str());

		SetBinary(HEXED_HEX, lpb, cb);
	}
	break;
	case HEXED_UNICODE: // Unicode string changed
	{
		size_t cchStr = NULL;
		lpb = (LPBYTE)GetEditBoxTextW(HEXED_UNICODE, &cchStr);

		SetStringW(HEXED_ANSI, (LPWSTR)lpb, cchStr);

		// What we just read includes a NULL terminator, in both the string and count.
		// When we write binary/base64, we don't want to include this NULL
		if (cchStr) cchStr -= 1;
		cb = cchStr * sizeof(WCHAR);

		szEncodeStr = Base64Encode(cb, lpb);
		SetStringW(HEXED_BASE64, szEncodeStr.c_str());

		SetBinary(HEXED_HEX, lpb, cb);
	}
	break;
	case HEXED_BASE64: // base64 changed
	{
		szTmpString = GetStringUseControl(HEXED_BASE64);

		// remove any whitespace before decoding
		CleanString(&szTmpString);

		cchEncodeStr = szTmpString.GetLength();
		WC_H(Base64Decode(LPCTSTRToWstring((LPCTSTR)szTmpString), &cb, &lpb));

		if (S_OK == hRes)
		{
			SetStringA(HEXED_ANSI, (LPCSTR)lpb, cb);
			if (!(cb % 2)) // Set Unicode String
			{
				SetStringW(HEXED_UNICODE, (LPWSTR)lpb, cb / sizeof(WCHAR));
			}
			else
			{
				SetString(HEXED_UNICODE, _T(""));
			}
			SetBinary(HEXED_HEX, lpb, cb);
		}
		else
		{
			SetString(HEXED_ANSI, _T(""));
			SetString(HEXED_UNICODE, _T(""));
			SetBinary(HEXED_HEX, 0, 0);
		}
		delete[] lpb;
	}
	break;
	case HEXED_HEX: // binary changed
	{
		if (GetBinaryUseControl(HEXED_HEX, &cb, &lpb))
		{
			// Treat as a NULL terminated string
			// GetBinaryUseControl includes extra NULLs at the end of the buffer to make this work
			SetStringA(HEXED_ANSI, (LPCSTR)lpb, cb + 1); // ansi string

			if (!(cb % 2)) // Set Unicode String
			{
				// Treat as a NULL terminated string
				// GetBinaryUseControl includes extra NULLs at the end of the buffer to make this work
				SetStringW(HEXED_UNICODE, (LPWSTR)lpb, cb / sizeof(WCHAR) + 1);
			}
			else
			{
				SetString(HEXED_UNICODE, _T(""));
			}

			szEncodeStr = Base64Encode(cb, lpb);
			SetStringW(HEXED_BASE64, szEncodeStr.c_str());
		}
		else
		{
			SetString(HEXED_ANSI, _T(""));
			SetString(HEXED_UNICODE, _T(""));
			SetString(HEXED_BASE64, _T(""));
		}
		delete[] lpb;

	}
	break;
	default:
		break;
	}

	if (HEXED_SMARTVIEW != i)
	{
		// length of base64 encoded string
		CountedTextPane* lpPane = (CountedTextPane*)GetControl(HEXED_BASE64);
		if (lpPane)
		{
			lpPane->SetCount(cchEncodeStr);
		}

		lpPane = (CountedTextPane*)GetControl(HEXED_HEX);
		if (lpPane)
		{
			lpPane->SetCount(cb);
		}
	}
	// Update any parsing we've got:
	UpdateParser();

	// Force the new layout
	OnRecalcLayout();

	return i;
}

void CHexEditor::UpdateParser()
{
	// Find out how to interpret the data
	SmartViewPane* lpPane = (SmartViewPane*)GetControl(HEXED_SMARTVIEW);
	if (lpPane)
	{
		SBinary Bin = { 0 };
		if (GetBinaryUseControl(HEXED_HEX, (size_t*)&Bin.cb, &Bin.lpb))
		{
			lpPane->Parse(Bin);
			delete[] Bin.lpb;
		}
	}
}

// Import
void CHexEditor::OnEditAction1()
{
	HRESULT hRes = S_OK;
	if (S_OK == hRes)
	{
		INT_PTR iDlgRet = IDOK;

		CStringW szFileSpec;
		EC_B(szFileSpec.LoadString(IDS_ALLFILES));

		CFileDialogExW dlgFilePicker;

		EC_D_DIALOG(dlgFilePicker.DisplayDialog(
			true,
			NULL,
			NULL,
			OFN_FILEMUSTEXIST,
			szFileSpec,
			this));
		if (iDlgRet == IDOK && dlgFilePicker.GetFileName())
		{
			if (m_lpMapiObjects) m_lpMapiObjects->MAPIInitialize(NULL);
			LPSTREAM lpStream = NULL;

			// Get a Stream interface on the input file
			EC_H(MyOpenStreamOnFile(
				MAPIAllocateBuffer,
				MAPIFreeBuffer,
				STGM_READ,
				dlgFilePicker.GetFileName(),
				NULL,
				&lpStream));

			if (lpStream)
			{
				TextPane* lpPane = (TextPane*)GetControl(HEXED_HEX);
				if (lpPane)
				{
					lpPane->InitEditFromBinaryStream(lpStream);
				}
				lpStream->Release();
			}
		}
	}
}

// Export
void CHexEditor::OnEditAction2()
{
	HRESULT hRes = S_OK;
	if (S_OK == hRes)
	{
		INT_PTR iDlgRet = IDOK;

		CStringW szFileSpec;
		EC_B(szFileSpec.LoadString(IDS_ALLFILES));

		CFileDialogExW dlgFilePicker;

		EC_D_DIALOG(dlgFilePicker.DisplayDialog(
			false,
			NULL,
			NULL,
			OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT,
			szFileSpec,
			this));
		if (iDlgRet == IDOK && dlgFilePicker.GetFileName())
		{
			if (m_lpMapiObjects) m_lpMapiObjects->MAPIInitialize(NULL);
			LPSTREAM lpStream = NULL;

			// Get a Stream interface on the input file
			EC_H(MyOpenStreamOnFile(
				MAPIAllocateBuffer,
				MAPIFreeBuffer,
				STGM_CREATE | STGM_READWRITE,
				dlgFilePicker.GetFileName(),
				NULL,
				&lpStream));

			if (lpStream)
			{
				TextPane* lpPane = (TextPane*)GetControl(HEXED_HEX);
				if (lpPane)
				{
					lpPane->WriteToBinaryStream(lpStream);
				}
				lpStream->Release();
			}
		}
	}
}

// Close
void CHexEditor::OnEditAction3()
{
	OnOK();
}