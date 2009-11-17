// HexEditor.cpp : implementation file
//

#include "stdafx.h"
#include "HexEditor.h"
#include "InterpretProp.h"
#include "MAPIFunctions.h"
#include "InterpretProp2.h"
#include "ParentWnd.h"

static TCHAR* CLASS = _T("CHexEditor");

enum __PropTagFields
{
	HEXED_ANSI,
	HEXED_UNICODE,
	HEXED_CCH,
	HEXED_BASE64,
	HEXED_CB,
	HEXED_HEX,
	HEXED_PICKER,
	HEXED_SMARTVIEW
};

CHexEditor::CHexEditor(CParentWnd* pParentWnd):
CEditor(pParentWnd,IDS_HEXEDITOR,IDS_HEXEDITORPROMPT,0,CEDITOR_BUTTON_OK)
{
	TRACE_CONSTRUCTOR(CLASS);
	CreateControls(8);
	InitMultiLine(HEXED_ANSI,IDS_ANSISTRING,NULL,false);
	InitMultiLine(HEXED_UNICODE,IDS_UNISTRING,NULL,false);
	InitSingleLine(HEXED_CCH,IDS_CCH,NULL,true);
	InitMultiLine(HEXED_BASE64,IDS_BASE64STRING,NULL,false);
	InitSingleLine(HEXED_CB,IDS_CB,NULL,true);
	InitMultiLine(HEXED_HEX,IDS_HEX,NULL,false);
	InitDropDown(HEXED_PICKER,IDS_STRUCTUREPICKERPROMPT,g_cuidParsingTypesDropDown,g_uidParsingTypesDropDown,true);
	InitMultiLine(HEXED_SMARTVIEW,IDS_PARSEDSTRUCTURE,NULL,true);
	HRESULT hRes = S_OK;
	m_lpszTemplateName = MAKEINTRESOURCE(IDD_BLANK_DIALOG);

	m_lpNonModalParent = pParentWnd;
	if (m_lpNonModalParent) m_lpNonModalParent->AddRef();

	m_hwndCenteringWindow = GetActiveWindow();

	HINSTANCE hInst = AfxFindResourceHandle(m_lpszTemplateName, RT_DIALOG);
	HRSRC hResource = NULL;
	EC_D(hResource,::FindResource(hInst, m_lpszTemplateName, RT_DIALOG));
	HGLOBAL hTemplate = NULL;
	EC_D(hTemplate,LoadResource(hInst, hResource));
	LPCDLGTEMPLATE lpDialogTemplate = (LPCDLGTEMPLATE)LockResource(hTemplate);
	EC_B(CreateDlgIndirect(lpDialogTemplate, m_lpNonModalParent, hInst));
}

CHexEditor::~CHexEditor()
{
	TRACE_DESTRUCTOR(CLASS);
	if (m_lpNonModalParent) m_lpNonModalParent->Release();
}

void CHexEditor::OnOK()
{
	ShowWindow(SW_HIDE);
	delete this;
}

void CHexEditor::OnCancel()
{
	OnOK();
}

// MFC will call this function to check if it ought to center the dialog
// We'll tell it no, but also place the dialog where we want it.
BOOL CHexEditor::CheckAutoCenter()
{
	// We can make the hex editor wider - OnSize will fix the height for us
	SetWindowPos(NULL,0,0,1000,0,NULL);
	CenterWindow(m_hwndCenteringWindow);
	return false;
} // CHexEditor::CheckAutoCenter

ULONG CHexEditor::HandleChange(UINT nID)
{
	HRESULT hRes = S_OK;
	ULONG i = CEditor::HandleChange(nID);

	if ((ULONG) -1 == i) return (ULONG) -1;

	CString szTmpString;
	SBinary Bin = {0};
	LPBYTE	lpb = NULL;
	size_t	cb = 0;
	LPTSTR	szEncodeStr = NULL;
	size_t	cchEncodeStr = 0;
	switch (i)
	{
	case(HEXED_ANSI):
		{
			size_t cchStr = NULL;
			lpb = (LPBYTE) GetEditBoxTextA(HEXED_ANSI,&cchStr);

			SetStringA(HEXED_UNICODE, (LPCSTR) lpb, cchStr);

			// What we just read includes a NULL terminator, in both the string and count.
			// When we write binary/base64, we don't want to include this NULL
			if (cchStr) cchStr -= 1;
			cb = cchStr * sizeof(CHAR);

			WC_H(Base64Encode(cb, lpb, &cchEncodeStr, &szEncodeStr));
			SetString(HEXED_BASE64,szEncodeStr);

			SetBinary(HEXED_HEX, lpb, cb);
		}
		break;
	case(HEXED_UNICODE): // Unicode string changed
		{
			size_t cchStr = NULL;
			lpb = (LPBYTE) GetEditBoxTextW(HEXED_UNICODE, &cchStr);

			SetStringW(HEXED_ANSI, (LPWSTR) lpb, cchStr);

			// What we just read includes a NULL terminator, in both the string and count.
			// When we write binary/base64, we don't want to include this NULL
			if (cchStr) cchStr -= 1;
			cb = cchStr * sizeof(WCHAR);

			WC_H(Base64Encode(cb, lpb, &cchEncodeStr, &szEncodeStr));
			SetString(HEXED_BASE64,szEncodeStr);

			SetBinary(HEXED_HEX, lpb, cb);
		}
		break;
	case(HEXED_BASE64): // base64 changed
		{
			szTmpString = GetStringUseControl(HEXED_BASE64);

			// remove any whitespace before decoding
			szTmpString.Replace(_T("\r"),_T("")); // STRING_OK
			szTmpString.Replace(_T("\n"),_T("")); // STRING_OK
			szTmpString.Replace(_T("\t"),_T("")); // STRING_OK
			szTmpString.Replace(_T(" "),_T("")); // STRING_OK

			cchEncodeStr = szTmpString.GetLength();
			WC_H(Base64Decode(szTmpString,&cb, &lpb));

			if (S_OK == hRes)
			{
				Bin.lpb = lpb;
				Bin.cb = (ULONG) cb;
				SetStringA(HEXED_ANSI, (LPCSTR)lpb, cb);
				if (!(cb % 2)) // Set Unicode String
				{
					SetStringW(HEXED_UNICODE, (LPWSTR)lpb, cb/sizeof(WCHAR));
				}
				SetBinary(HEXED_HEX, lpb, cb);
			}
			else
			{
				SetString(HEXED_ANSI,_T(""));
				SetString(HEXED_UNICODE,_T(""));
				SetBinary(HEXED_HEX, 0, 0);
			}
			delete[] lpb;
		}
		break;
	case(HEXED_HEX): // binary changed
		{
			if (GetBinaryUseControl(HEXED_HEX,&cb,&lpb))
			{
				// Treat as a NULL terminated string
				// GetBinaryUseControl includes extra NULLs at the end of the buffer to make this work
				SetStringA(HEXED_ANSI,(LPCSTR) lpb, cb+1); // ansi string

				if (!(cb % 2)) // Set Unicode String
				{
					// Treat as a NULL terminated string
					// GetBinaryUseControl includes extra NULLs at the end of the buffer to make this work
					SetStringW(HEXED_UNICODE,(LPWSTR) lpb, cb/sizeof(WCHAR)+1);
				}
				else
				{
					SetString(HEXED_UNICODE,_T(""));
				}

				WC_H(Base64Encode(cb, lpb, &cchEncodeStr, &szEncodeStr));
				SetString(HEXED_BASE64,szEncodeStr);
			}
			else
			{
				SetString(HEXED_ANSI,_T(""));
				SetString(HEXED_UNICODE,_T(""));
				SetString(HEXED_BASE64,_T(""));
			}
			delete[] lpb;

		}
		break;
	default:
		break;
	}

	if (HEXED_PICKER != i)
	{
		// length of base64 encoded string
		SetSize(HEXED_CCH, cchEncodeStr);
		// Length of binary/hex data
		SetSize(HEXED_CB, cb);
	}
	// Update any parsing we've got:
	UpdateParser();
	delete[] szEncodeStr;
	return i;
} // CHexEditor::HandleChange

void CHexEditor::UpdateParser()
{
	// Find out how to interpret the data
	DWORD_PTR iStructType = GetDropDownSelectionValue(HEXED_PICKER);

	LPTSTR szString = NULL;
	if (iStructType)
	{
		SBinary Bin = {0};
		if (GetBinaryUseControl(HEXED_HEX,(size_t*) &Bin.cb,&Bin.lpb))
		{
			// Get the string interpretation
			InterpretBinaryAsString(Bin,iStructType,NULL,NULL,&szString);
			delete[] Bin.lpb;
		}
	}

	if (szString)
	{
		SetString(HEXED_SMARTVIEW,szString);

		delete[] szString;
	}
	else
	{
		SetString(HEXED_SMARTVIEW,_T(""));
	}
}