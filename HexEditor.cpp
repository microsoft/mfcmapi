// HexEditor.cpp : implementation file
//

#include "stdafx.h"
#include "HexEditor.h"
#include "InterpretProp.h"
#include "MAPIFunctions.h"

static TCHAR* CLASS = _T("CHexEditor");

CHexEditor::CHexEditor(CWnd* pParentWnd):
CEditor(pParentWnd,IDS_HEXEDITOR,IDS_HEXEDITORPROMPT,0,CEDITOR_BUTTON_OK)
{
	TRACE_CONSTRUCTOR(CLASS);
	CreateControls(6);
	InitMultiLine(0,IDS_ANSISTRING,NULL,false);
	InitMultiLine(1,IDS_UNISTRING,NULL,false);
	InitSingleLine(2,IDS_CCH,NULL,true);
	InitMultiLine(3,IDS_BASE64STRING,NULL,false);
	InitSingleLine(4,IDS_CB,NULL,true);
	InitMultiLine(5,IDS_HEX,NULL,false);
}

void CHexEditor::InitAnsiString(
								LPCSTR szInputString)
{
	SetStringA(0,szInputString);
}

void CHexEditor::InitUnicodeString(
								  LPCWSTR szInputString)
{
	SetStringW(0,szInputString);
}

void CHexEditor::InitBase64(
							LPCTSTR szInputBase64String)
{
	SetString(2,szInputBase64String);
}

void CHexEditor::InitBinary(
							LPSBinary lpsInputBin)
{
	SetBinary(5,lpsInputBin->lpb,lpsInputBin->cb);
}

CHexEditor::~CHexEditor()
{
	TRACE_DESTRUCTOR(CLASS);
}

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
	case(0): // ANSI string changed
		{
			lpb = (LPBYTE) GetEditBoxTextA(0);
			if (lpb)
			{
				EC_H(StringCbLengthA((LPCSTR) lpb,STRSAFE_MAX_CCH * sizeof(char),&cb));
			}

			SetStringA(1,(LPCSTR) lpb);

			WC_H(Base64Encode(cb, lpb, &cchEncodeStr, &szEncodeStr));
			SetString(3,szEncodeStr);

			SetBinary(5, lpb, cb);
		}
		break;
	case(1): // Unicode string changed
		{
			lpb = (LPBYTE) GetEditBoxTextW(1);
			if (lpb)
			{
				EC_H(StringCbLengthW((LPCWSTR) lpb,STRSAFE_MAX_CCH * sizeof(WCHAR),&cb));
			}

			SetStringW(0,(LPWSTR) lpb);

			WC_H(Base64Encode(cb, lpb, &cchEncodeStr, &szEncodeStr));
			SetString(3,szEncodeStr);

			SetBinary(5, lpb, cb);
		}
		break;
	case(3): // base64 changed
		{
			szTmpString = GetStringUseControl(3);

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
				SetString(0,BinToTextString(&Bin,true));
				SetString(1,BinToTextString(&Bin,true));
				SetBinary(5, lpb, cb);
			}
			else
			{
				SetString(0,_T(""));
				SetString(1,_T(""));
				SetString(5,_T(""));
				SetBinary(5, 0, 0);
			}
			delete[] lpb;
		}
		break;
	case(5): // binary changed
		{
			if (GetBinaryUseControl(5,&cb,&lpb))
			{
				Bin.lpb = lpb;
				Bin.cb = (ULONG) cb;
				SetString(0,BinToTextString(&Bin,true)); // ansi string

				if (!(cb % 2)) // Set Unicode String
				{
					SetStringW(1,(LPWSTR) lpb);
				}
				else
				{
					SetString(1,_T(""));
				}

				WC_H(Base64Encode(cb, lpb, &cchEncodeStr, &szEncodeStr));
				SetString(3,szEncodeStr);
			}
			else
			{
				SetString(0,_T(""));
				SetString(1,_T(""));
				SetString(3,_T(""));
			}
			delete[] lpb;

		}
		break;
	default:
		break;
	}

	// length of base64 encoded string
	SetSize(2, cchEncodeStr);
	// Length of binary/hex data
	SetSize(4, cb);
	delete[] szEncodeStr;
	return i;
}

