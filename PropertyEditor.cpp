// PropertyEditor.cpp : implementation file
//

#include "stdafx.h"
#include "PropertyEditor.h"
#include "InterpretProp2.h"
#include "MAPIFunctions.h"
#include "SmartView.h"

_Check_return_ HRESULT DisplayPropertyEditor(_In_ CWnd* pParentWnd,
											 UINT uidTitle,
											 UINT uidPrompt,
											 BOOL bIsAB,
											 _In_opt_ LPVOID lpAllocParent,
											 _In_opt_ LPMAPIPROP lpMAPIProp,
											 ULONG ulPropTag,
											 BOOL bMVRow,
											 _In_opt_ LPSPropValue lpsPropValue,
											 _Inout_opt_ LPSPropValue* lpNewValue)
{
	HRESULT hRes = S_OK;
	BOOL bShouldFreeInputValue = false;

	// We got a MAPI prop object and no input value, go look one up
	if (lpMAPIProp && !lpsPropValue)
	{
		SPropTagArray sTag = {0};
		sTag.cValues = 1;
		sTag.aulPropTag[0] = (PT_ERROR == PROP_TYPE(ulPropTag))?CHANGE_PROP_TYPE(ulPropTag,PT_UNSPECIFIED):ulPropTag;
		ULONG ulValues = NULL;

		WC_H(lpMAPIProp->GetProps(&sTag,NULL,&ulValues,&lpsPropValue));

		// Suppress MAPI_E_NOT_FOUND error when the source type is non error
		if (lpsPropValue &&
			PT_ERROR == PROP_TYPE(lpsPropValue->ulPropTag) &&
			MAPI_E_NOT_FOUND == lpsPropValue->Value.err &&
			PT_ERROR != PROP_TYPE(ulPropTag)
			)
		{
			MAPIFreeBuffer(lpsPropValue);
			lpsPropValue = NULL;
		}

		// In all cases where we got a value back, we need to reset our property tag to the value we got
		// This will address when the source is PT_UNSPECIFIED, when the returned value is PT_ERROR,
		// or any other case where the returned value has a different type than requested
		if (SUCCEEDED(hRes) && lpsPropValue)
			ulPropTag = lpsPropValue->ulPropTag;

		bShouldFreeInputValue = true;
	}
	else if (lpsPropValue && !ulPropTag)
	{
		ulPropTag = lpsPropValue->ulPropTag;
	}

	// Check for the multivalue prop case
	if (PROP_TYPE(ulPropTag) & MV_FLAG)
	{
		CMultiValuePropertyEditor MyPropertyEditor(
			pParentWnd,
			uidTitle,
			uidPrompt,
			bIsAB,
			lpAllocParent,
			lpMAPIProp,
			ulPropTag,
			lpsPropValue);
		WC_H(MyPropertyEditor.DisplayDialog());

		if (lpNewValue) *lpNewValue = MyPropertyEditor.DetachModifiedSPropValue();
	}
	// Or the single value prop case
	else
	{
		CPropertyEditor MyPropertyEditor(
			pParentWnd,
			uidTitle,
			uidPrompt,
			bIsAB,
			bMVRow,
			lpAllocParent,
			lpMAPIProp,
			ulPropTag,
			lpsPropValue);
		WC_H(MyPropertyEditor.DisplayDialog());

		if (lpNewValue) *lpNewValue = MyPropertyEditor.DetachModifiedSPropValue();
	}

	if (bShouldFreeInputValue)
		MAPIFreeBuffer(lpsPropValue);

	return hRes;
} // DisplayPropertyEditor

static TCHAR* SVCLASS = _T("CPropertyEditor"); // STRING_OK

// Create an editor for a MAPI property
CPropertyEditor::CPropertyEditor(
								 _In_ CWnd* pParentWnd,
								 UINT uidTitle,
								 UINT uidPrompt,
								 BOOL bIsAB,
								 BOOL bMVRow,
								 _In_opt_ LPVOID lpAllocParent,
								 _In_opt_ LPMAPIPROP lpMAPIProp,
								 ULONG ulPropTag,
								 _In_opt_ LPSPropValue lpsPropValue):
CEditor(pParentWnd,uidTitle,uidPrompt,0,CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL)
{
	TRACE_CONSTRUCTOR(SVCLASS);

	m_bIsAB = bIsAB;
	m_bMVRow = bMVRow;
	m_lpAllocParent = lpAllocParent;
	m_lpsOutputValue = NULL;
	m_bDirty = false;

	m_lpMAPIProp = lpMAPIProp;
	if (m_lpMAPIProp) m_lpMAPIProp->AddRef();
	m_ulPropTag = ulPropTag;
	m_lpsInputValue = lpsPropValue;

	// If we didn't have an input value, we are creating a new property
	// So by definition, we're already dirty
	if (!m_lpsInputValue) m_bDirty = true;

	CString szPromptPostFix;
	szPromptPostFix.Format(_T("\r\n%s"),(LPCTSTR) TagToString(m_ulPropTag | (m_bMVRow?MV_FLAG:NULL),m_lpMAPIProp,m_bIsAB,false)); // STRING_OK

	SetPromptPostFix(szPromptPostFix);

	// Let's crack our property open and see what kind of controls we'll need for it
	CreatePropertyControls();

	InitPropertyControls();
} // CPropertyEditor::CPropertyEditor

CPropertyEditor::~CPropertyEditor()
{
	TRACE_DESTRUCTOR(SVCLASS);
	if (m_lpMAPIProp) m_lpMAPIProp->Release();
} // CPropertyEditor::~CPropertyEditor

_Check_return_ BOOL CPropertyEditor::OnInitDialog()
{
	BOOL bRet = CEditor::OnInitDialog();
	return bRet;
} // CPropertyEditor::OnInitDialog

void CPropertyEditor::OnOK()
{
	// This is where we write our changes back
	WriteStringsToSPropValue();

	// Write the property to the object if we're not editing a row of a MV property
	if (!m_bMVRow) WriteSPropValueToObject();
	CDialog::OnOK(); // don't need to call CEditor::OnOK
} // CPropertyEditor::OnOK

void CPropertyEditor::CreatePropertyControls()
{
	switch (PROP_TYPE(m_ulPropTag))
	{
	case(PT_APPTIME):
	case(PT_BOOLEAN):
	case(PT_DOUBLE):
	case(PT_OBJECT):
	case(PT_R4):
		CreateControls(1);
		break;
	case(PT_ERROR):
		CreateControls(2);
		break;
	case(PT_BINARY):
		CreateControls(4);
		break;
	case(PT_CURRENCY):
	case(PT_I8):
	case(PT_LONG):
	case(PT_I2):
	case(PT_SYSTIME):
	case(PT_STRING8):
	case(PT_UNICODE):
		CreateControls(3);
		break;
	case(PT_CLSID):
		CreateControls(1);
		break;
	case(PT_ACTIONS):
		CreateControls(1);
		break;
	case(PT_SRESTRICTION):
		CreateControls(1);
		break;
	default:
		CreateControls(2);
		break;
	}
} // CPropertyEditor::CreatePropertyControls

void CPropertyEditor::InitPropertyControls()
{
	LPTSTR szSmartView = NULL;

	InterpretPropSmartView(m_lpsInputValue,
		m_lpMAPIProp,
		NULL,
		NULL,
		m_bMVRow,
		&szSmartView); // Built from lpProp & lpMAPIProp

	CString szTemp1;
	CString szTemp2;
	switch (PROP_TYPE(m_ulPropTag))
	{
	case(PT_APPTIME):
		InitSingleLine(0,IDS_DOUBLE,NULL,false);
		if (m_lpsInputValue)
		{
			SetStringf(0,_T("%f"),m_lpsInputValue->Value.at); // STRING_OK
		}
		else
		{
			SetDecimal(0,0);
		}
		break;
	case(PT_BOOLEAN):
		InitCheck(0,IDS_BOOLEAN,m_lpsInputValue?m_lpsInputValue->Value.b:false,false);
		break;
	case(PT_DOUBLE):
		InitSingleLine(0,IDS_DOUBLE,NULL,false);
		if (m_lpsInputValue)
		{
			SetStringf(0,_T("%f"),m_lpsInputValue->Value.dbl); // STRING_OK
		}
		else
		{
			SetDecimal(0,0);
		}
		break;
	case(PT_OBJECT):
		InitSingleLine(0,IDS_OBJECT,IDS_OBJECTVALUE,true);
		break;
	case(PT_R4):
		InitSingleLine(0,IDS_FLOAT,NULL,false);
		if (m_lpsInputValue)
		{
			SetStringf(0,_T("%f"),m_lpsInputValue->Value.flt); // STRING_OK
		}
		else
		{
			SetDecimal(0,0);
		}
		break;
	case(PT_STRING8):
		InitMultiLine(0,IDS_ANSISTRING,NULL,false);
		InitSingleLine(1,IDS_CB,NULL,true);
		SetSize(1,0);
		InitMultiLine(2,IDS_BIN,NULL,false);
		if (m_lpsInputValue && CheckStringProp(m_lpsInputValue,PT_STRING8))
		{
			SetStringA(0,m_lpsInputValue->Value.lpszA);

			size_t cbStr = 0;
			HRESULT hRes = S_OK;
			EC_H(StringCbLengthA(m_lpsInputValue->Value.lpszA,STRSAFE_MAX_CCH * sizeof(char),&cbStr));

			SetSize(1, cbStr);

			SetBinary(
				2,
				(LPBYTE) m_lpsInputValue->Value.lpszA,
				cbStr);
		}

		break;
	case(PT_UNICODE):
		InitMultiLine(0,IDS_UNISTRING,NULL,false);
		InitSingleLine(1,IDS_CB,NULL,true);
		SetSize(1,0);
		InitMultiLine(2,IDS_BIN,NULL,false);
		if (m_lpsInputValue && CheckStringProp(m_lpsInputValue,PT_UNICODE))
		{
			SetStringW(0, m_lpsInputValue->Value.lpszW);

			size_t cbStr = 0;
			HRESULT hRes = S_OK;
			EC_H(StringCbLengthW(m_lpsInputValue->Value.lpszW,STRSAFE_MAX_CCH * sizeof(WCHAR),&cbStr));

			SetSize(1, cbStr);

			SetBinary(
				2,
				(LPBYTE) m_lpsInputValue->Value.lpszW,
				cbStr);
		}

		break;
	case(PT_CURRENCY):
		InitSingleLine(0,IDS_HI,NULL,false);
		InitSingleLine(1,IDS_LO,NULL,false);
		InitSingleLine(2,IDS_CURRENCY,NULL,false);
		if (m_lpsInputValue)
		{
			SetHex(0,m_lpsInputValue->Value.cur.Hi);
			SetHex(1,m_lpsInputValue->Value.cur.Lo);
			SetString(2,CurrencyToString(m_lpsInputValue->Value.cur));
		}
		else
		{
			SetHex(0,0);
			SetHex(1,0);
			SetString(2,_T("0.0000")); // STRING_OK
		}
		break;
	case(PT_ERROR):
		InitSingleLine(0,IDS_ERRORCODEHEX,NULL,true);
		InitSingleLine(1,IDS_ERRORNAME,NULL,true);
		if (m_lpsInputValue)
		{
			SetHex(0, m_lpsInputValue->Value.err);
			SetString(1,ErrorNameFromErrorCode(m_lpsInputValue->Value.err));
		}
		break;
	case(PT_I2):
		InitSingleLine(0,IDS_SIGNEDDECIMAL,NULL,false);
		InitSingleLine(1,IDS_HEX,NULL,false);
		InitMultiLine(2,IDS_COLSMART_VIEW,NULL,true);
		if (m_lpsInputValue)
		{
			SetDecimal(0,m_lpsInputValue->Value.i);
			SetHex(1,m_lpsInputValue->Value.i);

			if (szSmartView) SetString(2,szSmartView);
		}
		else
		{
			SetDecimal(0,0);
			SetHex(1,0);
			SetHex(2,0);
		}
		break;
	case(PT_I8):
		InitSingleLine(0,IDS_HIGHPART,NULL,false);
		InitSingleLine(1,IDS_LOWPART,NULL,false);
		InitSingleLine(2,IDS_DECIMAL,NULL,false);

		if (m_lpsInputValue)
		{
			SetHex(0,(int) m_lpsInputValue->Value.li.HighPart);
			SetHex(1,(int) m_lpsInputValue->Value.li.LowPart);
			SetStringf(2,_T("%I64d"),m_lpsInputValue->Value.li.QuadPart); // STRING_OK
		}
		else
		{
			SetHex(0,0);
			SetHex(1,0);
			SetDecimal(2,0);
		}
		break;
	case(PT_BINARY):
		InitSingleLine(0,IDS_CB,NULL,true);
		SetHex(0,0);
		if (m_lpsInputValue)
		{
			SetSize(0,m_lpsInputValue->Value.bin.cb);
			InitMultiLine(1,IDS_BIN,BinToHexString(&m_lpsInputValue->Value.bin,false),false);
			InitMultiLine(2,IDS_TEXT,NULL,false);
			SetStringA(2,(LPCSTR)m_lpsInputValue->Value.bin.lpb,m_lpsInputValue->Value.bin.cb+1);
			InitMultiLine(3,IDS_COLSMART_VIEW,szSmartView,true);
		}
		else
		{
			InitMultiLine(1,IDS_BIN,NULL,false);
			InitMultiLine(2,IDS_TEXT,NULL,false);
			InitMultiLine(3,IDS_COLSMART_VIEW,szSmartView,true);
		}
		break;
	case(PT_LONG):
		InitSingleLine(0,IDS_UNSIGNEDDECIMAL,NULL,false);
		InitSingleLine(1,IDS_HEX,NULL,false);
		InitMultiLine(2,IDS_COLSMART_VIEW,NULL,true);
		if (m_lpsInputValue)
		{
			SetStringf(0,_T("%u"),m_lpsInputValue->Value.l); // STRING_OK
			SetHex(1,m_lpsInputValue->Value.l);
			if (szSmartView) SetString(2,szSmartView);
		}
		else
		{
			SetDecimal(0,0);
			SetHex(1,0);
			SetHex(2,0);
		}
		break;
	case(PT_SYSTIME):
		InitSingleLine(0,IDS_LOWDATETIME,NULL,false);
		InitSingleLine(1,IDS_HIGHDATETIME,NULL,false);
		InitSingleLine(2,IDS_DATE,NULL,true);
		if (m_lpsInputValue)
		{
			SetHex(0,(int) m_lpsInputValue->Value.ft.dwLowDateTime);
			SetHex(1,(int) m_lpsInputValue->Value.ft.dwHighDateTime);
			FileTimeToString(&m_lpsInputValue->Value.ft,&szTemp1,NULL);
			SetString(2,szTemp1);
		}
		else
		{
			SetHex(0,0);
			SetHex(1,0);
		}
		break;
	case(PT_CLSID):
		{
			InitSingleLine(0,IDS_GUID,NULL,false);
			LPTSTR szGuid = NULL;
			if (m_lpsInputValue)
			{
				szGuid = GUIDToStringAndName(m_lpsInputValue->Value.lpguid);
			}
			else
			{
				szGuid = GUIDToStringAndName(0);
			}
			SetString(0,szGuid);
			delete[] szGuid;
		}
		break;
	case(PT_SRESTRICTION):
		InitMultiLine(0,IDS_RESTRICTION,NULL,true);
		InterpretProp(m_lpsInputValue,&szTemp1,NULL);
		SetString(0,szTemp1);
		break;
	case(PT_ACTIONS):
		InitMultiLine(0,IDS_ACTIONS,NULL,true);
		InterpretProp(m_lpsInputValue,&szTemp1,NULL);
		SetString(0,szTemp1);
		break;
	default:
		InterpretProp(m_lpsInputValue,&szTemp1,&szTemp2);
		InitMultiLine(0,IDS_VALUE,szTemp1,true);
		InitMultiLine(1,IDS_ALTERNATEVIEW,szTemp2,true);
		break;
	}
	delete[] szSmartView;
} // CPropertyEditor::InitPropertyControls

void CPropertyEditor::WriteStringsToSPropValue()
{
	HRESULT hRes = S_OK;
	CString szTmpString;
	ULONG ulStrLen = NULL;

	// Check first if we'll have anything to write
	switch (PROP_TYPE(m_ulPropTag))
	{
	case(PT_OBJECT): // Nothing to write back - not supported
	case(PT_SRESTRICTION):
	case(PT_ACTIONS):
		return;
	case (PT_BINARY):
		// Check that we've got valid hex string before we allocate anything. Note that we're
		// reading szTmpString now and will assume it's read when we get to the real PT_BINARY case
		szTmpString = GetStringUseControl(1);

		// remove any whitespace before decoding
		CleanHexString(&szTmpString);

		ulStrLen = szTmpString.GetLength();
		if (ulStrLen & 1) return; // can't use an odd length string
		break;
	case (PT_STRING8):
	case (PT_UNICODE):
		// Check that we've got valid hex string before we allocate anything. Note that we're
		// reading szTmpString now and will assume it's read when we get to the real PT_STRING8/PT_UNICODE cases
		szTmpString = GetStringUseControl(2);

		// remove any whitespace before decoding
		CleanHexString(&szTmpString);

		ulStrLen = szTmpString.GetLength();
		if (ulStrLen & 1) return; // can't use an odd length string
		break;
	default: break;
	}

	// If nothing has changed, we're done.
	if (!m_bDirty) return;

	if (!m_lpsOutputValue)
	{
		if (m_lpAllocParent)
		{
			EC_H(MAPIAllocateMore(
				sizeof(SPropValue),
				m_lpAllocParent,
				(LPVOID*)&m_lpsOutputValue));
		}
		else
		{
			EC_H(MAPIAllocateBuffer(
				sizeof(SPropValue),
				(LPVOID*) &m_lpsOutputValue));
			m_lpAllocParent = m_lpsOutputValue;
		}
	}

	if (m_lpsOutputValue)
	{
		BOOL bFailed = false; // set true if we fail to get a prop and have to clean up memory
		m_lpsOutputValue->ulPropTag = m_ulPropTag;
		m_lpsOutputValue->dwAlignPad = NULL;
		switch (PROP_TYPE(m_ulPropTag))
		{
		case(PT_I2): // treat as signed long
			{
				short int iVal = 0;
				szTmpString = GetStringUseControl(0);
				iVal = (short int) _tcstol(szTmpString,NULL,10);
				m_lpsOutputValue->Value.i = iVal;
			}
			break;
		case(PT_LONG): // treat as unsigned long
			{
				LONG lVal = 0;
				szTmpString = GetStringUseControl(0);
				lVal = (LONG) _tcstoul(szTmpString,NULL,10);
				m_lpsOutputValue->Value.l = lVal;
			}
			break;
		case(PT_R4):
			{
				float fVal = 0;
				szTmpString = GetStringUseControl(0);
				fVal = (float) _tcstod(szTmpString,NULL);
				m_lpsOutputValue->Value.flt = fVal;
			}
			break;
		case(PT_DOUBLE):
			{
				double dVal = 0;
				szTmpString = GetStringUseControl(0);
				dVal = (double) _tcstod(szTmpString,NULL);
				m_lpsOutputValue->Value.dbl = dVal;
			}
			break;
		case(PT_CURRENCY):
			{
				CURRENCY curVal = {0};
				szTmpString = GetStringUseControl(0);
				curVal.Hi = _tcstoul(szTmpString,NULL,16);
				szTmpString = GetStringUseControl(1);
				curVal.Lo = _tcstoul(szTmpString,NULL,16);
				m_lpsOutputValue->Value.cur = curVal;
			}
			break;
		case(PT_APPTIME):
			{
				double atVal = 0;
				szTmpString = GetStringUseControl(0);
				atVal = (double) _tcstod(szTmpString,NULL);
				m_lpsOutputValue->Value.at = atVal;
			}
			break;
		case(PT_ERROR): // unsigned
			{
				SCODE errVal = 0;
				szTmpString = GetStringUseControl(0);
				errVal = (SCODE) _tcstoul(szTmpString,NULL,0);
				m_lpsOutputValue->Value.err = errVal;
			}
			break;
		case(PT_BOOLEAN):
			{
				BOOL bVal = false;
				bVal = GetCheckUseControl(0);
				m_lpsOutputValue->Value.b = (unsigned short) bVal;
			}
			break;
		case(PT_I8):
			{
				LARGE_INTEGER liVal = {0};
				szTmpString = GetStringUseControl(0);
				liVal.HighPart = (long) _tcstoul(szTmpString,NULL,0);
				szTmpString = GetStringUseControl(1);
				liVal.LowPart = (long) _tcstoul(szTmpString,NULL,0);
				m_lpsOutputValue->Value.li = liVal;
			}
			break;
		case(PT_STRING8):
			{
				// We read strings out of the hex control in order to preserve any hex level tweaks the user
				// may have done. The RichEdit control likes throwing them away.
				m_lpsOutputValue->Value.lpszA = NULL;

				// remember we already read szTmpString and ulStrLen and found ulStrLen was even
				ULONG cbBin = ulStrLen / 2;
				ULONG cchString = cbBin;
				EC_H(MAPIAllocateMore(
					(cchString+1)*sizeof(CHAR), // NULL terminator
					m_lpAllocParent,
					(LPVOID*)&m_lpsOutputValue->Value.lpszA));
				if (FAILED(hRes)) bFailed = true;
				else
				{
					if (cbBin) MyBinFromHex(
						(LPCTSTR) szTmpString,
						(LPBYTE) m_lpsOutputValue->Value.lpszA,
						cbBin);
					m_lpsOutputValue->Value.lpszA[cchString] = NULL;
				}
			}
			break;
		case(PT_UNICODE):
			{
				// We read strings out of the hex control in order to preserve any hex level tweaks the user
				// may have done. The RichEdit control likes throwing them away.
				m_lpsOutputValue->Value.lpszW = NULL;

				// remember we already read szTmpString and ulStrLen and found ulStrLen was even
				ULONG cbBin = ulStrLen / 2;
				// Since this is a unicode string, cbBin must also be even
				if (cbBin & 1) bFailed = true;
				else
				{
					ULONG cchString = cbBin / 2;
					EC_H(MAPIAllocateMore(
						(cchString+1)*sizeof(WCHAR), // NULL terminator
						m_lpAllocParent,
						(LPVOID*)&m_lpsOutputValue->Value.lpszW));
					if (FAILED(hRes)) bFailed = true;
					else
					{
						if (cbBin) MyBinFromHex(
							(LPCTSTR) szTmpString,
							(LPBYTE) m_lpsOutputValue->Value.lpszW,
							cbBin);
						m_lpsOutputValue->Value.lpszW[cchString] = NULL;
					}
				}
			}
			break;
		case(PT_SYSTIME):
			{
				FILETIME ftVal = {0};
				szTmpString = GetStringUseControl(0);
				ftVal.dwLowDateTime = _tcstoul(szTmpString,NULL,16);
				szTmpString = GetStringUseControl(1);
				ftVal.dwHighDateTime = _tcstoul(szTmpString,NULL,16);
				m_lpsOutputValue->Value.ft = ftVal;
			}
			break;
		case(PT_CLSID):
			{
				EC_H(MAPIAllocateMore(
					sizeof(GUID),
					m_lpAllocParent,
					(LPVOID*)&m_lpsOutputValue->Value.lpguid));
				if (m_lpsOutputValue->Value.lpguid)
				{
					szTmpString = GetStringUseControl(0);
					EC_H(StringToGUID((LPCTSTR) szTmpString,m_lpsOutputValue->Value.lpguid));
					if (FAILED(hRes)) bFailed = true;
				}
			}
			break;
		case(PT_BINARY):
			{
				// remember we already read szTmpString and ulStrLen and found ulStrLen was even
				ULONG cbBin = ulStrLen / 2;
				m_lpsOutputValue->Value.bin.cb = cbBin;
				if (0 == m_lpsOutputValue->Value.bin.cb)
				{
					m_lpsOutputValue->Value.bin.lpb = 0;
				}
				else
				{
					EC_H(MAPIAllocateMore(
						m_lpsOutputValue->Value.bin.cb,
						m_lpAllocParent,
						(LPVOID*)&m_lpsOutputValue->Value.bin.lpb));
					if (FAILED(hRes)) bFailed = true;
					else
					{
						MyBinFromHex(
							(LPCTSTR) szTmpString,
							m_lpsOutputValue->Value.bin.lpb,
							m_lpsOutputValue->Value.bin.cb);
					}
				}
			}
			break;
		default:
			// We shouldn't ever get here unless some new prop type shows up
			bFailed = true;
			break;
		}
		if (bFailed)
		{
			// If we don't have a parent or we are the parent, then we can free here
			if (!m_lpAllocParent || m_lpAllocParent == m_lpsOutputValue)
			{
				MAPIFreeBuffer(m_lpsOutputValue);
				m_lpsOutputValue = NULL;
				m_lpAllocParent = NULL;
			}
			else
			{
				// If m_lpsOutputValue was allocated off a parent, we can't free it here
				// Just drop the reference and m_lpAllocParent's free will clean it up
				m_lpsOutputValue = NULL;
			}
		}
	}
} // CPropertyEditor::WriteStringsToSPropValue

void CPropertyEditor::WriteSPropValueToObject()
{
	if (!m_lpsOutputValue || !m_lpMAPIProp) return;

	HRESULT hRes = S_OK;

	LPSPropProblemArray lpProblemArray = NULL;

	EC_H(m_lpMAPIProp->SetProps(
		1,
		m_lpsOutputValue,
		&lpProblemArray));

	EC_PROBLEMARRAY(lpProblemArray);
	MAPIFreeBuffer(lpProblemArray);

	EC_H(m_lpMAPIProp->SaveChanges(KEEP_OPEN_READWRITE));
} // CPropertyEditor::WriteSPropValueToObject

// Callers beware: Detatches and returns the modified prop value - this must be MAPIFreeBuffered!
_Check_return_ LPSPropValue CPropertyEditor::DetachModifiedSPropValue()
{
	LPSPropValue m_lpRet = m_lpsOutputValue;
	m_lpsOutputValue = NULL;
	return m_lpRet;
} // CPropertyEditor::DetachModifiedSPropValue

_Check_return_ ULONG CPropertyEditor::HandleChange(UINT nID)
{
	ULONG i = CEditor::HandleChange(nID);

	if ((ULONG) -1 == i) return (ULONG) -1;

	CString szTmpString;

	// If we get here, something changed - set the dirty flag
	m_bDirty = true;

	switch (PROP_TYPE(m_ulPropTag))
	{
	case(PT_I2): // signed 16 bit
		{
			short int iVal = 0;
			szTmpString = GetStringUseControl(i);
			if (0 == i)
			{
				iVal = (short int) _tcstol(szTmpString,NULL,10);
				SetHex(1, iVal);
			}
			else if (1 == i)
			{
				iVal = (short int) _tcstol(szTmpString,NULL,16);
				SetDecimal(0,iVal);
			}

			LPTSTR szSmartView = NULL;
			SPropValue sProp = {0};
			sProp.ulPropTag = m_ulPropTag;
			sProp.Value.i = iVal;

			InterpretPropSmartView(&sProp,
				m_lpMAPIProp,
				NULL,
				NULL,
				m_bMVRow,
				&szSmartView);

			if (szSmartView) SetString(2,szSmartView);
			delete[] szSmartView;
			szSmartView = NULL;
		}
		break;
	case(PT_LONG): // unsigned 32 bit
		{
			LONG lVal = 0;
			szTmpString = GetStringUseControl(i);
			if (0 == i)
			{
				lVal = (LONG) _tcstoul(szTmpString,NULL,10);
				SetHex(1,lVal);
			}
			else if (1 == i)
			{
				lVal = (LONG) _tcstoul(szTmpString,NULL,16);
				SetStringf(0,_T("%u"),lVal); // STRING_OK
			}

			LPTSTR szSmartView = NULL;
			SPropValue sProp = {0};
			sProp.ulPropTag = m_ulPropTag;
			sProp.Value.l = lVal;

			InterpretPropSmartView(&sProp,
				m_lpMAPIProp,
				NULL,
				NULL,
				m_bMVRow,
				&szSmartView);

			if (szSmartView) SetString(2,szSmartView);
			delete[] szSmartView;
			szSmartView = NULL;
		}
		break;
	case(PT_CURRENCY):
		{
			CURRENCY curVal = {0};
			if (0 == i || 1 == i)
			{
				szTmpString = GetStringUseControl(0);
				curVal.Hi = _tcstoul(szTmpString,NULL,16);
				szTmpString = GetStringUseControl(1);
				curVal.Lo= _tcstoul(szTmpString,NULL,16);
				SetString(2,CurrencyToString(curVal));
			}
			else if (2 == i)
			{
				szTmpString = GetStringUseControl(i);
				szTmpString.Remove(_T('.'));
				curVal.int64 = _ttoi64(szTmpString);
				SetHex(0,(int) curVal.Hi);
				SetHex(1,(int) curVal.Lo);
			}
		}
		break;
	case(PT_I8):
		{
			LARGE_INTEGER liVal = {0};
			if (0 == i || 1 == i)
			{
				szTmpString = GetStringUseControl(0);
				liVal.HighPart = (long) _tcstoul(szTmpString,NULL,0);
				szTmpString = GetStringUseControl(1);
				liVal.LowPart = (long) _tcstoul(szTmpString,NULL,0);
				SetStringf(2,_T("%I64d"),liVal.QuadPart); // STRING_OK
			}
			else if (2 == i)
			{
				szTmpString = GetStringUseControl(i);
				liVal.QuadPart = _ttoi64(szTmpString);
				SetHex(0,(int) liVal.HighPart);
				SetHex(1,(int) liVal.LowPart);
			}
		}
		break;
	case(PT_SYSTIME): // components are unsigned hex
		{
			FILETIME ftVal = {0};
			szTmpString = GetStringUseControl(0);
			ftVal.dwLowDateTime = _tcstoul(szTmpString,NULL,16);
			szTmpString = GetStringUseControl(1);
			ftVal.dwHighDateTime = _tcstoul(szTmpString,NULL,16);

			FileTimeToString(&ftVal,&szTmpString,NULL);
			SetString(2,szTmpString);
		}
		break;
	case(PT_BINARY):
		{
			LPBYTE	lpb = NULL;
			size_t	cb = 0;
			SBinary Bin = {0};

			if (1 == i)
			{
				if (GetBinaryUseControl(1,&cb,&lpb))
				{
					// Treat as a NULL terminated string
					// GetBinaryUseControl includes extra NULLs at the end of the buffer to make this work
					SetStringA(2,(LPCSTR) lpb, cb+1); // ansi string
					Bin.lpb = lpb;
				}
			}
			else if (2 == i)
			{
				size_t cchStr = NULL;
				LPSTR lpszA = GetEditBoxTextA(2, &cchStr);

				// What we just read includes a NULL terminator, in both the string and count.
				// When we write binary, we don't want to include this NULL
				if (cchStr) cchStr -= 1;
				cb = cchStr * sizeof(CHAR);

				SetBinary(1, (LPBYTE) lpszA, cb);
				Bin.lpb = (LPBYTE) lpszA;
			}

			Bin.cb = (ULONG) cb;
			SetSize(0, cb);

			LPTSTR szSmartView = NULL;
			SPropValue sProp = {0};
			sProp.ulPropTag = m_ulPropTag;
			sProp.Value.bin = Bin;

			InterpretPropSmartView(&sProp,
				m_lpMAPIProp,
				NULL,
				NULL,
				m_bMVRow,
				&szSmartView);

			SetString(3,szSmartView);
			delete[] szSmartView;
			szSmartView = NULL;

			delete[] lpb;
		}
		break;
	case(PT_STRING8):
		{
			if (0 == i)
			{
				size_t cbStr = 0;
				size_t cchStr = 0;
				LPSTR lpszA = GetEditBoxTextA(0,&cchStr);

				// What we just read includes a NULL terminator, in both the string and count.
				// When we write binary, we don't want to include this NULL
				if (cchStr) cchStr -= 1;
				cbStr = cchStr * sizeof(CHAR);

				// Even if we don't have a string, still make the call to SetBinary
				// This will blank out the binary control when lpszA is NULL
				SetBinary(2,(LPBYTE) lpszA, cbStr);

				SetSize(1, cbStr);
			}
			else if (2 == i)
			{
				LPBYTE	lpb = NULL;
				size_t	cb = 0;

				if (GetBinaryUseControl(2,&cb,&lpb))
				{
					// GetBinaryUseControl includes extra NULLs at the end of the buffer to make this work
					SetStringA(0,(LPCSTR) lpb, cb+1);
					SetSize(1, cb);
				}
				delete[] lpb;
			}
		}
		break;
	case(PT_UNICODE):
		{
			if (0 == i)
			{
				size_t cbStr = 0;
				size_t cchStr = 0;
				LPWSTR lpszW = GetEditBoxTextW(0,&cchStr);

				// What we just read includes a NULL terminator, in both the string and count.
				// When we write binary, we don't want to include this NULL
				if (cchStr) cchStr -= 1;
				cbStr = cchStr * sizeof(WCHAR);

				// Even if we don't have a string, still make the call to SetBinary
				// This will blank out the binary control when lpszW is NULL
				SetBinary(2,(LPBYTE) lpszW, cbStr);

				SetSize(1, cbStr);
			}
			else if (2 == i)
			{
				LPBYTE	lpb = NULL;
				size_t	cb = 0;

				if (GetBinaryUseControl(2,&cb,&lpb))
				{
					if (!(cb % sizeof(WCHAR)))
					{
						// GetBinaryUseControl includes extra NULLs at the end of the buffer to make this work
						SetStringW(0,(LPCWSTR) lpb, cb+1);
						SetSize(1, cb);
					}
				}
				delete[] lpb;
			}
		}
		break;
	default:
		break;
	}
	return i;
} // CPropertyEditor::HandleChange

static TCHAR* MVCLASS = _T("CMultiValuePropertyEditor"); // STRING_OK

// Create an editor for a MAPI property
CMultiValuePropertyEditor::CMultiValuePropertyEditor(
	_In_ CWnd* pParentWnd,
	UINT uidTitle,
	UINT uidPrompt,
	BOOL bIsAB,
	_In_opt_ LPVOID lpAllocParent,
	_In_opt_ LPMAPIPROP lpMAPIProp,
	ULONG ulPropTag,
	_In_opt_ LPSPropValue lpsPropValue):
CEditor(pParentWnd,uidTitle,uidPrompt,0,CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL)
{
	TRACE_CONSTRUCTOR(MVCLASS);

	m_bIsAB = bIsAB;
	m_lpAllocParent = lpAllocParent;
	m_lpsOutputValue = NULL;

	m_lpMAPIProp = lpMAPIProp;
	if (m_lpMAPIProp) m_lpMAPIProp->AddRef();
	m_ulPropTag = ulPropTag;
	m_lpsInputValue = lpsPropValue;

	CString szPromptPostFix;
	szPromptPostFix.Format(_T("\r\n%s"),(LPCTSTR) TagToString(m_ulPropTag,m_lpMAPIProp,m_bIsAB,false)); // STRING_OK

	SetPromptPostFix(szPromptPostFix);

	// Let's crack our property open and see what kind of controls we'll need for it
	CreatePropertyControls();

	InitPropertyControls();
} // CMultiValuePropertyEditor::CMultiValuePropertyEditor

CMultiValuePropertyEditor::~CMultiValuePropertyEditor()
{
	TRACE_DESTRUCTOR(MVCLASS);
	if (m_lpMAPIProp) m_lpMAPIProp->Release();
} // CMultiValuePropertyEditor::~CMultiValuePropertyEditor

_Check_return_ BOOL CMultiValuePropertyEditor::OnInitDialog()
{
	BOOL bRet = CEditor::OnInitDialog();

	ReadMultiValueStringsFromProperty(0);
	ResizeList(0,false);
	UpdateSmartView(0);
	UpdateListButtons();

	return bRet;
} // CMultiValuePropertyEditor::OnInitDialog

void CMultiValuePropertyEditor::OnOK()
{
	// This is where we write our changes back
	WriteMultiValueStringsToSPropValue(0);
	WriteSPropValueToObject();
	CDialog::OnOK(); // don't need to call CEditor::OnOK
} // CMultiValuePropertyEditor::OnOK

void CMultiValuePropertyEditor::CreatePropertyControls()
{
	if (PT_MV_BINARY == PROP_TYPE(m_ulPropTag) ||
		PT_MV_LONG == PROP_TYPE(m_ulPropTag))
	{
		CreateControls(2);
	}
	else
	{
		CreateControls(1);
	}
} // CMultiValuePropertyEditor::CreatePropertyControls

void CMultiValuePropertyEditor::InitPropertyControls()
{
	InitList(0,IDS_PROPVALUES,false,false);
	if (PT_MV_BINARY == PROP_TYPE(m_ulPropTag) ||
		PT_MV_LONG == PROP_TYPE(m_ulPropTag))
	{
		InitMultiLine(1,IDS_COLSMART_VIEW,NULL,true);
	}
} // CMultiValuePropertyEditor::InitPropertyControls

// Function must be called AFTER dialog controls have been created, not before
void CMultiValuePropertyEditor::ReadMultiValueStringsFromProperty(ULONG ulListNum)
{
	if (!IsValidList(ulListNum)) return;

	InsertColumn(ulListNum,0,IDS_ENTRY);
	InsertColumn(ulListNum,1,IDS_VALUE);
	InsertColumn(ulListNum,2,IDS_ALTERNATEVIEW);
	if (PT_MV_LONG == PROP_TYPE(m_ulPropTag) ||
		PT_MV_BINARY == PROP_TYPE(m_ulPropTag))
	{
		InsertColumn(ulListNum,3,IDS_COLSMART_VIEW);
	}

	if (!m_lpsInputValue) return;
	if (!(PROP_TYPE(m_lpsInputValue->ulPropTag) & MV_FLAG)) return;

	CString szTmp;
	ULONG iMVCount = 0;
	// All the MV structures are basically the same, so we can cheat when we pull the count
	ULONG cValues = m_lpsInputValue->Value.MVi.cValues;
	for (iMVCount = 0; iMVCount < cValues; iMVCount++)
	{
		szTmp.Format(_T("%d"),iMVCount); // STRING_OK
		SortListData* lpData = InsertListRow(ulListNum,iMVCount,szTmp);

		if (lpData)
		{
			lpData->ulSortDataType = SORTLIST_MVPROP;
			switch(PROP_TYPE(m_lpsInputValue->ulPropTag))
			{
			case(PT_MV_I2):
				lpData->data.MV.val.i = m_lpsInputValue->Value.MVi.lpi[iMVCount];
				break;
			case(PT_MV_LONG):
				lpData->data.MV.val.l = m_lpsInputValue->Value.MVl.lpl[iMVCount];
				break;
			case(PT_MV_DOUBLE):
				lpData->data.MV.val.dbl = m_lpsInputValue->Value.MVdbl.lpdbl[iMVCount];
				break;
			case(PT_MV_CURRENCY):
				lpData->data.MV.val.cur = m_lpsInputValue->Value.MVcur.lpcur[iMVCount];
				break;
			case(PT_MV_APPTIME):
				lpData->data.MV.val.at = m_lpsInputValue->Value.MVat.lpat[iMVCount];
				break;
			case(PT_MV_SYSTIME):
				lpData->data.MV.val.ft = m_lpsInputValue->Value.MVft.lpft[iMVCount];
				break;
			case(PT_MV_I8):
				lpData->data.MV.val.li = m_lpsInputValue->Value.MVli.lpli[iMVCount];
				break;
			case(PT_MV_R4):
				lpData->data.MV.val.flt = m_lpsInputValue->Value.MVflt.lpflt[iMVCount];
				break;
			case(PT_MV_STRING8):
				lpData->data.MV.val.lpszA = m_lpsInputValue->Value.MVszA.lppszA[iMVCount];
				break;
			case(PT_MV_UNICODE):
				lpData->data.MV.val.lpszW = m_lpsInputValue->Value.MVszW.lppszW[iMVCount];
				break;
			case(PT_MV_BINARY):
				lpData->data.MV.val.bin = m_lpsInputValue->Value.MVbin.lpbin[iMVCount];
				break;
			case(PT_MV_CLSID):
				lpData->data.MV.val.lpguid = &m_lpsInputValue->Value.MVguid.lpguid[iMVCount];
				break;
			default:
				break;
			}

			SPropValue sProp = {0};
			sProp.ulPropTag = CHANGE_PROP_TYPE(m_lpsInputValue->ulPropTag,PROP_TYPE(m_lpsInputValue->ulPropTag) & ~MV_FLAG);
			sProp.Value = lpData->data.MV.val;
			UpdateListRow(&sProp,ulListNum,iMVCount);

			lpData->bItemFullyLoaded = true;
		}
	}
} // CMultiValuePropertyEditor::ReadMultiValueStringsFromProperty

// Perisist the data in the controls to m_lpsOutputValue
void CMultiValuePropertyEditor::WriteMultiValueStringsToSPropValue(ULONG ulListNum)
{
	if (!IsValidList(ulListNum)) return;

	// If we're not dirty, don't write
	// Unless we had no input value. Then we're creating a new property.
	// So we're implicitly dirty.
	if (!ListDirty(ulListNum) && m_lpsInputValue) return;

	HRESULT hRes = S_OK;
	// Take care of allocations first
	if (!m_lpsOutputValue)
	{
		if (m_lpAllocParent)
		{
			EC_H(MAPIAllocateMore(
				sizeof(SPropValue),
				m_lpAllocParent,
				(LPVOID*) &m_lpsOutputValue));
		}
		else
		{
			EC_H(MAPIAllocateBuffer(
				sizeof(SPropValue),
				(LPVOID*) &m_lpsOutputValue));
			m_lpAllocParent = m_lpsOutputValue;
		}
	}

	if (m_lpsOutputValue)
	{
		WriteMultiValueStringsToSPropValue(ulListNum, (LPVOID) m_lpAllocParent, m_lpsOutputValue);
	}
} // CMultiValuePropertyEditor::WriteMultiValueStringsToSPropValue

// Given a pointer to an SPropValue structure which has already been allocated, fill out the values
void CMultiValuePropertyEditor::WriteMultiValueStringsToSPropValue(ULONG ulListNum, _In_ LPVOID lpParent, _In_ LPSPropValue lpsProp)
{
	if (!lpParent || !lpsProp) return;

	HRESULT hRes = S_OK;
	ULONG ulNumVals = GetListCount(ulListNum);
	ULONG iMVCount = 0;

	lpsProp->ulPropTag = m_ulPropTag;
	lpsProp->dwAlignPad = NULL;

	switch(PROP_TYPE(lpsProp->ulPropTag))
	{
	case(PT_MV_I2):
		EC_H(MAPIAllocateMore(sizeof(short int) * ulNumVals,lpParent,(LPVOID*)&lpsProp->Value.MVi.lpi));
		lpsProp->Value.MVi.cValues = ulNumVals;
		break;
	case(PT_MV_LONG):
		EC_H(MAPIAllocateMore(sizeof(LONG) * ulNumVals,lpParent,(LPVOID*)&lpsProp->Value.MVl.lpl));
		lpsProp->Value.MVl.cValues = ulNumVals;
		break;
	case(PT_MV_DOUBLE):
		EC_H(MAPIAllocateMore(sizeof(double) * ulNumVals,lpParent,(LPVOID*)&lpsProp->Value.MVdbl.lpdbl));
		lpsProp->Value.MVdbl.cValues = ulNumVals;
		break;
	case(PT_MV_CURRENCY):
		EC_H(MAPIAllocateMore(sizeof(CURRENCY) * ulNumVals,lpParent,(LPVOID*)&lpsProp->Value.MVcur.lpcur));
		lpsProp->Value.MVcur.cValues = ulNumVals;
		break;
	case(PT_MV_APPTIME):
		EC_H(MAPIAllocateMore(sizeof(double) * ulNumVals,lpParent,(LPVOID*)&lpsProp->Value.MVat.lpat));
		lpsProp->Value.MVat.cValues = ulNumVals;
		break;
	case(PT_MV_SYSTIME):
		EC_H(MAPIAllocateMore(sizeof(FILETIME) * ulNumVals,lpParent,(LPVOID*)&lpsProp->Value.MVft.lpft));
		lpsProp->Value.MVft.cValues = ulNumVals;
		break;
	case(PT_MV_I8):
		EC_H(MAPIAllocateMore(sizeof(LARGE_INTEGER) * ulNumVals,lpParent,(LPVOID*)&lpsProp->Value.MVli.lpli));
		lpsProp->Value.MVli.cValues = ulNumVals;
		break;
	case(PT_MV_R4):
		EC_H(MAPIAllocateMore(sizeof(float) * ulNumVals,lpParent,(LPVOID*)&lpsProp->Value.MVflt.lpflt));
		lpsProp->Value.MVflt.cValues = ulNumVals;
		break;
	case(PT_MV_STRING8):
		EC_H(MAPIAllocateMore(sizeof(LPSTR) * ulNumVals,lpParent,(LPVOID*)&lpsProp->Value.MVszA.lppszA));
		lpsProp->Value.MVszA.cValues = ulNumVals;
		break;
	case(PT_MV_UNICODE):
		EC_H(MAPIAllocateMore(sizeof(LPWSTR) * ulNumVals,lpParent,(LPVOID*)&lpsProp->Value.MVszW.lppszW));
		lpsProp->Value.MVszW.cValues = ulNumVals;
		break;
	case(PT_MV_BINARY):
		EC_H(MAPIAllocateMore(sizeof(SBinary) * ulNumVals,lpParent,(LPVOID*)&lpsProp->Value.MVbin.lpbin));
		lpsProp->Value.MVbin.cValues = ulNumVals;
		break;
	case(PT_MV_CLSID):
		EC_H(MAPIAllocateMore(sizeof(GUID) * ulNumVals,lpParent,(LPVOID*)&lpsProp->Value.MVguid.lpguid));
		lpsProp->Value.MVguid.cValues = ulNumVals;
		break;
	default:
		break;
	}
	// Allocation is now done

	// Now write our data into the space we allocated
	for (iMVCount = 0; iMVCount < ulNumVals; iMVCount++)
	{
		SortListData* lpData = GetListRowData(ulListNum,iMVCount);

		if (lpData)
		{
			switch(PROP_TYPE(lpsProp->ulPropTag))
			{
			case(PT_MV_I2):
				lpsProp->Value.MVi.lpi[iMVCount] = lpData->data.MV.val.i;
				break;
			case(PT_MV_LONG):
				lpsProp->Value.MVl.lpl[iMVCount] = lpData->data.MV.val.l;
				break;
			case(PT_MV_DOUBLE):
				lpsProp->Value.MVdbl.lpdbl[iMVCount] = lpData->data.MV.val.dbl;
				break;
			case(PT_MV_CURRENCY):
				lpsProp->Value.MVcur.lpcur[iMVCount] = lpData->data.MV.val.cur;
				break;
			case(PT_MV_APPTIME):
				lpsProp->Value.MVat.lpat[iMVCount] = lpData->data.MV.val.at;
				break;
			case(PT_MV_SYSTIME):
				lpsProp->Value.MVft.lpft[iMVCount] = lpData->data.MV.val.ft;
				break;
			case(PT_MV_I8):
				lpsProp->Value.MVli.lpli[iMVCount] = lpData->data.MV.val.li;
				break;
			case(PT_MV_R4):
				lpsProp->Value.MVflt.lpflt[iMVCount] = lpData->data.MV.val.flt;
				break;
			case(PT_MV_STRING8):
				EC_H(CopyStringA(&lpsProp->Value.MVszA.lppszA[iMVCount],lpData->data.MV.val.lpszA,lpParent));
				break;
			case(PT_MV_UNICODE):
				EC_H(CopyStringW(&lpsProp->Value.MVszW.lppszW[iMVCount],lpData->data.MV.val.lpszW,lpParent));
				break;
			case(PT_MV_BINARY):
				EC_H(CopySBinary(&lpsProp->Value.MVbin.lpbin[iMVCount],&lpData->data.MV.val.bin,lpParent));
				break;
			case(PT_MV_CLSID):
				if (lpData->data.MV.val.lpguid)
				{
					lpsProp->Value.MVguid.lpguid[iMVCount] = *lpData->data.MV.val.lpguid;
				}

				break;
			default:
				break;
			}
		}
	}
} // CMultiValuePropertyEditor::WriteMultiValueStringsToSPropValue

void CMultiValuePropertyEditor::WriteSPropValueToObject()
{
	if (!m_lpsOutputValue || !m_lpMAPIProp) return;

	HRESULT hRes = S_OK;

	LPSPropProblemArray lpProblemArray = NULL;

	EC_H(m_lpMAPIProp->SetProps(
		1,
		m_lpsOutputValue,
		&lpProblemArray));

	EC_PROBLEMARRAY(lpProblemArray);
	MAPIFreeBuffer(lpProblemArray);

	EC_H(m_lpMAPIProp->SaveChanges(KEEP_OPEN_READWRITE));
} // CMultiValuePropertyEditor::WriteSPropValueToObject

// Callers beware: Detatches and returns the modified prop value - this must be MAPIFreeBuffered!
_Check_return_ LPSPropValue CMultiValuePropertyEditor::DetachModifiedSPropValue()
{
	LPSPropValue m_lpRet = m_lpsOutputValue;
	m_lpsOutputValue = NULL;
	return m_lpRet;
} // CMultiValuePropertyEditor::DetachModifiedSPropValue

_Check_return_ BOOL CMultiValuePropertyEditor::DoListEdit(ULONG ulListNum, int iItem, _In_ SortListData* lpData)
{
	if (!lpData) return false;
	if (!IsValidList(ulListNum)) return false;

	HRESULT hRes = S_OK;
	SPropValue tmpPropVal = {0};
	// Strip off MV_FLAG since we're displaying only a row
	tmpPropVal.ulPropTag = m_ulPropTag & ~MV_FLAG;
	tmpPropVal.Value = lpData->data.MV.val;

	LPSPropValue lpNewValue = NULL;
	WC_H(DisplayPropertyEditor(
		this,
		IDS_EDITROW,
		IDS_EDITROWPROMPT,
		m_bIsAB,
		NULL, // not passing an allocation parent because we know we're gonna free the result
		m_lpMAPIProp,
		NULL,
		true, // This is a row from a multivalued property. Only case we pass true here.
		&tmpPropVal,
		&lpNewValue));

	if (S_OK == hRes && lpNewValue)
	{
		ULONG ulBufSize = 0;
		// This handles most cases by default - cases needing a buffer copied are handled below
		lpData->data.MV.val = lpNewValue->Value;
		switch (PROP_TYPE(lpNewValue->ulPropTag))
		{
		case(PT_STRING8):
			{
				// When the lpData is ultimately freed, MAPI will take care of freeing this.
				// This will be true even if we do this multiple times off the same lpData!
				size_t cbStr = 0;
				EC_H(StringCbLengthA(lpNewValue->Value.lpszA,STRSAFE_MAX_CCH * sizeof(char),&cbStr));
				cbStr += sizeof(char);

				EC_H(MAPIAllocateMore(
					(ULONG) cbStr,
					lpData,
					(LPVOID*) &lpData->data.MV.val.lpszA));

				if (S_OK == hRes)
				{
					memcpy(lpData->data.MV.val.lpszA,lpNewValue->Value.lpszA,cbStr);
				}
			}
			break;
		case(PT_UNICODE):
			{
				size_t cbStr = 0;
				EC_H(StringCbLengthW(lpNewValue->Value.lpszW,STRSAFE_MAX_CCH * sizeof(WCHAR),&cbStr));
				cbStr += sizeof(WCHAR);

				EC_H(MAPIAllocateMore(
					(ULONG) cbStr,
					lpData,
					(LPVOID*) &lpData->data.MV.val.lpszW));

				if (S_OK == hRes)
				{
					memcpy(lpData->data.MV.val.lpszW,lpNewValue->Value.lpszW,cbStr);
				}
			}
			break;
		case(PT_BINARY):
			ulBufSize = lpNewValue->Value.bin.cb;
			EC_H(MAPIAllocateMore(
				ulBufSize,
				lpData,
				(LPVOID*) &lpData->data.MV.val.bin.lpb));

			if (S_OK == hRes)
			{
				memcpy(lpData->data.MV.val.bin.lpb,lpNewValue->Value.bin.lpb,ulBufSize);
			}
			break;
		case(PT_CLSID):
			ulBufSize = sizeof(GUID);
			EC_H(MAPIAllocateMore(
				ulBufSize,
				lpData,
				(LPVOID*) &lpData->data.MV.val.lpguid));

			if (S_OK == hRes)
			{
				memcpy(lpData->data.MV.val.lpguid,lpNewValue->Value.lpguid,ulBufSize);
			}
			break;
		default:
			break;
		}

		// update the UI
		UpdateListRow(lpNewValue,ulListNum,iItem);
		UpdateSmartView(ulListNum);
		return true;
	}

	// Remember we didn't have an allocation parent - this is safe
	MAPIFreeBuffer(lpNewValue);
	return false;
} // CMultiValuePropertyEditor::DoListEdit

void CMultiValuePropertyEditor::UpdateListRow(_In_ LPSPropValue lpProp, ULONG ulListNum, ULONG iMVCount)
{
	CString szTmp;
	CString szAltTmp;

	InterpretProp(lpProp,&szTmp,&szAltTmp);
	SetListString(ulListNum,iMVCount,1,szTmp);
	SetListString(ulListNum,iMVCount,2,szAltTmp);

	if (PT_MV_LONG == PROP_TYPE(m_ulPropTag) ||
		PT_MV_BINARY == PROP_TYPE(m_ulPropTag))
	{
		LPTSTR szSmartView = NULL;

		InterpretPropSmartView(lpProp,
			m_lpMAPIProp,
			NULL,
			NULL,
			true,
			&szSmartView);

		if (szSmartView) SetListString(ulListNum,iMVCount,3,szSmartView);
		delete[] szSmartView;
		szSmartView = NULL;
	}
} // CMultiValuePropertyEditor::UpdateListRow

void CMultiValuePropertyEditor::UpdateSmartView(ULONG ulListNum)
{
	HRESULT hRes = S_OK;
	LPSPropValue lpsProp = NULL;
	EC_H(MAPIAllocateBuffer(
		sizeof(SPropValue),
		(LPVOID*) &lpsProp));
	if (lpsProp)
	{
		WriteMultiValueStringsToSPropValue(ulListNum, (LPVOID) lpsProp, lpsProp);

		LPTSTR szSmartView = NULL;
		InterpretPropSmartView(lpsProp,
			m_lpMAPIProp,
			NULL,
			NULL,
			true,
			&szSmartView);
		if (szSmartView) SetString(1,szSmartView);
		delete[] szSmartView;
	}
	MAPIFreeBuffer(lpsProp);
} // CMultiValuePropertyEditor::UpdateSmartView