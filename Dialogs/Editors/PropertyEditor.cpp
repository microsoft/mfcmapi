#include "stdafx.h"
#include "PropertyEditor.h"
#include "InterpretProp2.h"
#include "MAPIFunctions.h"
#include "SmartView/SmartView.h"
#include "SortList/MVPropData.h"
#include "ViewPane/CountedTextPane.h"

_Check_return_ HRESULT DisplayPropertyEditor(_In_ CWnd* pParentWnd,
	UINT uidTitle,
	UINT uidPrompt,
	bool bIsAB,
	_In_opt_ LPVOID lpAllocParent,
	_In_opt_ LPMAPIPROP lpMAPIProp,
	ULONG ulPropTag,
	bool bMVRow,
	_In_opt_ LPSPropValue lpsPropValue,
	_Inout_opt_ LPSPropValue* lpNewValue)
{
	auto hRes = S_OK;
	auto bShouldFreeInputValue = false;

	// We got a MAPI prop object and no input value, go look one up
	if (lpMAPIProp && !lpsPropValue)
	{
		SPropTagArray sTag = { 0 };
		sTag.cValues = 1;
		sTag.aulPropTag[0] = PT_ERROR == PROP_TYPE(ulPropTag) ? CHANGE_PROP_TYPE(ulPropTag, PT_UNSPECIFIED) : ulPropTag;
		ULONG ulValues = NULL;

		WC_MAPI(lpMAPIProp->GetProps(&sTag, NULL, &ulValues, &lpsPropValue));

		// Suppress MAPI_E_NOT_FOUND error when the source type is non error
		if (lpsPropValue &&
			PT_ERROR == PROP_TYPE(lpsPropValue->ulPropTag) &&
			MAPI_E_NOT_FOUND == lpsPropValue->Value.err &&
			PT_ERROR != PROP_TYPE(ulPropTag)
			)
		{
			MAPIFreeBuffer(lpsPropValue);
			lpsPropValue = nullptr;
		}

		if (MAPI_E_CALL_FAILED == hRes)
		{
			// Just suppress this - let the user edit anyway
			hRes = S_OK;
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
}

static wstring SVCLASS = L"CPropertyEditor"; // STRING_OK

// Create an editor for a MAPI property
CPropertyEditor::CPropertyEditor(
	_In_ CWnd* pParentWnd,
	UINT uidTitle,
	UINT uidPrompt,
	bool bIsAB,
	bool bMVRow,
	_In_opt_ LPVOID lpAllocParent,
	_In_opt_ LPMAPIPROP lpMAPIProp,
	ULONG ulPropTag,
	_In_opt_ LPSPropValue lpsPropValue) :
	CEditor(pParentWnd, uidTitle, uidPrompt, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL)
{
	TRACE_CONSTRUCTOR(SVCLASS);

	m_bIsAB = bIsAB;
	m_bMVRow = bMVRow;
	m_lpAllocParent = lpAllocParent;
	m_lpsOutputValue = nullptr;
	m_bDirty = false;

	m_lpMAPIProp = lpMAPIProp;
	if (m_lpMAPIProp) m_lpMAPIProp->AddRef();
	m_ulPropTag = ulPropTag;
	m_lpsInputValue = lpsPropValue;
	m_lpSmartView = nullptr;

	// If we didn't have an input value, we are creating a new property
	// So by definition, we're already dirty
	if (!m_lpsInputValue) m_bDirty = true;

	auto szPromptPostFix = format(L"%ws%ws", uidPrompt ? L"\r\n" : L"", TagToString(m_ulPropTag | (m_bMVRow ? MV_FLAG : NULL), m_lpMAPIProp, m_bIsAB, false).c_str()); // STRING_OK
	SetPromptPostFix(szPromptPostFix);

	InitPropertyControls();
}

CPropertyEditor::~CPropertyEditor()
{
	TRACE_DESTRUCTOR(SVCLASS);
	if (m_lpMAPIProp) m_lpMAPIProp->Release();
}

BOOL CPropertyEditor::OnInitDialog()
{
	auto bRet = CEditor::OnInitDialog();
	return bRet;
}

void CPropertyEditor::OnOK()
{
	// This is where we write our changes back
	WriteStringsToSPropValue();

	// Write the property to the object if we're not editing a row of a MV property
	if (!m_bMVRow) WriteSPropValueToObject();
	CMyDialog::OnOK(); // don't need to call CEditor::OnOK
}

void CPropertyEditor::InitPropertyControls()
{
	switch (PROP_TYPE(m_ulPropTag))
	{
	case PT_I8:
	case PT_I2:
	case PT_BINARY:
	case PT_LONG:
		// This will be freed by the pane that we pass it to.
		m_lpSmartView = static_cast<SmartViewPane*>(SmartViewPane::Create(IDS_SMARTVIEW));
	}

	auto iStructType = FindSmartViewParserForProp(m_lpsInputValue ? m_lpsInputValue->ulPropTag : m_ulPropTag, NULL, nullptr, m_bMVRow);
	auto szSmartView = InterpretPropSmartView(
		m_lpsInputValue,
		m_lpMAPIProp,
		nullptr,
		nullptr,
		m_bIsAB,
		m_bMVRow); // Built from lpProp & lpMAPIProp

	wstring szTemp1;
	wstring szTemp2;
	CountedTextPane* lpPane = nullptr;
	size_t cbStr = 0;
	wstring szGuid;

	switch (PROP_TYPE(m_ulPropTag))
	{
	case PT_APPTIME:
		InitPane(0, TextPane::CreateSingleLinePane(IDS_DOUBLE, false));
		if (m_lpsInputValue)
		{
			SetStringf(0, L"%f", m_lpsInputValue->Value.at); // STRING_OK
		}
		else
		{
			SetDecimal(0, 0);
		}

		break;
	case PT_BOOLEAN:
		InitPane(0, CheckPane::Create(IDS_BOOLEAN, m_lpsInputValue ? 0 != m_lpsInputValue->Value.b : false, false));
		break;
	case PT_DOUBLE:
		InitPane(0, TextPane::CreateSingleLinePane(IDS_DOUBLE, false));
		if (m_lpsInputValue)
		{
			SetStringf(0, L"%f", m_lpsInputValue->Value.dbl); // STRING_OK
		}
		else
		{
			SetDecimal(0, 0);
		}

		break;
	case PT_OBJECT:
		InitPane(0, TextPane::CreateSingleLinePaneID(IDS_OBJECT, IDS_OBJECTVALUE, true));
		break;
	case PT_R4:
		InitPane(0, TextPane::CreateSingleLinePane(IDS_FLOAT, false));
		if (m_lpsInputValue)
		{
			SetStringf(0, L"%f", m_lpsInputValue->Value.flt); // STRING_OK
		}
		else
		{
			SetDecimal(0, 0);
		}

		break;
	case PT_STRING8:
		InitPane(0, CountedTextPane::Create(IDS_ANSISTRING, false, IDS_CCH));
		InitPane(1, CountedTextPane::Create(IDS_BIN, false, IDS_CB));
		if (m_lpsInputValue && CheckStringProp(m_lpsInputValue, PT_STRING8))
		{
			SetStringA(0, m_lpsInputValue->Value.lpszA);

			lpPane = static_cast<CountedTextPane*>(GetPane(1));
			if (lpPane)
			{
				auto hRes = S_OK;
				EC_H(StringCbLengthA(m_lpsInputValue->Value.lpszA, STRSAFE_MAX_CCH * sizeof(char), &cbStr));

				lpPane->SetCount(cbStr);
				lpPane->SetBinary(
					reinterpret_cast<LPBYTE>(m_lpsInputValue->Value.lpszA),
					cbStr);
			}

			lpPane = static_cast<CountedTextPane*>(GetPane(0));
			if (lpPane) lpPane->SetCount(cbStr);
		}

		break;
	case PT_UNICODE:
		InitPane(0, CountedTextPane::Create(IDS_UNISTRING, false, IDS_CCH));
		InitPane(1, CountedTextPane::Create(IDS_BIN, false, IDS_CB));
		if (m_lpsInputValue && CheckStringProp(m_lpsInputValue, PT_UNICODE))
		{
			SetStringW(0, m_lpsInputValue->Value.lpszW);

			lpPane = static_cast<CountedTextPane*>(GetPane(1));
			if (lpPane)
			{
				auto hRes = S_OK;
				EC_H(StringCbLengthW(m_lpsInputValue->Value.lpszW, STRSAFE_MAX_CCH * sizeof(WCHAR), &cbStr));

				lpPane->SetCount(cbStr);
				lpPane->SetBinary(
					reinterpret_cast<LPBYTE>(m_lpsInputValue->Value.lpszW),
					cbStr);
			}

			lpPane = static_cast<CountedTextPane*>(GetPane(0));
			if (lpPane) lpPane->SetCount(cbStr % sizeof(WCHAR) ? 0 : cbStr / sizeof(WCHAR));
		}

		break;
	case PT_CURRENCY:
		InitPane(0, TextPane::CreateSingleLinePane(IDS_HI, false));
		InitPane(1, TextPane::CreateSingleLinePane(IDS_LO, false));
		InitPane(2, TextPane::CreateSingleLinePane(IDS_CURRENCY, false));
		if (m_lpsInputValue)
		{
			SetHex(0, m_lpsInputValue->Value.cur.Hi);
			SetHex(1, m_lpsInputValue->Value.cur.Lo);
			SetStringW(2, CurrencyToString(m_lpsInputValue->Value.cur));
		}
		else
		{
			SetHex(0, 0);
			SetHex(1, 0);
			SetStringW(2, L"0.0000"); // STRING_OK
		}

		break;
	case PT_ERROR:
		InitPane(0, TextPane::CreateSingleLinePane(IDS_ERRORCODEHEX, true));
		InitPane(1, TextPane::CreateSingleLinePane(IDS_ERRORNAME, true));
		if (m_lpsInputValue)
		{
			SetHex(0, m_lpsInputValue->Value.err);
			SetStringW(1, ErrorNameFromErrorCode(m_lpsInputValue->Value.err));
		}

		break;
	case PT_I2:
		InitPane(0, TextPane::CreateSingleLinePane(IDS_SIGNEDDECIMAL, false));
		InitPane(1, TextPane::CreateSingleLinePane(IDS_HEX, false));
		InitPane(2, m_lpSmartView);
		if (m_lpsInputValue)
		{
			SetDecimal(0, m_lpsInputValue->Value.i);
			SetHex(1, m_lpsInputValue->Value.i);
		}
		else
		{
			SetDecimal(0, 0);
			SetHex(1, 0);
		}

		if (m_lpSmartView)
		{
			m_lpSmartView->DisableDropDown();
			m_lpSmartView->SetStringW(szSmartView);
		}

		break;
	case PT_I8:
		InitPane(0, TextPane::CreateSingleLinePane(IDS_HIGHPART, false));
		InitPane(1, TextPane::CreateSingleLinePane(IDS_LOWPART, false));
		InitPane(2, TextPane::CreateSingleLinePane(IDS_DECIMAL, false));
		InitPane(3, m_lpSmartView);

		if (m_lpsInputValue)
		{
			SetHex(0, static_cast<int>(m_lpsInputValue->Value.li.HighPart));
			SetHex(1, static_cast<int>(m_lpsInputValue->Value.li.LowPart));
			SetStringf(2, L"%I64d", m_lpsInputValue->Value.li.QuadPart); // STRING_OK
		}
		else
		{
			SetHex(0, 0);
			SetHex(1, 0);
			SetDecimal(2, 0);
		}

		if (m_lpSmartView)
		{
			m_lpSmartView->DisableDropDown();
			m_lpSmartView->SetStringW(szSmartView);
		}

		break;
	case PT_BINARY:
		lpPane = static_cast<CountedTextPane*>(CountedTextPane::Create(IDS_BIN, false, IDS_CB));
		InitPane(0, lpPane);
		InitPane(1, CountedTextPane::Create(IDS_TEXT, false, IDS_CCH));
		InitPane(2, m_lpSmartView);

		if (m_lpsInputValue)
		{
			if (lpPane)
			{
				lpPane->SetCount(m_lpsInputValue->Value.bin.cb);
				lpPane->SetStringW(BinToHexString(&m_lpsInputValue->Value.bin, false));
				SetStringA(1, string(LPCSTR(m_lpsInputValue->Value.bin.lpb), m_lpsInputValue->Value.bin.cb));
			}

			lpPane = static_cast<CountedTextPane*>(GetPane(1));
			if (lpPane) lpPane->SetCount(m_lpsInputValue->Value.bin.cb);
		}

		if (m_lpSmartView)
		{
			m_lpSmartView->SetParser(iStructType);
			m_lpSmartView->SetStringW(szSmartView);
		}

		break;
	case PT_LONG:
		InitPane(0, TextPane::CreateSingleLinePane(IDS_UNSIGNEDDECIMAL, false));
		InitPane(1, TextPane::CreateSingleLinePane(IDS_HEX, false));
		InitPane(2, m_lpSmartView);
		if (m_lpsInputValue)
		{
			SetStringf(0, L"%d", m_lpsInputValue->Value.l); // STRING_OK
			SetHex(1, m_lpsInputValue->Value.l);
		}
		else
		{
			SetDecimal(0, 0);
			SetHex(1, 0);
			SetHex(2, 0);
		}

		if (m_lpSmartView)
		{
			m_lpSmartView->DisableDropDown();
			m_lpSmartView->SetStringW(szSmartView);
		}

		break;
	case PT_SYSTIME:
		InitPane(0, TextPane::CreateSingleLinePane(IDS_LOWDATETIME, false));
		InitPane(1, TextPane::CreateSingleLinePane(IDS_HIGHDATETIME, false));
		InitPane(2, TextPane::CreateSingleLinePane(IDS_DATE, true));
		if (m_lpsInputValue)
		{
			SetHex(0, static_cast<int>(m_lpsInputValue->Value.ft.dwLowDateTime));
			SetHex(1, static_cast<int>(m_lpsInputValue->Value.ft.dwHighDateTime));
			FileTimeToString(&m_lpsInputValue->Value.ft, szTemp1, szTemp2);
			SetStringW(2, szTemp1);
		}
		else
		{
			SetHex(0, 0);
			SetHex(1, 0);
		}

		break;
	case PT_CLSID:
		InitPane(0, TextPane::CreateSingleLinePane(IDS_GUID, false));
		if (m_lpsInputValue)
		{
			szGuid = GUIDToStringAndName(m_lpsInputValue->Value.lpguid);
		}
		else
		{
			szGuid = GUIDToStringAndName(nullptr);
		}

		SetStringW(0, szGuid);
		break;
	case PT_SRESTRICTION:
		InitPane(0, CollapsibleTextPane::Create(IDS_RESTRICTION, true));
		InterpretProp(m_lpsInputValue, &szTemp1, nullptr);
		SetStringW(0, szTemp1);
		break;
	case PT_ACTIONS:
		InitPane(0, CollapsibleTextPane::Create(IDS_ACTIONS, true));
		InterpretProp(m_lpsInputValue, &szTemp1, nullptr);
		SetStringW(0, szTemp1);
		break;
	default:
		InterpretProp(m_lpsInputValue, &szTemp1, &szTemp2);
		InitPane(0, CollapsibleTextPane::Create(IDS_VALUE, true));
		InitPane(1, CollapsibleTextPane::Create(IDS_ALTERNATEVIEW, true));
		SetStringW(IDS_VALUE, szTemp1);
		SetStringW(IDS_ALTERNATEVIEW, szTemp2);
		break;
	}
}

void CPropertyEditor::WriteStringsToSPropValue()
{
	auto hRes = S_OK;

	// Check first if we'll have anything to write
	switch (PROP_TYPE(m_ulPropTag))
	{
	case PT_OBJECT: // Nothing to write back - not supported
	case PT_SRESTRICTION:
	case PT_ACTIONS:
		return;
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
				reinterpret_cast<LPVOID*>(&m_lpsOutputValue)));
		}
		else
		{
			EC_H(MAPIAllocateBuffer(
				sizeof(SPropValue),
				reinterpret_cast<LPVOID*>(&m_lpsOutputValue)));
			m_lpAllocParent = m_lpsOutputValue;
		}
	}

	if (m_lpsOutputValue)
	{
		auto bFailed = false; // set true if we fail to get a prop and have to clean up memory
		m_lpsOutputValue->ulPropTag = m_ulPropTag;
		m_lpsOutputValue->dwAlignPad = NULL;
		vector<BYTE> bin;
		wstring szTmpString;

		switch (PROP_TYPE(m_ulPropTag))
		{
		case PT_I2: // treat as signed long
			szTmpString = GetStringUseControl(0);
			m_lpsOutputValue->Value.i = static_cast<short int>(wstringToLong(szTmpString, 10));
			break;
		case PT_LONG: // treat as unsigned long
			szTmpString = GetStringUseControl(0);
			m_lpsOutputValue->Value.l = static_cast<LONG>(wstringToUlong(szTmpString, 10));
			break;
		case PT_R4:
			szTmpString = GetStringUseControl(0);
			m_lpsOutputValue->Value.flt = static_cast<float>(wstringToDouble(szTmpString));
			break;
		case PT_DOUBLE:
			szTmpString = GetStringUseControl(0);
			m_lpsOutputValue->Value.dbl = wstringToDouble(szTmpString);
			break;
		case PT_CURRENCY:
			szTmpString = GetStringUseControl(0);
			m_lpsOutputValue->Value.cur.Hi = wstringToUlong(szTmpString, 16);
			szTmpString = GetStringUseControl(1);
			m_lpsOutputValue->Value.cur.Lo = wstringToUlong(szTmpString, 16);
			break;
		case PT_APPTIME:
			szTmpString = GetStringUseControl(0);
			m_lpsOutputValue->Value.at = wstringToDouble(szTmpString);
			break;
		case PT_ERROR: // unsigned
			szTmpString = GetStringUseControl(0);
			m_lpsOutputValue->Value.err = static_cast<SCODE>(wstringToUlong(szTmpString, 16));
			break;
		case PT_BOOLEAN:
			m_lpsOutputValue->Value.b = static_cast<unsigned short>(GetCheckUseControl(0));
			break;
		case PT_I8:
			szTmpString = GetStringUseControl(0);
			m_lpsOutputValue->Value.li.HighPart = static_cast<long>(wstringToUlong(szTmpString, 16));
			szTmpString = GetStringUseControl(1);
			m_lpsOutputValue->Value.li.LowPart = static_cast<long>(wstringToUlong(szTmpString, 16));
			break;
		case PT_STRING8:
			// We read strings out of the hex control in order to preserve any hex level tweaks the user
			// may have done. The RichEdit control likes throwing them away.
			szTmpString = GetStringUseControl(1);
			bin = HexStringToBin(szTmpString);
			m_lpsOutputValue->Value.lpszA = reinterpret_cast<LPSTR>(ByteVectorToMAPI(bin, m_lpAllocParent));
			break;
		case PT_UNICODE:
			// We read strings out of the hex control in order to preserve any hex level tweaks the user
			// may have done. The RichEdit control likes throwing them away.
			szTmpString = GetStringUseControl(1);
			bin = HexStringToBin(szTmpString);
			m_lpsOutputValue->Value.lpszW = reinterpret_cast<LPWSTR>(ByteVectorToMAPI(bin, m_lpAllocParent));
			break;
		case PT_SYSTIME:
			szTmpString = GetStringUseControl(0);
			m_lpsOutputValue->Value.ft.dwLowDateTime = wstringToUlong(szTmpString, 16);
			szTmpString = GetStringUseControl(1);
			m_lpsOutputValue->Value.ft.dwHighDateTime = wstringToUlong(szTmpString, 16);
			break;
		case PT_CLSID:
			EC_H(MAPIAllocateMore(
				sizeof(GUID),
				m_lpAllocParent,
				reinterpret_cast<LPVOID*>(&m_lpsOutputValue->Value.lpguid)));
			if (m_lpsOutputValue->Value.lpguid)
			{
				*m_lpsOutputValue->Value.lpguid = StringToGUID(GetStringUseControl(0));
			}

			break;
		case PT_BINARY:
			// remember we already read szTmpString and ulStrLen and found ulStrLen was even
			szTmpString = GetStringUseControl(0);
			bin = HexStringToBin(szTmpString);
			m_lpsOutputValue->Value.bin.lpb = ByteVectorToMAPI(bin, m_lpAllocParent);
			m_lpsOutputValue->Value.bin.cb = static_cast<ULONG>(bin.size());
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
				m_lpsOutputValue = nullptr;
				m_lpAllocParent = nullptr;
			}
			else
			{
				// If m_lpsOutputValue was allocated off a parent, we can't free it here
				// Just drop the reference and m_lpAllocParent's free will clean it up
				m_lpsOutputValue = nullptr;
			}
		}
	}
}

void CPropertyEditor::WriteSPropValueToObject() const
{
	if (!m_lpsOutputValue || !m_lpMAPIProp) return;

	auto hRes = S_OK;

	LPSPropProblemArray lpProblemArray = nullptr;

	EC_MAPI(m_lpMAPIProp->SetProps(
		1,
		m_lpsOutputValue,
		&lpProblemArray));

	EC_PROBLEMARRAY(lpProblemArray);
	MAPIFreeBuffer(lpProblemArray);

	EC_MAPI(m_lpMAPIProp->SaveChanges(KEEP_OPEN_READWRITE));
}

// Callers beware: Detatches and returns the modified prop value - this must be MAPIFreeBuffered!
_Check_return_ LPSPropValue CPropertyEditor::DetachModifiedSPropValue()
{
	auto m_lpRet = m_lpsOutputValue;
	m_lpsOutputValue = nullptr;
	return m_lpRet;
}

_Check_return_ ULONG CPropertyEditor::HandleChange(UINT nID)
{
	auto i = CEditor::HandleChange(nID);

	if (static_cast<ULONG>(-1) == i) return static_cast<ULONG>(-1);

	wstring szTmpString;
	wstring szTemp1;
	wstring szTemp2;
	wstring szSmartView;
	SPropValue sProp = { 0 };

	short int iVal = 0;
	LONG lVal = 0;
	CURRENCY curVal = { 0 };
	LARGE_INTEGER liVal = { 0 };
	FILETIME ftVal = { 0 };
	SBinary Bin = { 0 };
	vector<BYTE> bin;
	string lpszA;
	wstring lpszW;

	CountedTextPane* lpPane = nullptr;

	// If we get here, something changed - set the dirty flag
	m_bDirty = true;

	switch (PROP_TYPE(m_ulPropTag))
	{
	case PT_I2: // signed 16 bit
		szTmpString = GetStringUseControl(i);
		if (0 == i)
		{
			iVal = static_cast<short int>(wstringToLong(szTmpString, 10));
			SetHex(1, iVal);
		}
		else if (1 == i)
		{
			lVal = static_cast<short int>(wstringToLong(szTmpString, 16));
			SetDecimal(0, lVal);
		}

		sProp.ulPropTag = m_ulPropTag;
		sProp.Value.i = iVal;

		szSmartView = InterpretPropSmartView(&sProp,
			m_lpMAPIProp,
			nullptr,
			nullptr,
			m_bIsAB,
			m_bMVRow);

		if (m_lpSmartView) m_lpSmartView->SetStringW(szSmartView);

		break;
	case PT_LONG: // unsigned 32 bit
		szTmpString = GetStringUseControl(i);
		if (0 == i)
		{
			lVal = static_cast<LONG>(wstringToUlong(szTmpString, 10));
			SetHex(1, lVal);
		}
		else if (1 == i)
		{
			lVal = static_cast<LONG>(wstringToUlong(szTmpString, 16));
			SetStringf(0, L"%d", lVal); // STRING_OK
		}

		sProp.ulPropTag = m_ulPropTag;
		sProp.Value.l = lVal;

		szSmartView = InterpretPropSmartView(&sProp,
			m_lpMAPIProp,
			nullptr,
			nullptr,
			m_bIsAB,
			m_bMVRow);

		if (m_lpSmartView) m_lpSmartView->SetStringW(szSmartView);

		break;
	case PT_CURRENCY:
		if (0 == i || 1 == i)
		{
			szTmpString = GetStringUseControl(0);
			curVal.Hi = wstringToUlong(szTmpString, 16);
			szTmpString = GetStringUseControl(1);
			curVal.Lo = wstringToUlong(szTmpString, 16);
			SetStringW(2, CurrencyToString(curVal));
		}
		else if (2 == i)
		{
			szTmpString = GetStringUseControl(i);
			szTmpString = StripCharacter(szTmpString, L'.');
			curVal.int64 = wstringToInt64(szTmpString);
			SetHex(0, static_cast<int>(curVal.Hi));
			SetHex(1, static_cast<int>(curVal.Lo));
		}

		break;
	case PT_I8:
		if (0 == i || 1 == i)
		{
			szTmpString = GetStringUseControl(0);
			liVal.HighPart = static_cast<long>(wstringToUlong(szTmpString, 16));
			szTmpString = GetStringUseControl(1);
			liVal.LowPart = static_cast<long>(wstringToUlong(szTmpString, 16));
			SetStringf(2, L"%I64d", liVal.QuadPart); // STRING_OK
		}
		else if (2 == i)
		{
			szTmpString = GetStringUseControl(i);
			liVal.QuadPart = wstringToInt64(szTmpString);
			SetHex(0, static_cast<int>(liVal.HighPart));
			SetHex(1, static_cast<int>(liVal.LowPart));
		}

		sProp.ulPropTag = m_ulPropTag;
		sProp.Value.li = liVal;

		szSmartView = InterpretPropSmartView(&sProp,
			m_lpMAPIProp,
			nullptr,
			nullptr,
			m_bIsAB,
			m_bMVRow);

		if (m_lpSmartView) m_lpSmartView->SetStringW(szSmartView);

		break;
	case PT_SYSTIME: // components are unsigned hex
		szTmpString = GetStringUseControl(0);
		ftVal.dwLowDateTime = wstringToUlong(szTmpString, 16);
		szTmpString = GetStringUseControl(1);
		ftVal.dwHighDateTime = wstringToUlong(szTmpString, 16);

		FileTimeToString(&ftVal, szTemp1, szTemp2);
		SetStringW(2, szTemp1);
		break;
	case PT_BINARY:
		if (0 == i || 2 == i)
		{
			bin = GetBinaryUseControl(0);
			if (0 == i) SetStringA(1, string(LPCSTR(bin.data()), bin.size())); // ansi string
			Bin.lpb = bin.data();
			Bin.cb = ULONG(bin.size());
		}
		else if (1 == i)
		{
			lpszA = GetEditBoxTextA(1); // Do not free this
			Bin.lpb = LPBYTE(lpszA.c_str());
			Bin.cb = ULONG(lpszA.length() * sizeof(CHAR));

			SetBinary(0, Bin);
		}

		lpPane = static_cast<CountedTextPane*>(GetPane(0));
		if (lpPane) lpPane->SetCount(Bin.cb);

		lpPane = static_cast<CountedTextPane*>(GetPane(1));
		if (lpPane) lpPane->SetCount(Bin.cb);

		if (m_lpSmartView) m_lpSmartView->Parse(Bin);
		break;
	case PT_STRING8:
		if (0 == i)
		{
			size_t cbStr = 0;
			lpszA = GetEditBoxTextA(0);

			lpPane = static_cast<CountedTextPane*>(GetPane(1));
			if (lpPane)
			{
				cbStr = lpszA.length() * sizeof(CHAR);

				// Even if we don't have a string, still make the call to SetBinary
				// This will blank out the binary control when lpszA is NULL
				lpPane->SetBinary(LPBYTE(lpszA.c_str()), cbStr);
				lpPane->SetCount(cbStr);
			}

			lpPane = static_cast<CountedTextPane*>(GetPane(0));
			if (lpPane) lpPane->SetCount(cbStr);
		}
		else if (1 == i)
		{
			bin = GetBinaryUseControl(1);

			SetStringA(0, string(LPCSTR(bin.data()), bin.size()));

			lpPane = static_cast<CountedTextPane*>(GetPane(0));
			if (lpPane) lpPane->SetCount(bin.size());

			lpPane = static_cast<CountedTextPane*>(GetPane(1));
			if (lpPane) lpPane->SetCount(bin.size());
		}

		break;
	case PT_UNICODE:
		if (0 == i)
		{
			lpszW = GetEditBoxTextW(0);

			lpPane = static_cast<CountedTextPane*>(GetPane(1));
			if (lpPane)
			{
				auto cbStr = lpszW.length() * sizeof(WCHAR);

				// Even if we don't have a string, still make the call to SetBinary
				// This will blank out the binary control when lpszW is NULL
				lpPane->SetBinary(LPBYTE(lpszW.c_str()), cbStr);
				lpPane->SetCount(cbStr);
			}

			lpPane = static_cast<CountedTextPane*>(GetPane(0));
			if (lpPane) lpPane->SetCount(lpszW.length());
		}
		else if (1 == i)
		{
			lpPane = static_cast<CountedTextPane*>(GetPane(0));
			bin = GetBinaryUseControl(1);
			if (!(bin.size() % sizeof(WCHAR)))
			{
				SetStringW(0, wstring(LPCWSTR(bin.data()), bin.size() / sizeof(WCHAR)));
				if (lpPane) lpPane->SetCount(bin.size() / sizeof(WCHAR));
			}
			else
			{
				SetStringW(0, L"");
				if (lpPane) lpPane->SetCount(0);
			}

			lpPane = static_cast<CountedTextPane*>(GetPane(1));
			if (lpPane) lpPane->SetCount(bin.size());
		}

		break;
	default:
		break;
	}

	OnRecalcLayout();
	return i;
}

static wstring MVCLASS = L"CMultiValuePropertyEditor"; // STRING_OK

// Create an editor for a MAPI property
CMultiValuePropertyEditor::CMultiValuePropertyEditor(
	_In_ CWnd* pParentWnd,
	UINT uidTitle,
	UINT uidPrompt,
	bool bIsAB,
	_In_opt_ LPVOID lpAllocParent,
	_In_opt_ LPMAPIPROP lpMAPIProp,
	ULONG ulPropTag,
	_In_opt_ LPSPropValue lpsPropValue) :
	CEditor(pParentWnd, uidTitle, uidPrompt, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL)
{
	TRACE_CONSTRUCTOR(MVCLASS);

	m_bIsAB = bIsAB;
	m_lpAllocParent = lpAllocParent;
	m_lpsOutputValue = nullptr;

	m_lpMAPIProp = lpMAPIProp;
	if (m_lpMAPIProp) m_lpMAPIProp->AddRef();
	m_ulPropTag = ulPropTag;
	m_lpsInputValue = lpsPropValue;

	auto szPromptPostFix = format(L"\r\n%ws", TagToString(m_ulPropTag, m_lpMAPIProp, m_bIsAB, false).c_str()); // STRING_OK
	SetPromptPostFix(szPromptPostFix);

	InitPropertyControls();
}

CMultiValuePropertyEditor::~CMultiValuePropertyEditor()
{
	TRACE_DESTRUCTOR(MVCLASS);
	if (m_lpMAPIProp) m_lpMAPIProp->Release();
}

BOOL CMultiValuePropertyEditor::OnInitDialog()
{
	auto bRet = CEditor::OnInitDialog();

	ReadMultiValueStringsFromProperty();
	ResizeList(0, false);

	auto iStructType = FindSmartViewParserForProp(m_lpsInputValue->ulPropTag, NULL, nullptr, true);
	auto szSmartView = InterpretPropSmartView(
		m_lpsInputValue,
		m_lpMAPIProp,
		nullptr,
		nullptr,
		m_bIsAB,
		true);
	if (!szSmartView.empty())
	{
		auto lpPane = static_cast<SmartViewPane*>(GetPane(1));
		if (lpPane)
		{
			lpPane->SetParser(iStructType);
			lpPane->SetStringW(szSmartView);
		}
	}

	UpdateButtons();

	return bRet;
}

void CMultiValuePropertyEditor::OnOK()
{
	// This is where we write our changes back
	WriteMultiValueStringsToSPropValue();
	WriteSPropValueToObject();
	CMyDialog::OnOK(); // don't need to call CEditor::OnOK
}

void CMultiValuePropertyEditor::InitPropertyControls()
{
	InitPane(0, ListPane::Create(IDS_PROPVALUES, false, false, this));
	if (PT_MV_BINARY == PROP_TYPE(m_ulPropTag) ||
		PT_MV_LONG == PROP_TYPE(m_ulPropTag))
	{
		auto lpPane = static_cast<SmartViewPane*>(SmartViewPane::Create(IDS_SMARTVIEW));
		InitPane(1, lpPane);

		if (lpPane && PT_MV_LONG == PROP_TYPE(m_ulPropTag))
		{
			lpPane->DisableDropDown();
		}
	}
}

// Function must be called AFTER dialog controls have been created, not before
void CMultiValuePropertyEditor::ReadMultiValueStringsFromProperty() const
{
	InsertColumn(0, 0, IDS_ENTRY);
	InsertColumn(0, 1, IDS_VALUE);
	InsertColumn(0, 2, IDS_ALTERNATEVIEW);
	if (PT_MV_LONG == PROP_TYPE(m_ulPropTag) ||
		PT_MV_BINARY == PROP_TYPE(m_ulPropTag))
	{
		InsertColumn(0, 3, IDS_SMARTVIEW);
	}

	if (!m_lpsInputValue) return;
	if (!(PROP_TYPE(m_lpsInputValue->ulPropTag) & MV_FLAG)) return;

	// All the MV structures are basically the same, so we can cheat when we pull the count
	auto cValues = m_lpsInputValue->Value.MVi.cValues;
	for (ULONG iMVCount = 0; iMVCount < cValues; iMVCount++)
	{
		auto lpData = InsertListRow(0, iMVCount, format(L"%u", iMVCount)); // STRING_OK

		if (lpData)
		{
			lpData->InitializeMV(m_lpsInputValue, iMVCount);

			SPropValue sProp = { 0 };
			sProp.ulPropTag = CHANGE_PROP_TYPE(m_lpsInputValue->ulPropTag, PROP_TYPE(m_lpsInputValue->ulPropTag) & ~MV_FLAG);
			sProp.Value = lpData->MV()->m_val;
			UpdateListRow(&sProp, iMVCount);

			lpData->bItemFullyLoaded = true;
		}
	}
}

// Perisist the data in the controls to m_lpsOutputValue
void CMultiValuePropertyEditor::WriteMultiValueStringsToSPropValue()
{
	// If we're not dirty, don't write
	// Unless we had no input value. Then we're creating a new property.
	// So we're implicitly dirty.
	if (!IsDirty(0) && m_lpsInputValue) return;

	auto hRes = S_OK;
	// Take care of allocations first
	if (!m_lpsOutputValue)
	{
		if (m_lpAllocParent)
		{
			EC_H(MAPIAllocateMore(
				sizeof(SPropValue),
				m_lpAllocParent,
				reinterpret_cast<LPVOID*>(&m_lpsOutputValue)));
		}
		else
		{
			EC_H(MAPIAllocateBuffer(
				sizeof(SPropValue),
				reinterpret_cast<LPVOID*>(&m_lpsOutputValue)));
			m_lpAllocParent = m_lpsOutputValue;
		}
	}

	if (m_lpsOutputValue)
	{
		WriteMultiValueStringsToSPropValue(static_cast<LPVOID>(m_lpAllocParent), m_lpsOutputValue);
	}
}

// Given a pointer to an SPropValue structure which has already been allocated, fill out the values
void CMultiValuePropertyEditor::WriteMultiValueStringsToSPropValue(_In_ LPVOID lpParent, _In_ LPSPropValue lpsProp) const
{
	if (!lpParent || !lpsProp) return;

	auto hRes = S_OK;
	auto ulNumVals = GetListCount(0);

	lpsProp->ulPropTag = m_ulPropTag;
	lpsProp->dwAlignPad = NULL;

	switch (PROP_TYPE(lpsProp->ulPropTag))
	{
	case PT_MV_I2:
		EC_H(MAPIAllocateMore(sizeof(short int)* ulNumVals, lpParent, reinterpret_cast<LPVOID*>(&lpsProp->Value.MVi.lpi)));
		lpsProp->Value.MVi.cValues = ulNumVals;
		break;
	case PT_MV_LONG:
		EC_H(MAPIAllocateMore(sizeof(LONG)* ulNumVals, lpParent, reinterpret_cast<LPVOID*>(&lpsProp->Value.MVl.lpl)));
		lpsProp->Value.MVl.cValues = ulNumVals;
		break;
	case PT_MV_DOUBLE:
		EC_H(MAPIAllocateMore(sizeof(double)* ulNumVals, lpParent, reinterpret_cast<LPVOID*>(&lpsProp->Value.MVdbl.lpdbl)));
		lpsProp->Value.MVdbl.cValues = ulNumVals;
		break;
	case PT_MV_CURRENCY:
		EC_H(MAPIAllocateMore(sizeof(CURRENCY)* ulNumVals, lpParent, reinterpret_cast<LPVOID*>(&lpsProp->Value.MVcur.lpcur)));
		lpsProp->Value.MVcur.cValues = ulNumVals;
		break;
	case PT_MV_APPTIME:
		EC_H(MAPIAllocateMore(sizeof(double)* ulNumVals, lpParent, reinterpret_cast<LPVOID*>(&lpsProp->Value.MVat.lpat)));
		lpsProp->Value.MVat.cValues = ulNumVals;
		break;
	case PT_MV_SYSTIME:
		EC_H(MAPIAllocateMore(sizeof(FILETIME)* ulNumVals, lpParent, reinterpret_cast<LPVOID*>(&lpsProp->Value.MVft.lpft)));
		lpsProp->Value.MVft.cValues = ulNumVals;
		break;
	case PT_MV_I8:
		EC_H(MAPIAllocateMore(sizeof(LARGE_INTEGER)* ulNumVals, lpParent, reinterpret_cast<LPVOID*>(&lpsProp->Value.MVli.lpli)));
		lpsProp->Value.MVli.cValues = ulNumVals;
		break;
	case PT_MV_R4:
		EC_H(MAPIAllocateMore(sizeof(float)* ulNumVals, lpParent, reinterpret_cast<LPVOID*>(&lpsProp->Value.MVflt.lpflt)));
		lpsProp->Value.MVflt.cValues = ulNumVals;
		break;
	case PT_MV_STRING8:
		EC_H(MAPIAllocateMore(sizeof(LPSTR)* ulNumVals, lpParent, reinterpret_cast<LPVOID*>(&lpsProp->Value.MVszA.lppszA)));
		lpsProp->Value.MVszA.cValues = ulNumVals;
		break;
	case PT_MV_UNICODE:
		EC_H(MAPIAllocateMore(sizeof(LPWSTR)* ulNumVals, lpParent, reinterpret_cast<LPVOID*>(&lpsProp->Value.MVszW.lppszW)));
		lpsProp->Value.MVszW.cValues = ulNumVals;
		break;
	case PT_MV_BINARY:
		EC_H(MAPIAllocateMore(sizeof(SBinary)* ulNumVals, lpParent, reinterpret_cast<LPVOID*>(&lpsProp->Value.MVbin.lpbin)));
		lpsProp->Value.MVbin.cValues = ulNumVals;
		break;
	case PT_MV_CLSID:
		EC_H(MAPIAllocateMore(sizeof(GUID)* ulNumVals, lpParent, reinterpret_cast<LPVOID*>(&lpsProp->Value.MVguid.lpguid)));
		lpsProp->Value.MVguid.cValues = ulNumVals;
		break;
	default:
		break;
	}
	// Allocation is now done

	// Now write our data into the space we allocated
	for (ULONG iMVCount = 0; iMVCount < ulNumVals; iMVCount++)
	{
		auto lpData = GetListRowData(0, iMVCount);

		if (lpData && lpData->MV())
		{
			switch (PROP_TYPE(lpsProp->ulPropTag))
			{
			case PT_MV_I2:
				lpsProp->Value.MVi.lpi[iMVCount] = lpData->MV()->m_val.i;
				break;
			case PT_MV_LONG:
				lpsProp->Value.MVl.lpl[iMVCount] = lpData->MV()->m_val.l;
				break;
			case PT_MV_DOUBLE:
				lpsProp->Value.MVdbl.lpdbl[iMVCount] = lpData->MV()->m_val.dbl;
				break;
			case PT_MV_CURRENCY:
				lpsProp->Value.MVcur.lpcur[iMVCount] = lpData->MV()->m_val.cur;
				break;
			case PT_MV_APPTIME:
				lpsProp->Value.MVat.lpat[iMVCount] = lpData->MV()->m_val.at;
				break;
			case PT_MV_SYSTIME:
				lpsProp->Value.MVft.lpft[iMVCount] = lpData->MV()->m_val.ft;
				break;
			case PT_MV_I8:
				lpsProp->Value.MVli.lpli[iMVCount] = lpData->MV()->m_val.li;
				break;
			case PT_MV_R4:
				lpsProp->Value.MVflt.lpflt[iMVCount] = lpData->MV()->m_val.flt;
				break;
			case PT_MV_STRING8:
				EC_H(CopyStringA(&lpsProp->Value.MVszA.lppszA[iMVCount], lpData->MV()->m_val.lpszA, lpParent));
				break;
			case PT_MV_UNICODE:
				EC_H(CopyStringW(&lpsProp->Value.MVszW.lppszW[iMVCount], lpData->MV()->m_val.lpszW, lpParent));
				break;
			case PT_MV_BINARY:
				EC_H(CopySBinary(&lpsProp->Value.MVbin.lpbin[iMVCount], &lpData->MV()->m_val.bin, lpParent));
				break;
			case PT_MV_CLSID:
				if (lpData->MV()->m_val.lpguid)
				{
					lpsProp->Value.MVguid.lpguid[iMVCount] = *lpData->MV()->m_val.lpguid;
				}

				break;
			default:
				break;
			}
		}
	}
}

void CMultiValuePropertyEditor::WriteSPropValueToObject() const
{
	if (!m_lpsOutputValue || !m_lpMAPIProp) return;

	auto hRes = S_OK;

	LPSPropProblemArray lpProblemArray = nullptr;

	EC_MAPI(m_lpMAPIProp->SetProps(
		1,
		m_lpsOutputValue,
		&lpProblemArray));

	EC_PROBLEMARRAY(lpProblemArray);
	MAPIFreeBuffer(lpProblemArray);

	EC_MAPI(m_lpMAPIProp->SaveChanges(KEEP_OPEN_READWRITE));
}

// Callers beware: Detatches and returns the modified prop value - this must be MAPIFreeBuffered!
_Check_return_ LPSPropValue CMultiValuePropertyEditor::DetachModifiedSPropValue()
{
	auto m_lpRet = m_lpsOutputValue;
	m_lpsOutputValue = nullptr;
	return m_lpRet;
}

_Check_return_ bool CMultiValuePropertyEditor::DoListEdit(ULONG /*ulListNum*/, int iItem, _In_ SortListData* lpData)
{
	if (!lpData) return false;
	if (!lpData->MV())
	{
		lpData->InitializeMV(nullptr);
	}

	auto hRes = S_OK;
	SPropValue tmpPropVal = { 0 };
	// Strip off MV_FLAG since we're displaying only a row
	tmpPropVal.ulPropTag = m_ulPropTag & ~MV_FLAG;
	tmpPropVal.Value = lpData->MV()->m_val;

	LPSPropValue lpNewValue = nullptr;
	WC_H(DisplayPropertyEditor(
		this,
		IDS_EDITROW,
		NULL,
		m_bIsAB,
		NULL, // not passing an allocation parent because we know we're gonna free the result
		m_lpMAPIProp,
		NULL,
		true, // This is a row from a multivalued property. Only case we pass true here.
		&tmpPropVal,
		&lpNewValue));

	if (S_OK == hRes && lpNewValue)
	{
		lpData->InitializeMV(lpNewValue);

		// update the UI
		UpdateListRow(lpNewValue, iItem);
		UpdateSmartView();
		return true;
	}

	// Remember we didn't have an allocation parent - this is safe
	MAPIFreeBuffer(lpNewValue);
	return false;
}

void CMultiValuePropertyEditor::UpdateListRow(_In_ LPSPropValue lpProp, ULONG iMVCount) const
{
	wstring szTmp;
	wstring szAltTmp;

	InterpretProp(lpProp, &szTmp, &szAltTmp);
	SetListString(0, iMVCount, 1, szTmp);
	SetListString(0, iMVCount, 2, szAltTmp);

	if (PT_MV_LONG == PROP_TYPE(m_ulPropTag) ||
		PT_MV_BINARY == PROP_TYPE(m_ulPropTag))
	{
		auto szSmartView = InterpretPropSmartView(
			lpProp,
			m_lpMAPIProp,
			nullptr,
			nullptr,
			m_bIsAB,
			true);

		if (!szSmartView.empty()) SetListString(0, iMVCount, 3, szSmartView);
	}
}

void CMultiValuePropertyEditor::UpdateSmartView() const
{
	auto hRes = S_OK;
	auto lpPane = static_cast<SmartViewPane*>(GetPane(1));
	if (lpPane)
	{
		LPSPropValue lpsProp = nullptr;
		EC_H(MAPIAllocateBuffer(
			sizeof(SPropValue),
			reinterpret_cast<LPVOID*>(&lpsProp)));
		if (lpsProp)
		{
			WriteMultiValueStringsToSPropValue(static_cast<LPVOID>(lpsProp), lpsProp);

			wstring szSmartView;
			switch (PROP_TYPE(m_ulPropTag))
			{
			case PT_MV_LONG:
				szSmartView = InterpretPropSmartView(lpsProp, m_lpMAPIProp, nullptr, nullptr, m_bIsAB, true);
				break;
			case PT_MV_BINARY:
				auto iStructType = static_cast<__ParsingTypeEnum>(lpPane->GetDropDownSelectionValue());
				if (iStructType)
				{
					szSmartView = InterpretMVBinaryAsString(lpsProp->Value.MVbin, iStructType, m_lpMAPIProp);
				}
				break;
			}

			if (!szSmartView.empty())
			{
				lpPane->SetStringW(szSmartView);
			}
		}

		MAPIFreeBuffer(lpsProp);
	}
}

_Check_return_ ULONG CMultiValuePropertyEditor::HandleChange(UINT nID)
{
	auto i = static_cast<ULONG>(-1);

	// We check against the list pane first so we can track if it handled the change,
	// because if it did, we're going to recalculate smart view.
	auto lpPane = static_cast<ListPane*>(GetPane(0));
	if (lpPane)
	{
		i = lpPane->HandleChange(nID);
	}

	if (-1 == i)
	{
		i = CEditor::HandleChange(nID);
	}

	if (static_cast<ULONG>(-1) == i) return static_cast<ULONG>(-1);

	UpdateSmartView();
	OnRecalcLayout();

	return i;
}