#include <StdAfx.h>
#include <UI/Dialogs/Editors/PropertyEditor.h>
#include <Interpret/Guids.h>
#include <MAPI/MAPIFunctions.h>
#include <Interpret/SmartView/SmartView.h>
#include <UI/Controls/SortList/MVPropData.h>
#include <UI/ViewPane/CountedTextPane.h>
#include <UI/Dialogs/Editors/MultiValuePropertyEditor.h>
#include <Interpret/InterpretProp.h>

namespace dialog
{
	namespace editor
	{
		_Check_return_ HRESULT DisplayPropertyEditor(_In_ CWnd* pParentWnd,
			UINT uidTitle,
			UINT uidPrompt,
			bool bIsAB,
			_In_opt_ LPVOID lpAllocParent,
			_In_opt_ LPMAPIPROP lpMAPIProp,
			ULONG ulPropTag,
			bool bMVRow,
			_In_opt_ const _SPropValue* lpsPropValue,
			_Inout_opt_ LPSPropValue* lpNewValue)
		{
			auto hRes = S_OK;

			_SPropValue* sourceProp = nullptr;
			// We got a MAPI prop object and no input value, go look one up
			if (lpMAPIProp && !lpsPropValue)
			{
				SPropTagArray sTag = { 0 };
				sTag.cValues = 1;
				sTag.aulPropTag[0] = PT_ERROR == PROP_TYPE(ulPropTag) ? CHANGE_PROP_TYPE(ulPropTag, PT_UNSPECIFIED) : ulPropTag;
				ULONG ulValues = NULL;

				WC_MAPI(lpMAPIProp->GetProps(&sTag, NULL, &ulValues, &sourceProp));

				// Suppress MAPI_E_NOT_FOUND error when the source type is non error
				if (sourceProp &&
					PT_ERROR == PROP_TYPE(sourceProp->ulPropTag) &&
					MAPI_E_NOT_FOUND == sourceProp->Value.err &&
					PT_ERROR != PROP_TYPE(ulPropTag)
					)
				{
					MAPIFreeBuffer(sourceProp);
					sourceProp = nullptr;
				}

				if (MAPI_E_CALL_FAILED == hRes)
				{
					// Just suppress this - let the user edit anyway
					hRes = S_OK;
				}

				// In all cases where we got a value back, we need to reset our property tag to the value we got
				// This will address when the source is PT_UNSPECIFIED, when the returned value is PT_ERROR,
				// or any other case where the returned value has a different type than requested
				if (SUCCEEDED(hRes) && sourceProp)
				{
					ulPropTag = sourceProp->ulPropTag;
				}

				lpsPropValue = sourceProp;
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

			MAPIFreeBuffer(sourceProp);

			return hRes;
		}

		static std::wstring SVCLASS = L"CPropertyEditor"; // STRING_OK

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
			_In_opt_ const _SPropValue* lpsPropValue) :
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

			const auto szPromptPostFix = strings::format(L"%ws%ws", uidPrompt ? L"\r\n" : L"", interpretprop::TagToString(m_ulPropTag | (m_bMVRow ? MV_FLAG : NULL), m_lpMAPIProp, m_bIsAB, false).c_str()); // STRING_OK
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
			return CEditor::OnInitDialog();
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
				m_lpSmartView = viewpane::SmartViewPane::Create(IDS_SMARTVIEW);
			}

			const auto smartView = smartview::InterpretPropSmartView2(
				m_lpsInputValue,
				m_lpMAPIProp,
				nullptr,
				nullptr,
				m_bIsAB,
				m_bMVRow); // Built from lpProp & lpMAPIProp

			const auto iStructType = smartView.first;
			const auto szSmartView = smartView.second;

			std::wstring szTemp1;
			std::wstring szTemp2;
			viewpane::CountedTextPane* lpPane = nullptr;
			size_t cbStr = 0;
			std::wstring szGuid;

			switch (PROP_TYPE(m_ulPropTag))
			{
			case PT_APPTIME:
				InitPane(0, viewpane::TextPane::CreateSingleLinePane(IDS_DOUBLE, false));
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
				InitPane(0, viewpane::CheckPane::Create(IDS_BOOLEAN, m_lpsInputValue ? 0 != m_lpsInputValue->Value.b : false, false));
				break;
			case PT_DOUBLE:
				InitPane(0, viewpane::TextPane::CreateSingleLinePane(IDS_DOUBLE, false));
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
				InitPane(0, viewpane::TextPane::CreateSingleLinePaneID(IDS_OBJECT, IDS_OBJECTVALUE, true));
				break;
			case PT_R4:
				InitPane(0, viewpane::TextPane::CreateSingleLinePane(IDS_FLOAT, false));
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
				InitPane(0, viewpane::CountedTextPane::Create(IDS_ANSISTRING, false, IDS_CCH));
				InitPane(1, viewpane::CountedTextPane::Create(IDS_BIN, false, IDS_CB));
				if (m_lpsInputValue && CheckStringProp(m_lpsInputValue, PT_STRING8))
				{
					auto lpszA = std::string(m_lpsInputValue->Value.lpszA);
					SetStringA(0, lpszA);

					lpPane = dynamic_cast<viewpane::CountedTextPane*>(GetPane(1));
					if (lpPane)
					{
						cbStr = lpszA.length() * sizeof(CHAR);

						lpPane->SetBinary(LPBYTE(lpszA.c_str()), cbStr);
						lpPane->SetCount(cbStr);
					}

					lpPane = dynamic_cast<viewpane::CountedTextPane*>(GetPane(0));
					if (lpPane) lpPane->SetCount(cbStr);
				}

				break;
			case PT_UNICODE:
				InitPane(0, viewpane::CountedTextPane::Create(IDS_UNISTRING, false, IDS_CCH));
				InitPane(1, viewpane::CountedTextPane::Create(IDS_BIN, false, IDS_CB));
				if (m_lpsInputValue && CheckStringProp(m_lpsInputValue, PT_UNICODE))
				{
					auto lpszW = std::wstring(m_lpsInputValue->Value.lpszW);
					SetStringW(0, lpszW);

					lpPane = dynamic_cast<viewpane::CountedTextPane*>(GetPane(1));
					if (lpPane)
					{
						cbStr = lpszW.length() * sizeof(WCHAR);

						lpPane->SetBinary(LPBYTE(lpszW.c_str()), cbStr);
						lpPane->SetCount(cbStr);
					}

					lpPane = dynamic_cast<viewpane::CountedTextPane*>(GetPane(0));
					if (lpPane) lpPane->SetCount(lpszW.length());
				}

				break;
			case PT_CURRENCY:
				InitPane(0, viewpane::TextPane::CreateSingleLinePane(IDS_HI, false));
				InitPane(1, viewpane::TextPane::CreateSingleLinePane(IDS_LO, false));
				InitPane(2, viewpane::TextPane::CreateSingleLinePane(IDS_CURRENCY, false));
				if (m_lpsInputValue)
				{
					SetHex(0, m_lpsInputValue->Value.cur.Hi);
					SetHex(1, m_lpsInputValue->Value.cur.Lo);
					SetStringW(2, strings::CurrencyToString(m_lpsInputValue->Value.cur));
				}
				else
				{
					SetHex(0, 0);
					SetHex(1, 0);
					SetStringW(2, L"0.0000"); // STRING_OK
				}

				break;
			case PT_ERROR:
				InitPane(0, viewpane::TextPane::CreateSingleLinePane(IDS_ERRORCODEHEX, true));
				InitPane(1, viewpane::TextPane::CreateSingleLinePane(IDS_ERRORNAME, true));
				if (m_lpsInputValue)
				{
					SetHex(0, m_lpsInputValue->Value.err);
					SetStringW(1, error::ErrorNameFromErrorCode(m_lpsInputValue->Value.err));
				}

				break;
			case PT_I2:
				InitPane(0, viewpane::TextPane::CreateSingleLinePane(IDS_SIGNEDDECIMAL, false));
				InitPane(1, viewpane::TextPane::CreateSingleLinePane(IDS_HEX, false));
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
				InitPane(0, viewpane::TextPane::CreateSingleLinePane(IDS_HIGHPART, false));
				InitPane(1, viewpane::TextPane::CreateSingleLinePane(IDS_LOWPART, false));
				InitPane(2, viewpane::TextPane::CreateSingleLinePane(IDS_DECIMAL, false));
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
				lpPane = viewpane::CountedTextPane::Create(IDS_BIN, false, IDS_CB);
				InitPane(0, lpPane);
				InitPane(1, viewpane::CountedTextPane::Create(IDS_TEXT, false, IDS_CCH));
				InitPane(2, m_lpSmartView);

				if (m_lpsInputValue)
				{
					if (lpPane)
					{
						lpPane->SetCount(m_lpsInputValue->Value.bin.cb);
						if (m_lpsInputValue->Value.bin.cb != 0)
						{
							lpPane->SetStringW(strings::BinToHexString(&m_lpsInputValue->Value.bin, false));
						}

						SetStringA(1, std::string(LPCSTR(m_lpsInputValue->Value.bin.lpb), m_lpsInputValue->Value.bin.cb));
					}

					lpPane = dynamic_cast<viewpane::CountedTextPane*>(GetPane(1));
					if (lpPane) lpPane->SetCount(m_lpsInputValue->Value.bin.cb);
				}

				if (m_lpSmartView)
				{
					m_lpSmartView->SetParser(iStructType);
					m_lpSmartView->SetStringW(szSmartView);
				}

				break;
			case PT_LONG:
				InitPane(0, viewpane::TextPane::CreateSingleLinePane(IDS_UNSIGNEDDECIMAL, false));
				InitPane(1, viewpane::TextPane::CreateSingleLinePane(IDS_HEX, false));
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
				InitPane(0, viewpane::TextPane::CreateSingleLinePane(IDS_LOWDATETIME, false));
				InitPane(1, viewpane::TextPane::CreateSingleLinePane(IDS_HIGHDATETIME, false));
				InitPane(2, viewpane::TextPane::CreateSingleLinePane(IDS_DATE, true));
				if (m_lpsInputValue)
				{
					SetHex(0, static_cast<int>(m_lpsInputValue->Value.ft.dwLowDateTime));
					SetHex(1, static_cast<int>(m_lpsInputValue->Value.ft.dwHighDateTime));
					strings::FileTimeToString(m_lpsInputValue->Value.ft, szTemp1, szTemp2);
					SetStringW(2, szTemp1);
				}
				else
				{
					SetHex(0, 0);
					SetHex(1, 0);
				}

				break;
			case PT_CLSID:
				InitPane(0, viewpane::TextPane::CreateSingleLinePane(IDS_GUID, false));
				if (m_lpsInputValue)
				{
					szGuid = guid::GUIDToStringAndName(m_lpsInputValue->Value.lpguid);
				}
				else
				{
					szGuid = guid::GUIDToStringAndName(nullptr);
				}

				SetStringW(0, szGuid);
				break;
			case PT_SRESTRICTION:
				InitPane(0, viewpane::TextPane::CreateCollapsibleTextPane(IDS_RESTRICTION, true));
				interpretprop::InterpretProp(m_lpsInputValue, &szTemp1, nullptr);
				SetStringW(0, szTemp1);
				break;
			case PT_ACTIONS:
				InitPane(0, viewpane::TextPane::CreateCollapsibleTextPane(IDS_ACTIONS, true));
				interpretprop::InterpretProp(m_lpsInputValue, &szTemp1, nullptr);
				SetStringW(0, szTemp1);
				break;
			default:
				interpretprop::InterpretProp(m_lpsInputValue, &szTemp1, &szTemp2);
				InitPane(0, viewpane::TextPane::CreateCollapsibleTextPane(IDS_VALUE, true));
				InitPane(1, viewpane::TextPane::CreateCollapsibleTextPane(IDS_ALTERNATEVIEW, true));
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
				std::vector<BYTE> bin;

				switch (PROP_TYPE(m_ulPropTag))
				{
				case PT_I2: // treat as signed long
					m_lpsOutputValue->Value.i = static_cast<short int>(strings::wstringToLong(GetStringW(0), 10));
					break;
				case PT_LONG: // treat as unsigned long
					m_lpsOutputValue->Value.l = static_cast<LONG>(strings::wstringToUlong(GetStringW(0), 10));
					break;
				case PT_R4:
					m_lpsOutputValue->Value.flt = static_cast<float>(strings::wstringToDouble(GetStringW(0)));
					break;
				case PT_DOUBLE:
					m_lpsOutputValue->Value.dbl = strings::wstringToDouble(GetStringW(0));
					break;
				case PT_CURRENCY:
					m_lpsOutputValue->Value.cur.Hi = strings::wstringToUlong(GetStringW(0), 16);
					m_lpsOutputValue->Value.cur.Lo = strings::wstringToUlong(GetStringW(1), 16);
					break;
				case PT_APPTIME:
					m_lpsOutputValue->Value.at = strings::wstringToDouble(GetStringW(0));
					break;
				case PT_ERROR: // unsigned
					m_lpsOutputValue->Value.err = static_cast<SCODE>(strings::wstringToUlong(GetStringW(0), 16));
					break;
				case PT_BOOLEAN:
					m_lpsOutputValue->Value.b = static_cast<unsigned short>(GetCheck(0));
					break;
				case PT_I8:
					m_lpsOutputValue->Value.li.HighPart = static_cast<long>(strings::wstringToUlong(GetStringW(0), 16));
					m_lpsOutputValue->Value.li.LowPart = static_cast<long>(strings::wstringToUlong(GetStringW(1), 16));
					break;
				case PT_STRING8:
					// We read strings out of the hex control in order to preserve any hex level tweaks the user
					// may have done. The RichEdit control likes throwing them away.
					bin = strings::HexStringToBin(GetStringW(1));
					m_lpsOutputValue->Value.lpszA = reinterpret_cast<LPSTR>(ByteVectorToMAPI(bin, m_lpAllocParent));
					break;
				case PT_UNICODE:
					// We read strings out of the hex control in order to preserve any hex level tweaks the user
					// may have done. The RichEdit control likes throwing them away.
					bin = strings::HexStringToBin(GetStringW(1));
					m_lpsOutputValue->Value.lpszW = reinterpret_cast<LPWSTR>(ByteVectorToMAPI(bin, m_lpAllocParent));
					break;
				case PT_SYSTIME:
					m_lpsOutputValue->Value.ft.dwLowDateTime = strings::wstringToUlong(GetStringW(0), 16);
					m_lpsOutputValue->Value.ft.dwHighDateTime = strings::wstringToUlong(GetStringW(1), 16);
					break;
				case PT_CLSID:
					EC_H(MAPIAllocateMore(
						sizeof(GUID),
						m_lpAllocParent,
						reinterpret_cast<LPVOID*>(&m_lpsOutputValue->Value.lpguid)));
					if (m_lpsOutputValue->Value.lpguid)
					{
						*m_lpsOutputValue->Value.lpguid = guid::StringToGUID(GetStringW(0));
					}

					break;
				case PT_BINARY:
					// remember we already read szTmpString and ulStrLen and found ulStrLen was even
					bin = strings::HexStringToBin(GetStringW(0));
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
			const auto m_lpRet = m_lpsOutputValue;
			m_lpsOutputValue = nullptr;
			return m_lpRet;
		}

		_Check_return_ ULONG CPropertyEditor::HandleChange(UINT nID)
		{
			const auto i = CEditor::HandleChange(nID);

			if (static_cast<ULONG>(-1) == i) return static_cast<ULONG>(-1);

			std::wstring szTmpString;
			std::wstring szTemp1;
			std::wstring szTemp2;
			std::wstring szSmartView;
			SPropValue sProp = { 0 };

			short int iVal = 0;
			LONG lVal = 0;
			CURRENCY curVal = { 0 };
			LARGE_INTEGER liVal = { 0 };
			FILETIME ftVal = { 0 };
			SBinary Bin = { 0 };
			std::vector<BYTE> bin;
			std::string lpszA;
			std::wstring lpszW;

			viewpane::CountedTextPane* lpPane = nullptr;

			// If we get here, something changed - set the dirty flag
			m_bDirty = true;

			switch (PROP_TYPE(m_ulPropTag))
			{
			case PT_I2: // signed 16 bit
				szTmpString = GetStringW(i);
				if (0 == i)
				{
					iVal = static_cast<short int>(strings::wstringToLong(szTmpString, 10));
					SetHex(1, iVal);
				}
				else if (1 == i)
				{
					lVal = static_cast<short int>(strings::wstringToLong(szTmpString, 16));
					SetDecimal(0, lVal);
				}

				sProp.ulPropTag = m_ulPropTag;
				sProp.Value.i = iVal;

				szSmartView = smartview::InterpretPropSmartView(&sProp,
					m_lpMAPIProp,
					nullptr,
					nullptr,
					m_bIsAB,
					m_bMVRow);

				if (m_lpSmartView) m_lpSmartView->SetStringW(szSmartView);

				break;
			case PT_LONG: // unsigned 32 bit
				szTmpString = GetStringW(i);
				if (0 == i)
				{
					lVal = static_cast<LONG>(strings::wstringToUlong(szTmpString, 10));
					SetHex(1, lVal);
				}
				else if (1 == i)
				{
					lVal = static_cast<LONG>(strings::wstringToUlong(szTmpString, 16));
					SetStringf(0, L"%d", lVal); // STRING_OK
				}

				sProp.ulPropTag = m_ulPropTag;
				sProp.Value.l = lVal;

				szSmartView = smartview::InterpretPropSmartView(&sProp,
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
					szTmpString = GetStringW(0);
					curVal.Hi = strings::wstringToUlong(szTmpString, 16);
					szTmpString = GetStringW(1);
					curVal.Lo = strings::wstringToUlong(szTmpString, 16);
					SetStringW(2, strings::CurrencyToString(curVal));
				}
				else if (2 == i)
				{
					szTmpString = GetStringW(i);
					szTmpString = strings::StripCharacter(szTmpString, L'.');
					curVal.int64 = strings::wstringToInt64(szTmpString);
					SetHex(0, static_cast<int>(curVal.Hi));
					SetHex(1, static_cast<int>(curVal.Lo));
				}

				break;
			case PT_I8:
				if (0 == i || 1 == i)
				{
					szTmpString = GetStringW(0);
					liVal.HighPart = static_cast<long>(strings::wstringToUlong(szTmpString, 16));
					szTmpString = GetStringW(1);
					liVal.LowPart = static_cast<long>(strings::wstringToUlong(szTmpString, 16));
					SetStringf(2, L"%I64d", liVal.QuadPart); // STRING_OK
				}
				else if (2 == i)
				{
					szTmpString = GetStringW(i);
					liVal.QuadPart = strings::wstringToInt64(szTmpString);
					SetHex(0, static_cast<int>(liVal.HighPart));
					SetHex(1, static_cast<int>(liVal.LowPart));
				}

				sProp.ulPropTag = m_ulPropTag;
				sProp.Value.li = liVal;

				szSmartView = smartview::InterpretPropSmartView(&sProp,
					m_lpMAPIProp,
					nullptr,
					nullptr,
					m_bIsAB,
					m_bMVRow);

				if (m_lpSmartView) m_lpSmartView->SetStringW(szSmartView);

				break;
			case PT_SYSTIME: // components are unsigned hex
				szTmpString = GetStringW(0);
				ftVal.dwLowDateTime = strings::wstringToUlong(szTmpString, 16);
				szTmpString = GetStringW(1);
				ftVal.dwHighDateTime = strings::wstringToUlong(szTmpString, 16);

				strings::FileTimeToString(ftVal, szTemp1, szTemp2);
				SetStringW(2, szTemp1);
				break;
			case PT_BINARY:
				if (0 == i || 2 == i)
				{
					bin = GetBinary(0);
					if (0 == i) SetStringA(1, std::string(LPCSTR(bin.data()), bin.size())); // ansi string
					Bin.lpb = bin.data();
					Bin.cb = ULONG(bin.size());
				}
				else if (1 == i)
				{
					lpszA = GetStringA(1); // Do not free this
					Bin.lpb = LPBYTE(lpszA.c_str());
					Bin.cb = ULONG(sizeof(CHAR) * lpszA.length());

					SetBinary(0, Bin.lpb, Bin.cb);
				}

				lpPane = dynamic_cast<viewpane::CountedTextPane*>(GetPane(0));
				if (lpPane) lpPane->SetCount(Bin.cb);

				lpPane = dynamic_cast<viewpane::CountedTextPane*>(GetPane(1));
				if (lpPane) lpPane->SetCount(Bin.cb);

				if (m_lpSmartView) m_lpSmartView->Parse(Bin);
				break;
			case PT_STRING8:
				if (0 == i)
				{
					size_t cbStr = 0;
					lpszA = GetStringA(0);

					lpPane = dynamic_cast<viewpane::CountedTextPane*>(GetPane(1));
					if (lpPane)
					{
						cbStr = lpszA.length() * sizeof(CHAR);

						// Even if we don't have a string, still make the call to SetBinary
						// This will blank out the binary control when lpszA is NULL
						lpPane->SetBinary(LPBYTE(lpszA.c_str()), cbStr);
						lpPane->SetCount(cbStr);
					}

					lpPane = dynamic_cast<viewpane::CountedTextPane*>(GetPane(0));
					if (lpPane) lpPane->SetCount(cbStr);
				}
				else if (1 == i)
				{
					bin = GetBinary(1);

					SetStringA(0, std::string(LPCSTR(bin.data()), bin.size()));

					lpPane = dynamic_cast<viewpane::CountedTextPane*>(GetPane(0));
					if (lpPane) lpPane->SetCount(bin.size());

					lpPane = dynamic_cast<viewpane::CountedTextPane*>(GetPane(1));
					if (lpPane) lpPane->SetCount(bin.size());
				}

				break;
			case PT_UNICODE:
				if (0 == i)
				{
					lpszW = GetStringW(0);

					lpPane = dynamic_cast<viewpane::CountedTextPane*>(GetPane(1));
					if (lpPane)
					{
						const auto cbStr = lpszW.length() * sizeof(WCHAR);

						// Even if we don't have a string, still make the call to SetBinary
						// This will blank out the binary control when lpszW is NULL
						lpPane->SetBinary(LPBYTE(lpszW.c_str()), cbStr);
						lpPane->SetCount(cbStr);
					}

					lpPane = dynamic_cast<viewpane::CountedTextPane*>(GetPane(0));
					if (lpPane) lpPane->SetCount(lpszW.length());
				}
				else if (1 == i)
				{
					lpPane = dynamic_cast<viewpane::CountedTextPane*>(GetPane(0));
					bin = GetBinary(1);
					if (!(bin.size() % sizeof(WCHAR)))
					{
						SetStringW(0, std::wstring(LPCWSTR(bin.data()), bin.size() / sizeof(WCHAR)));
						if (lpPane) lpPane->SetCount(bin.size() / sizeof(WCHAR));
					}
					else
					{
						SetStringW(0, L"");
						if (lpPane) lpPane->SetCount(0);
					}

					lpPane = dynamic_cast<viewpane::CountedTextPane*>(GetPane(1));
					if (lpPane) lpPane->SetCount(bin.size());
				}

				break;
			default:
				break;
			}

			OnRecalcLayout();
			return i;
		}
	}
}