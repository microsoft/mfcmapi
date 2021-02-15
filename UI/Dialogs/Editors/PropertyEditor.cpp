#include <StdAfx.h>
#include <UI/Dialogs/Editors/PropertyEditor.h>
#include <core/interpret/guid.h>
#include <core/mapi/mapiFunctions.h>
#include <core/smartview/SmartView.h>
#include <core/sortlistdata/mvPropData.h>
#include <UI/ViewPane/CountedTextPane.h>
#include <UI/Dialogs/Editors/MultiValuePropertyEditor.h>
#include <core/mapi/mapiMemory.h>
#include <core/utility/strings.h>
#include <core/utility/output.h>
#include <core/interpret/proptags.h>
#include <core/addin/mfcmapi.h>
#include <core/property/parseProperty.h>

namespace dialog::editor
{
	_Check_return_ std::shared_ptr<IPropEditor> DisplayPropertyEditor(
		_In_ CWnd* pParentWnd,
		UINT uidTitle,
		bool bIsAB,
		_In_opt_ LPVOID lpAllocParent,
		_In_opt_ LPMAPIPROP lpMAPIProp,
		ULONG ulPropTag,
		bool bMVRow,
		_In_opt_ const _SPropValue* lpsPropValue)
	{
		auto MyPropertyEditor = std::shared_ptr<IPropEditor>{};
		auto hRes = S_OK;

		_SPropValue* sourceProp = nullptr;
		// We got a MAPI prop object and no input value, go look one up
		if (lpMAPIProp && !lpsPropValue)
		{
			auto sTag = SPropTagArray{
				1, PROP_TYPE(ulPropTag) == PT_ERROR ? CHANGE_PROP_TYPE(ulPropTag, PT_UNSPECIFIED) : ulPropTag};
			ULONG ulValues = NULL;

			hRes = WC_MAPI(lpMAPIProp->GetProps(&sTag, NULL, &ulValues, &sourceProp));

			// Suppress MAPI_E_NOT_FOUND error when the source type is non error
			if (sourceProp && PROP_TYPE(sourceProp->ulPropTag) == PT_ERROR &&
				sourceProp->Value.err == MAPI_E_NOT_FOUND && PROP_TYPE(ulPropTag) != PT_ERROR)
			{
				MAPIFreeBuffer(sourceProp);
				sourceProp = nullptr;
			}

			if (hRes == MAPI_E_CALL_FAILED)
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
			MyPropertyEditor = std::make_shared<CMultiValuePropertyEditor>(
				pParentWnd, uidTitle, bIsAB, lpAllocParent, lpMAPIProp, ulPropTag, lpsPropValue);
		}
		// Or the single value prop case
		else
		{
			MyPropertyEditor = std::make_shared<CPropertyEditor>(
				pParentWnd, uidTitle, bIsAB, bMVRow, lpAllocParent, lpMAPIProp, ulPropTag, lpsPropValue);
		}

		if (MyPropertyEditor && !MyPropertyEditor->DisplayDialog())
		{
			MyPropertyEditor = {};
		}

		MAPIFreeBuffer(sourceProp);

		return MyPropertyEditor;
	}

	static std::wstring CLASS = L"CPropertyEditor"; // STRING_OK

	// Create an editor for a MAPI property
	CPropertyEditor::CPropertyEditor(
		_In_ CWnd* pParentWnd,
		UINT uidTitle,
		bool bIsAB,
		bool bMVRow,
		_In_opt_ LPVOID lpAllocParent,
		_In_opt_ LPMAPIPROP lpMAPIProp,
		ULONG ulPropTag,
		_In_opt_ const _SPropValue* lpsPropValue)
		: IPropEditor(pParentWnd, uidTitle, NULL, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL), m_bIsAB(bIsAB),
		  m_bMVRow(bMVRow), m_lpAllocParent(lpAllocParent), m_lpMAPIProp(lpMAPIProp), m_ulPropTag(ulPropTag),
		  m_lpsInputValue(lpsPropValue)
	{
		TRACE_CONSTRUCTOR(CLASS);

		if (m_lpMAPIProp) m_lpMAPIProp->AddRef();

		// If we didn't have an input value, we are creating a new property
		// So by definition, we're already dirty
		if (!m_lpsInputValue) m_bDirty = true;

		SetPromptPostFix(
			proptags::TagToString(m_ulPropTag | (m_bMVRow ? MV_FLAG : NULL), m_lpMAPIProp, m_bIsAB, false));

		InitPropertyControls();
	}

	CPropertyEditor::~CPropertyEditor()
	{
		TRACE_DESTRUCTOR(CLASS);
		if (m_lpMAPIProp) m_lpMAPIProp->Release();
	}

	BOOL CPropertyEditor::OnInitDialog() { return CEditor::OnInitDialog(); }

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
		std::wstring szTemp1;
		std::wstring szTemp2;

		switch (PROP_TYPE(m_ulPropTag))
		{
		case PT_APPTIME:
			AddPane(viewpane::TextPane::CreateSingleLinePane(0, IDS_DOUBLE, false));
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
			AddPane(viewpane::CheckPane::Create(
				0, IDS_BOOLEAN, m_lpsInputValue ? 0 != m_lpsInputValue->Value.b : false, false));
			break;
		case PT_DOUBLE:
			AddPane(viewpane::TextPane::CreateSingleLinePane(0, IDS_DOUBLE, false));
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
			AddPane(viewpane::TextPane::CreateSingleLinePaneID(0, IDS_OBJECT, IDS_OBJECTVALUE, true));
			break;
		case PT_R4:
			AddPane(viewpane::TextPane::CreateSingleLinePane(0, IDS_FLOAT, false));
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
			AddPane(viewpane::CountedTextPane::Create(0, IDS_ANSISTRING, false, IDS_CCH));
			AddPane(viewpane::CountedTextPane::Create(1, IDS_BIN, false, IDS_CB));
			if (m_lpsInputValue && strings::CheckStringProp(m_lpsInputValue, PT_STRING8))
			{
				const auto lpszA = std::string(m_lpsInputValue->Value.lpszA);
				const auto len = lpszA.length();
				SetStringA(0, lpszA);

				auto lpPane = std::dynamic_pointer_cast<viewpane::CountedTextPane>(GetPane(1));
				if (lpPane)
				{
					const auto cbStr = len * sizeof(CHAR);

					lpPane->SetBinary(reinterpret_cast<const BYTE*>(lpszA.c_str()), cbStr);
					lpPane->SetCount(cbStr);
				}

				lpPane = std::dynamic_pointer_cast<viewpane::CountedTextPane>(GetPane(0));
				if (lpPane) lpPane->SetCount(len);
			}

			break;
		case PT_UNICODE:
			AddPane(viewpane::CountedTextPane::Create(0, IDS_UNISTRING, false, IDS_CCH));
			AddPane(viewpane::CountedTextPane::Create(1, IDS_BIN, false, IDS_CB));
			if (m_lpsInputValue && strings::CheckStringProp(m_lpsInputValue, PT_UNICODE))
			{
				const auto lpszW = std::wstring(m_lpsInputValue->Value.lpszW);
				const auto len = lpszW.length();
				SetStringW(0, lpszW);

				auto lpPane = std::dynamic_pointer_cast<viewpane::CountedTextPane>(GetPane(1));
				if (lpPane)
				{
					const auto cbStr = len * sizeof(WCHAR);

					lpPane->SetBinary(reinterpret_cast<const BYTE*>(lpszW.c_str()), cbStr);
					lpPane->SetCount(cbStr);
				}

				lpPane = std::dynamic_pointer_cast<viewpane::CountedTextPane>(GetPane(0));
				if (lpPane) lpPane->SetCount(len);
			}

			break;
		case PT_CURRENCY:
			AddPane(viewpane::TextPane::CreateSingleLinePane(0, IDS_HI, false));
			AddPane(viewpane::TextPane::CreateSingleLinePane(1, IDS_LO, false));
			AddPane(viewpane::TextPane::CreateSingleLinePane(2, IDS_CURRENCY, false));
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
			AddPane(viewpane::TextPane::CreateSingleLinePane(0, IDS_ERRORCODEHEX, true));
			AddPane(viewpane::TextPane::CreateSingleLinePane(1, IDS_ERRORNAME, true));
			if (m_lpsInputValue)
			{
				SetHex(0, m_lpsInputValue->Value.err);
				SetStringW(1, error::ErrorNameFromErrorCode(m_lpsInputValue->Value.err));
			}

			break;
		case PT_I2:
			AddPane(viewpane::TextPane::CreateSingleLinePane(0, IDS_SIGNEDDECIMAL, false));
			AddPane(viewpane::TextPane::CreateSingleLinePane(1, IDS_HEX, false));
			AddPane(viewpane::TextPane::CreateMultiLinePane(
				2,
				IDS_SMARTVIEW,
				smartview::parsePropertySmartView(m_lpsInputValue, m_lpMAPIProp, nullptr, nullptr, m_bIsAB, m_bMVRow),
				true));
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

			break;
		case PT_I8:
			AddPane(viewpane::TextPane::CreateSingleLinePane(0, IDS_HIGHPART, false));
			AddPane(viewpane::TextPane::CreateSingleLinePane(1, IDS_LOWPART, false));
			AddPane(viewpane::TextPane::CreateSingleLinePane(2, IDS_DECIMAL, false));
			AddPane(viewpane::TextPane::CreateMultiLinePane(
				3,
				IDS_SMARTVIEW,
				smartview::parsePropertySmartView(m_lpsInputValue, m_lpMAPIProp, nullptr, nullptr, m_bIsAB, m_bMVRow),
				true));

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

			break;
		case PT_BINARY:
		{
			auto lpPane = viewpane::CountedTextPane::Create(0, IDS_BIN, false, IDS_CB);
			AddPane(lpPane);
			AddPane(viewpane::CountedTextPane::Create(1, IDS_TEXT, false, IDS_CCH));
			auto smartViewPane = viewpane::SmartViewPane::Create(2, IDS_SMARTVIEW);
			AddPane(smartViewPane);

			if (m_lpsInputValue)
			{
				const auto bin = mapi::getBin(m_lpsInputValue);
				if (lpPane)
				{
					lpPane->SetCount(bin.cb);
					if (bin.cb != 0)
					{
						lpPane->SetStringW(strings::BinToHexString(&bin, false));
					}

					SetStringA(1, std::string(reinterpret_cast<LPCSTR>(bin.lpb), bin.cb));
				}

				lpPane = std::dynamic_pointer_cast<viewpane::CountedTextPane>(GetPane(1));
				if (lpPane) lpPane->SetCount(bin.cb);
			}

			if (smartViewPane)
			{
				smartViewPane->SetParser(smartview::FindSmartViewParserForProp(
					m_lpsInputValue, m_lpMAPIProp, nullptr, nullptr, m_bIsAB, m_bMVRow));

				smartViewPane->Parse(std::vector<BYTE>(
					m_lpsInputValue ? mapi::getBin(m_lpsInputValue).lpb : nullptr,
					m_lpsInputValue ? mapi::getBin(m_lpsInputValue).lpb + mapi::getBin(m_lpsInputValue).cb : nullptr));

				smartViewPane->OnItemSelected = [&](auto _1) { return HighlightHex(0, _1); };
			}
		}

		break;
		case PT_LONG:
			AddPane(viewpane::TextPane::CreateSingleLinePane(0, IDS_UNSIGNEDDECIMAL, false));
			AddPane(viewpane::TextPane::CreateSingleLinePane(1, IDS_HEX, false));
			AddPane(viewpane::TextPane::CreateMultiLinePane(
				2,
				IDS_SMARTVIEW,
				smartview::parsePropertySmartView(m_lpsInputValue, m_lpMAPIProp, nullptr, nullptr, m_bIsAB, m_bMVRow),
				true));

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

			break;
		case PT_SYSTIME:
			AddPane(viewpane::TextPane::CreateSingleLinePane(0, IDS_LOWDATETIME, false));
			AddPane(viewpane::TextPane::CreateSingleLinePane(1, IDS_HIGHDATETIME, false));
			AddPane(viewpane::TextPane::CreateSingleLinePane(2, IDS_DATE, true));
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
			AddPane(viewpane::TextPane::CreateSingleLinePane(0, IDS_GUID, false));
			if (m_lpsInputValue)
			{
				SetStringW(0, guid::GUIDToStringAndName(m_lpsInputValue->Value.lpguid));
			}
			else
			{
				SetStringW(0, guid::GUIDToStringAndName(nullptr));
			}

			break;
		case PT_SRESTRICTION:
			AddPane(viewpane::TextPane::CreateCollapsibleTextPane(0, IDS_RESTRICTION, true));
			property::parseProperty(m_lpsInputValue, &szTemp1, nullptr);
			SetStringW(0, szTemp1);
			break;
		case PT_ACTIONS:
			AddPane(viewpane::TextPane::CreateCollapsibleTextPane(0, IDS_ACTIONS, true));
			property::parseProperty(m_lpsInputValue, &szTemp1, nullptr);
			SetStringW(0, szTemp1);
			break;
		default:
			property::parseProperty(m_lpsInputValue, &szTemp1, &szTemp2);
			AddPane(viewpane::TextPane::CreateCollapsibleTextPane(0, IDS_VALUE, true));
			AddPane(viewpane::TextPane::CreateCollapsibleTextPane(1, IDS_ALTERNATEVIEW, true));
			SetStringW(IDS_VALUE, szTemp1);
			SetStringW(IDS_ALTERNATEVIEW, szTemp2);
			break;
		}
	}

	void CPropertyEditor::WriteStringsToSPropValue()
	{
		// Check first if we'll have anything to write
		switch (PROP_TYPE(m_ulPropTag))
		{
		case PT_OBJECT: // Nothing to write back - not supported
		case PT_SRESTRICTION:
		case PT_ACTIONS:
			return;
		default:
			break;
		}

		// If nothing has changed, we're done.
		if (!m_bDirty) return;

		if (!m_lpsOutputValue)
		{
			m_lpsOutputValue = mapi::allocate<LPSPropValue>(sizeof(SPropValue), m_lpAllocParent);
			if (!m_lpAllocParent)
			{
				m_lpAllocParent = m_lpsOutputValue;
			}
		}

		if (m_lpsOutputValue)
		{
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
				m_lpsOutputValue->Value.lpszA = reinterpret_cast<LPSTR>(mapi::ByteVectorToMAPI(bin, m_lpAllocParent));
				break;
			case PT_UNICODE:
				// We read strings out of the hex control in order to preserve any hex level tweaks the user
				// may have done. The RichEdit control likes throwing them away.
				bin = strings::HexStringToBin(GetStringW(1));
				m_lpsOutputValue->Value.lpszW = reinterpret_cast<LPWSTR>(mapi::ByteVectorToMAPI(bin, m_lpAllocParent));
				break;
			case PT_SYSTIME:
				m_lpsOutputValue->Value.ft.dwLowDateTime = strings::wstringToUlong(GetStringW(0), 16);
				m_lpsOutputValue->Value.ft.dwHighDateTime = strings::wstringToUlong(GetStringW(1), 16);
				break;
			case PT_CLSID:
				m_lpsOutputValue->Value.lpguid = mapi::allocate<GUID*>(sizeof(GUID), m_lpAllocParent);
				if (m_lpsOutputValue->Value.lpguid)
				{
					*m_lpsOutputValue->Value.lpguid = guid::StringToGUID(GetStringW(0));
				}

				break;
			case PT_BINARY:
				// remember we already read szTmpString and ulStrLen and found ulStrLen was even
				bin = strings::HexStringToBin(GetStringW(0));
				mapi::setBin(m_lpsOutputValue) = {
					static_cast<ULONG>(bin.size()), mapi::ByteVectorToMAPI(bin, m_lpAllocParent)};
				break;
			default:
				// We shouldn't ever get here unless some new prop type shows up
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
				break;
			}
		}
	}

	void CPropertyEditor::WriteSPropValueToObject() const
	{
		if (!m_lpsOutputValue || !m_lpMAPIProp) return;

		LPSPropProblemArray lpProblemArray = nullptr;

		const auto hRes = EC_MAPI(m_lpMAPIProp->SetProps(1, m_lpsOutputValue, &lpProblemArray));

		EC_PROBLEMARRAY(lpProblemArray);
		MAPIFreeBuffer(lpProblemArray);

		if (SUCCEEDED(hRes))
		{
			EC_MAPI_S(m_lpMAPIProp->SaveChanges(KEEP_OPEN_READWRITE));
		}
	}

	// Returns the modified prop value - caller is responsible for freeing
	_Check_return_ LPSPropValue CPropertyEditor::getValue() noexcept { return m_lpsOutputValue; }

	_Check_return_ ULONG CPropertyEditor::HandleChange(UINT nID)
	{
		const auto paneID = CEditor::HandleChange(nID);

		if (paneID == static_cast<ULONG>(-1)) return static_cast<ULONG>(-1);

		auto sProp = SPropValue{};

		// If we get here, something changed - set the dirty flag
		m_bDirty = true;

		switch (PROP_TYPE(m_ulPropTag))
		{
		case PT_I2: // signed 16 bit
			if (paneID == 0)
			{
				sProp.Value.i = static_cast<short int>(strings::wstringToLong(GetStringW(0), 10));
				SetHex(1, sProp.Value.i);
			}
			else if (paneID == 1)
			{
				sProp.Value.i = static_cast<short int>(strings::wstringToLong(GetStringW(1), 16));
				SetDecimal(0, sProp.Value.i);
			}

			sProp.ulPropTag = m_ulPropTag;

			SetStringW(2, smartview::parsePropertySmartView(&sProp, m_lpMAPIProp, nullptr, nullptr, m_bIsAB, m_bMVRow));

			break;
		case PT_LONG: // unsigned 32 bit
			if (paneID == 0)
			{
				sProp.Value.l = static_cast<LONG>(strings::wstringToUlong(GetStringW(0), 10));
				SetHex(1, sProp.Value.l);
			}
			else if (paneID == 1)
			{
				sProp.Value.l = static_cast<LONG>(strings::wstringToUlong(GetStringW(1), 16));
				SetStringf(0, L"%d", sProp.Value.l); // STRING_OK
			}

			sProp.ulPropTag = m_ulPropTag;

			SetStringW(2, smartview::parsePropertySmartView(&sProp, m_lpMAPIProp, nullptr, nullptr, m_bIsAB, m_bMVRow));

			break;
		case PT_CURRENCY:
			if (paneID == 0 || paneID == 1)
			{
				sProp.Value.cur.Hi = strings::wstringToUlong(GetStringW(0), 16, false);
				sProp.Value.cur.Lo = strings::wstringToUlong(GetStringW(1), 16, false);
				SetStringW(2, strings::CurrencyToString(sProp.Value.cur));
			}
			else if (paneID == 2)
			{
				sProp.Value.cur.int64 = strings::wstringToCurrency(GetStringW(2));
				SetHex(0, static_cast<int>(sProp.Value.cur.Hi));
				SetHex(1, static_cast<int>(sProp.Value.cur.Lo));
			}

			break;
		case PT_I8:
			if (paneID == 0 || paneID == 1)
			{
				sProp.Value.li.HighPart = static_cast<long>(strings::wstringToUlong(GetStringW(0), 16, false));
				sProp.Value.li.LowPart = static_cast<long>(strings::wstringToUlong(GetStringW(1), 16, false));
				SetStringf(2, L"%I64d", sProp.Value.li.QuadPart); // STRING_OK
			}
			else if (paneID == 2)
			{
				sProp.Value.li.QuadPart = strings::wstringToInt64(GetStringW(2));
				SetHex(0, static_cast<int>(sProp.Value.li.HighPart));
				SetHex(1, static_cast<int>(sProp.Value.li.LowPart));
			}

			sProp.ulPropTag = m_ulPropTag;

			SetStringW(3, smartview::parsePropertySmartView(&sProp, m_lpMAPIProp, nullptr, nullptr, m_bIsAB, m_bMVRow));

			break;
		case PT_SYSTIME: // components are unsigned hex
		{

			sProp.Value.ft.dwLowDateTime = strings::wstringToUlong(GetStringW(0), 16);
			sProp.Value.ft.dwHighDateTime = strings::wstringToUlong(GetStringW(1), 16);

			std::wstring prop;
			std::wstring altprop;
			strings::FileTimeToString(sProp.Value.ft, prop, altprop);
			SetStringW(2, prop);
		}
		break;
		case PT_BINARY:
		{
			std::vector<BYTE> bin;
			if (paneID == 0 || paneID == 2)
			{
				ClearHighlight(0);
				bin = GetBinary(0);
				if (paneID == 0)
					SetStringA(1, std::string(reinterpret_cast<LPCSTR>(bin.data()), bin.size())); // ansi string
			}
			else if (paneID == 1)
			{
				const auto lpszA = GetStringA(1); // Do not free this
				bin =
					std::vector<BYTE>{lpszA.c_str(), lpszA.c_str() + static_cast<ULONG>(sizeof(CHAR) * lpszA.length())};
				SetBinary(0, bin.data(), static_cast<ULONG>(bin.size()));
			}

			auto lpPane = std::dynamic_pointer_cast<viewpane::CountedTextPane>(GetPane(0));
			if (lpPane) lpPane->SetCount(bin.size());

			lpPane = std::dynamic_pointer_cast<viewpane::CountedTextPane>(GetPane(1));
			if (lpPane) lpPane->SetCount(bin.size());

			auto smartViewPane = std::dynamic_pointer_cast<viewpane::SmartViewPane>(GetPane(2));
			if (smartViewPane)
			{
				smartViewPane->Parse(bin);
			}
		}

		break;
		case PT_STRING8:
			if (paneID == 0)
			{
				size_t cbStr = 0;
				const auto lpszA = GetStringA(0);

				auto lpPane = std::dynamic_pointer_cast<viewpane::CountedTextPane>(GetPane(1));
				if (lpPane)
				{
					cbStr = lpszA.length() * sizeof(CHAR);

					// Even if we don't have a string, still make the call to SetBinary
					// This will blank out the binary control when lpszA is NULL
					lpPane->SetBinary(reinterpret_cast<const BYTE*>(lpszA.c_str()), cbStr);
					lpPane->SetCount(cbStr);
				}

				lpPane = std::dynamic_pointer_cast<viewpane::CountedTextPane>(GetPane(0));
				if (lpPane) lpPane->SetCount(cbStr);
			}
			else if (paneID == 1)
			{
				const auto bin = GetBinary(1);

				SetStringA(0, std::string(reinterpret_cast<LPCSTR>(bin.data()), bin.size()));

				auto lpPane = std::dynamic_pointer_cast<viewpane::CountedTextPane>(GetPane(0));
				if (lpPane) lpPane->SetCount(bin.size());

				lpPane = std::dynamic_pointer_cast<viewpane::CountedTextPane>(GetPane(1));
				if (lpPane) lpPane->SetCount(bin.size());
			}

			break;
		case PT_UNICODE:
			if (paneID == 0)
			{
				const auto lpszW = GetStringW(0);

				auto lpPane = std::dynamic_pointer_cast<viewpane::CountedTextPane>(GetPane(1));
				if (lpPane)
				{
					const auto cbStr = lpszW.length() * sizeof(WCHAR);

					// Even if we don't have a string, still make the call to SetBinary
					// This will blank out the binary control when lpszW is NULL
					lpPane->SetBinary(reinterpret_cast<const BYTE*>(lpszW.c_str()), cbStr);
					lpPane->SetCount(cbStr);
				}

				lpPane = std::dynamic_pointer_cast<viewpane::CountedTextPane>(GetPane(0));
				if (lpPane) lpPane->SetCount(lpszW.length());
			}
			else if (paneID == 1)
			{
				auto lpPane = std::dynamic_pointer_cast<viewpane::CountedTextPane>(GetPane(0));
				const auto bin = GetBinary(1);
				if (!(bin.size() % sizeof(WCHAR)))
				{
					SetStringW(0, std::wstring(reinterpret_cast<LPCWSTR>(bin.data()), bin.size() / sizeof(WCHAR)));
					if (lpPane) lpPane->SetCount(bin.size() / sizeof(WCHAR));
				}
				else
				{
					SetStringW(0, L"");
					if (lpPane) lpPane->SetCount(0);
				}

				lpPane = std::dynamic_pointer_cast<viewpane::CountedTextPane>(GetPane(1));
				if (lpPane) lpPane->SetCount(bin.size());
			}

			break;
		default:
			break;
		}

		OnRecalcLayout();
		return paneID;
	}
} // namespace dialog::editor