#include <StdAfx.h>
#include <UI/Dialogs/Editors/propeditor/PropertyEditor.h>
#include <UI/Dialogs/Editors/propeditor/MultiValuePropertyEditor.h>
#include <UI/ViewPane/SmartViewPane.h>
#include <core/interpret/guid.h>
#include <core/mapi/mapiFunctions.h>
#include <core/smartview/SmartView.h>
#include <core/sortlistdata/mvPropData.h>
#include <UI/ViewPane/CountedTextPane.h>
#include <core/mapi/mapiMemory.h>
#include <core/utility/strings.h>
#include <core/utility/output.h>
#include <core/interpret/proptags.h>
#include <core/addin/mfcmapi.h>
#include <core/property/parseProperty.h>
#include <UI/Dialogs/MFCUtilityFunctions.h>

namespace dialog::editor
{
	static std::wstring CLASS = L"CPropertyEditor"; // STRING_OK

	// Create an editor for a MAPI property
	CPropertyEditor::CPropertyEditor(
		_In_ CWnd* pParentWnd,
		_In_ std::shared_ptr<cache::CMapiObjects> lpMapiObjects,
		UINT uidTitle,
		const std::wstring& name,
		bool bIsAB,
		bool bMVRow,
		_In_opt_ LPMAPIPROP lpMAPIProp,
		ULONG ulPropTag,
		_In_opt_ const _SPropValue* lpsPropValue)
		: IPropEditor(pParentWnd, uidTitle, NULL, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL), m_bIsAB(bIsAB),
		  m_bMVRow(bMVRow),
		  m_lpMAPIProp(lpMAPIProp), m_ulPropTag(ulPropTag), m_lpsInputValue(lpsPropValue), m_name(name)
	{
		TRACE_CONSTRUCTOR(CLASS);
		m_lpMapiObjects = lpMapiObjects;

		if (m_lpMAPIProp) m_lpMAPIProp->AddRef();

		// If we didn't have an input value, we are creating a new property
		// So by definition, we're already dirty
		if (!m_lpsInputValue) m_bDirty = true;

		const auto propName =
			!name.empty() && PROP_ID(m_ulPropTag) == PROP_ID_NULL
				? proptags::TagToString(name, ulPropTag)
				: proptags::TagToString(m_ulPropTag | (m_bMVRow ? MV_FLAG : NULL), m_lpMAPIProp, m_bIsAB, false);
		SetPromptPostFix(propName);

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
				smartViewPane->OnActionButton = [&](auto _1) { return OpenEntry(_1); };
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

		m_sOutputValue.ulPropTag = m_ulPropTag;

		switch (PROP_TYPE(m_ulPropTag))
		{
		case PT_I2: // treat as signed long
			m_sOutputValue.Value.i = static_cast<short int>(strings::wstringToLong(GetStringW(0), 10));
			break;
		case PT_LONG: // treat as unsigned long
			m_sOutputValue.Value.l = static_cast<LONG>(strings::wstringToUlong(GetStringW(0), 10));
			break;
		case PT_R4:
			m_sOutputValue.Value.flt = static_cast<float>(strings::wstringToDouble(GetStringW(0)));
			break;
		case PT_DOUBLE:
			m_sOutputValue.Value.dbl = strings::wstringToDouble(GetStringW(0));
			break;
		case PT_CURRENCY:
			m_sOutputValue.Value.cur.Hi = strings::wstringToUlong(GetStringW(0), 16);
			m_sOutputValue.Value.cur.Lo = strings::wstringToUlong(GetStringW(1), 16);
			break;
		case PT_APPTIME:
			m_sOutputValue.Value.at = strings::wstringToDouble(GetStringW(0));
			break;
		case PT_ERROR: // unsigned
			m_sOutputValue.Value.err = static_cast<SCODE>(strings::wstringToUlong(GetStringW(0), 16));
			break;
		case PT_BOOLEAN:
			m_sOutputValue.Value.b = static_cast<unsigned short>(GetCheck(0));
			break;
		case PT_I8:
			m_sOutputValue.Value.li.HighPart = static_cast<long>(strings::wstringToUlong(GetStringW(0), 16));
			m_sOutputValue.Value.li.LowPart = static_cast<long>(strings::wstringToUlong(GetStringW(1), 16));
			break;
		case PT_STRING8:
			// We read strings out of the hex control in order to preserve any hex level tweaks the user
			// may have done. The RichEdit control likes throwing them away.
			m_bin = strings::HexStringToBin(GetStringW(1));
			m_bin.push_back(0); // Add null terminator
			m_sOutputValue.Value.lpszA = reinterpret_cast<LPSTR>(m_bin.data());
			break;
		case PT_UNICODE:
			// We read strings out of the hex control in order to preserve any hex level tweaks the user
			// may have done. The RichEdit control likes throwing them away.
			m_bin = strings::HexStringToBin(GetStringW(1));
			m_bin.push_back(0); // Add null terminator
			m_bin.push_back(0); // Add null terminator
			m_sOutputValue.Value.lpszW = reinterpret_cast<LPWSTR>(m_bin.data());
			break;
		case PT_SYSTIME:
			m_sOutputValue.Value.ft.dwLowDateTime = strings::wstringToUlong(GetStringW(0), 16);
			m_sOutputValue.Value.ft.dwHighDateTime = strings::wstringToUlong(GetStringW(1), 16);
			break;
		case PT_CLSID:
			m_guid = guid::StringToGUID(GetStringW(0));
			m_sOutputValue.Value.lpguid = &m_guid;
			break;
		case PT_BINARY:
			// remember we already read szTmpString and ulStrLen and found ulStrLen was even
			m_bin = strings::HexStringToBin(GetStringW(0));
			mapi::setBin(&m_sOutputValue) = {static_cast<ULONG>(m_bin.size()), m_bin.data()};
			break;
		default:
			m_sOutputValue = {};
			break;
		}
	}

	// Returns the modified prop value - caller is responsible for freeing
	_Check_return_ LPSPropValue CPropertyEditor::getValue() noexcept { return &m_sOutputValue; }

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

	void CPropertyEditor::OpenEntry(_In_ const SBinary& bin)
	{
		ULONG ulObjType = NULL;
		auto obj = mapi::CallOpenEntry<LPMAPIPROP>(
			m_lpMapiObjects->GetMDB(),
			m_lpMapiObjects->GetAddrBook(false),
			nullptr,
			m_lpMapiObjects->GetSession(),
			&bin,
			nullptr,
			MAPI_BEST_ACCESS,
			&ulObjType);

		if (obj)
		{
			WC_H_S(dialog::DisplayObject(obj, ulObjType, dialog::objectType::otDefault, nullptr, m_lpMapiObjects));
			obj->Release();
		}
	}
} // namespace dialog::editor