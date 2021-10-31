#include <StdAfx.h>
#include <UI/Dialogs/Editors/propeditor/MultiValuePropertyEditor.h>
#include <UI/ViewPane/SmartViewPane.h>
#include <core/smartview/SmartView.h>
#include <core/sortlistdata/mvPropData.h>
#include <core/mapi/mapiMemory.h>
#include <core/utility/strings.h>
#include <core/utility/output.h>
#include <core/interpret/proptags.h>
#include <core/addin/mfcmapi.h>
#include <core/mapi/mapiFunctions.h>
#include <core/property/parseProperty.h>

namespace dialog::editor
{
	static std::wstring CLASS = L"CMultiValuePropertyEditor"; // STRING_OK

	// Create an editor for a MAPI property
	CMultiValuePropertyEditor::CMultiValuePropertyEditor(
		_In_ CWnd* pParentWnd,
		_In_ std::shared_ptr<cache::CMapiObjects> lpMapiObjects,
		UINT uidTitle,
		const std::wstring& name,
		bool bIsAB,
		_In_opt_ LPMAPIPROP lpMAPIProp,
		ULONG ulPropTag,
		_In_opt_ const _SPropValue* lpsPropValue)
		: IPropEditor(pParentWnd, uidTitle, NULL, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL), m_bIsAB(bIsAB),
		  m_lpMAPIProp(lpMAPIProp), m_ulPropTag(ulPropTag), m_lpsInputValue(lpsPropValue), m_name(name)
	{
		TRACE_CONSTRUCTOR(CLASS);
		m_lpMapiObjects = lpMapiObjects;

		if (m_lpMAPIProp) m_lpMAPIProp->AddRef();

		const auto propName = !name.empty() && PROP_ID(m_ulPropTag) == PROP_ID_NULL
								  ? proptags::TagToString(name, ulPropTag)
								  : proptags::TagToString(m_ulPropTag, m_lpMAPIProp, m_bIsAB, false);
		SetPromptPostFix(propName);

		InitPropertyControls();
	}

	CMultiValuePropertyEditor::~CMultiValuePropertyEditor()
	{
		TRACE_DESTRUCTOR(CLASS);
		if (m_lpMAPIProp) m_lpMAPIProp->Release();
	}

	BOOL CMultiValuePropertyEditor::OnInitDialog()
	{
		const auto bRet = IPropEditor::OnInitDialog();

		ReadMultiValueStringsFromProperty();
		ResizeList(0, false);

		if (PT_MV_BINARY == PROP_TYPE(m_ulPropTag))
		{
			auto smartViewPane = std::dynamic_pointer_cast<viewpane::SmartViewPane>(GetPane(1));
			if (smartViewPane)
			{
				smartViewPane->SetParser(smartview::FindSmartViewParserForProp(
					m_lpsInputValue, m_lpMAPIProp, nullptr, nullptr, m_bIsAB, true));
			}
		}

		UpdateSmartView();
		UpdateButtons();

		return bRet;
	}

	void CMultiValuePropertyEditor::OnOK()
	{
		// This is where we write our changes back
		WriteMultiValueStringsToSPropValue();
		CMyDialog::OnOK(); // don't need to call CEditor::OnOK
	}

	void CMultiValuePropertyEditor::InitPropertyControls()
	{
		AddPane(viewpane::ListPane::Create(0, IDS_PROPVALUES, false, false, ListEditCallBack(this)));
		SetListID(0);

		if (PT_MV_BINARY == PROP_TYPE(m_ulPropTag))
		{
			AddPane(viewpane::SmartViewPane::Create(1, IDS_SMARTVIEW));
		}
		else if (PT_MV_LONG == PROP_TYPE(m_ulPropTag))
		{
			AddPane(viewpane::TextPane::CreateMultiLinePane(1, IDS_SMARTVIEW, strings::emptystring, true));
		}
	}

	// Function must be called AFTER dialog controls have been created, not before
	void CMultiValuePropertyEditor::ReadMultiValueStringsFromProperty() const
	{
		InsertColumn(0, 0, IDS_ENTRY);
		InsertColumn(0, 1, IDS_VALUE);
		InsertColumn(0, 2, IDS_ALTERNATEVIEW);
		if (PT_MV_LONG == PROP_TYPE(m_ulPropTag) || PT_MV_BINARY == PROP_TYPE(m_ulPropTag))
		{
			InsertColumn(0, 3, IDS_SMARTVIEW);
		}

		if (!m_lpsInputValue) return;
		if (!(PROP_TYPE(m_lpsInputValue->ulPropTag) & MV_FLAG)) return;

		// All the MV structures are basically the same, so we can cheat when we pull the count
		const auto cValues = m_lpsInputValue->Value.MVi.cValues;
		for (ULONG iMVCount = 0; iMVCount < cValues; iMVCount++)
		{
			auto lpData = InsertListRow(0, iMVCount, std::to_wstring(iMVCount));

			if (lpData)
			{
				sortlistdata::mvPropData::init(lpData, m_lpsInputValue, iMVCount);

				auto sProp = SPropValue{};
				sProp.ulPropTag =
					CHANGE_PROP_TYPE(m_lpsInputValue->ulPropTag, PROP_TYPE(m_lpsInputValue->ulPropTag) & ~MV_FLAG);
				const auto mvprop = lpData->cast<sortlistdata::mvPropData>();
				if (mvprop)
				{
					sProp.Value = mvprop->getVal();
				}

				UpdateListRow(&sProp, iMVCount);

				lpData->setFullyLoaded(true);
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

		const auto ulNumVals = GetListCount(0);

		m_sOutputValue.ulPropTag = m_ulPropTag;

		switch (PROP_TYPE(m_sOutputValue.ulPropTag))
		{
		case PT_MV_I2:
			m_bin = std::vector<BYTE>(sizeof(short int) * ulNumVals);
			m_sOutputValue.Value.MVi.lpi = reinterpret_cast<short int*>(m_bin.data());
			m_sOutputValue.Value.MVi.cValues = ulNumVals;
			break;
		case PT_MV_LONG:
			m_bin = std::vector<BYTE>(sizeof(LONG) * ulNumVals);
			m_sOutputValue.Value.MVl.lpl = reinterpret_cast<LONG*>(m_bin.data());
			m_sOutputValue.Value.MVl.cValues = ulNumVals;
			break;
		case PT_MV_DOUBLE:
			m_bin = std::vector<BYTE>(sizeof(double) * ulNumVals);
			m_sOutputValue.Value.MVdbl.lpdbl = reinterpret_cast<double*>(m_bin.data());
			m_sOutputValue.Value.MVdbl.cValues = ulNumVals;
			break;
		case PT_MV_CURRENCY:
			m_bin = std::vector<BYTE>(sizeof(CURRENCY) * ulNumVals);
			m_sOutputValue.Value.MVcur.lpcur = reinterpret_cast<CURRENCY*>(m_bin.data());
			m_sOutputValue.Value.MVcur.cValues = ulNumVals;
			break;
		case PT_MV_APPTIME:
			m_bin = std::vector<BYTE>(sizeof(double) * ulNumVals);
			m_sOutputValue.Value.MVat.lpat = reinterpret_cast<double*>(m_bin.data());
			m_sOutputValue.Value.MVat.cValues = ulNumVals;
			break;
		case PT_MV_SYSTIME:
			m_bin = std::vector<BYTE>(sizeof(FILETIME) * ulNumVals);
			m_sOutputValue.Value.MVft.lpft = reinterpret_cast<FILETIME*>(m_bin.data());
			m_sOutputValue.Value.MVft.cValues = ulNumVals;
			break;
		case PT_MV_I8:
			m_bin = std::vector<BYTE>(sizeof(LARGE_INTEGER) * ulNumVals);
			m_sOutputValue.Value.MVli.lpli = reinterpret_cast<LARGE_INTEGER*>(m_bin.data());
			m_sOutputValue.Value.MVli.cValues = ulNumVals;
			break;
		case PT_MV_R4:
			m_bin = std::vector<BYTE>(sizeof(float) * ulNumVals);
			m_sOutputValue.Value.MVflt.lpflt = reinterpret_cast<float*>(m_bin.data());
			m_sOutputValue.Value.MVflt.cValues = ulNumVals;
			break;
		case PT_MV_STRING8:
			m_bin = std::vector<BYTE>(sizeof(LPSTR) * ulNumVals);
			m_mvA = std::vector<std::string>(ulNumVals);
			m_sOutputValue.Value.MVszA.lppszA = reinterpret_cast<LPSTR*>(m_bin.data());
			m_sOutputValue.Value.MVszA.cValues = ulNumVals;
			break;
		case PT_MV_UNICODE:
			m_bin = std::vector<BYTE>(sizeof(LPWSTR) * ulNumVals);
			m_mvW = std::vector<std::wstring>(ulNumVals);
			m_sOutputValue.Value.MVszW.lppszW = reinterpret_cast<LPWSTR*>(m_bin.data());
			m_sOutputValue.Value.MVszW.cValues = ulNumVals;
			break;
		case PT_MV_BINARY:
			m_bin = std::vector<BYTE>(sizeof(SBinary) * ulNumVals);
			m_mvBin = std::vector<std::vector<BYTE>>(ulNumVals);
			m_sOutputValue.Value.MVbin.lpbin = reinterpret_cast<SBinary*>(m_bin.data());
			m_sOutputValue.Value.MVbin.cValues = ulNumVals;
			break;
		case PT_MV_CLSID:
			m_bin = std::vector<BYTE>(sizeof(GUID) * ulNumVals);
			m_mvGuid = std::vector<GUID>(ulNumVals);
			m_sOutputValue.Value.MVguid.lpguid = reinterpret_cast<GUID*>(m_bin.data());
			m_sOutputValue.Value.MVguid.cValues = ulNumVals;
			break;
		default:
			break;
		}
		// Allocation is now done

		// Now write our data into the space we allocated
		for (ULONG iMVCount = 0; iMVCount < ulNumVals; iMVCount++)
		{
			const auto lpData = GetListRowData(0, iMVCount);

			if (lpData)
			{
				const auto mvprop = lpData->cast<sortlistdata::mvPropData>();
				if (mvprop)
				{
					switch (PROP_TYPE(m_sOutputValue.ulPropTag))
					{
					case PT_MV_I2:
						m_sOutputValue.Value.MVi.lpi[iMVCount] = mvprop->getVal().i;
						break;
					case PT_MV_LONG:
						m_sOutputValue.Value.MVl.lpl[iMVCount] = mvprop->getVal().l;
						break;
					case PT_MV_DOUBLE:
						m_sOutputValue.Value.MVdbl.lpdbl[iMVCount] = mvprop->getVal().dbl;
						break;
					case PT_MV_CURRENCY:
						m_sOutputValue.Value.MVcur.lpcur[iMVCount] = mvprop->getVal().cur;
						break;
					case PT_MV_APPTIME:
						m_sOutputValue.Value.MVat.lpat[iMVCount] = mvprop->getVal().at;
						break;
					case PT_MV_SYSTIME:
						m_sOutputValue.Value.MVft.lpft[iMVCount] = mvprop->getVal().ft;
						break;
					case PT_MV_I8:
						m_sOutputValue.Value.MVli.lpli[iMVCount] = mvprop->getVal().li;
						break;
					case PT_MV_R4:
						m_sOutputValue.Value.MVflt.lpflt[iMVCount] = mvprop->getVal().flt;
						break;
					case PT_MV_STRING8:
						m_mvA[iMVCount] = mvprop->getVal().lpszA;
						m_sOutputValue.Value.MVszA.lppszA[iMVCount] = m_mvA[iMVCount].data();
						break;
					case PT_MV_UNICODE:
						m_mvW[iMVCount] = mvprop->getVal().lpszW;
						m_sOutputValue.Value.MVszW.lppszW[iMVCount] = m_mvW[iMVCount].data();
						break;
					case PT_MV_BINARY:
						m_mvBin[iMVCount].assign(
							mvprop->getVal().bin.lpb, mvprop->getVal().bin.lpb + mvprop->getVal().bin.cb);
						m_sOutputValue.Value.MVbin.lpbin[iMVCount].cb = ULONG(m_mvBin[iMVCount].size());
						m_sOutputValue.Value.MVbin.lpbin[iMVCount].lpb = m_mvBin[iMVCount].data();
						break;
					case PT_MV_CLSID:
						if (mvprop->getVal().lpguid)
						{
							m_mvGuid[iMVCount] = *mvprop->getVal().lpguid;
							m_sOutputValue.Value.MVguid.lpguid[iMVCount] = m_mvGuid[iMVCount];
						}

						break;
					default:
						break;
					}
				}
			}
		}
	}

	// Returns the modified prop value - caller is responsible for freeing
	_Check_return_ LPSPropValue CMultiValuePropertyEditor::getValue() noexcept { return &m_sOutputValue; }

	_Check_return_ bool
	CMultiValuePropertyEditor::DoListEdit(ULONG /*ulListNum*/, int iItem, _In_ sortlistdata::sortListData* lpData)
	{
		if (!lpData) return false;
		if (!lpData->cast<sortlistdata::mvPropData>())
		{
			sortlistdata::mvPropData::init(lpData, nullptr);
		}

		const auto mvprop = lpData->cast<sortlistdata::mvPropData>();
		if (!mvprop) return false;

		SPropValue tmpPropVal = {};
		// Strip off MV_FLAG since we're displaying only a row
		tmpPropVal.ulPropTag = m_ulPropTag & ~MV_FLAG;
		tmpPropVal.Value = mvprop->getVal();

		const auto propEditor = DisplayPropertyEditor(
			this,
			m_lpMapiObjects,
			IDS_EDITROW,
			m_name,
			m_bIsAB,
			m_lpMAPIProp,
			NULL,
			true, // This is a row from a multivalued property. Only case we pass true here.
			&tmpPropVal);
		if (propEditor)
		{
			const auto lpNewValue = propEditor->getValue();
			if (lpNewValue)
			{
				sortlistdata::mvPropData::init(lpData, lpNewValue);

				// update the UI
				UpdateListRow(lpNewValue, iItem);
				UpdateSmartView();
			}

			return true;
		}

		return false;
	}

	void CMultiValuePropertyEditor::UpdateListRow(_In_ LPSPropValue lpProp, ULONG iMVCount) const
	{
		std::wstring szTmp;
		std::wstring szAltTmp;

		property::parseProperty(lpProp, &szTmp, &szAltTmp);
		SetListString(0, iMVCount, 1, szTmp);
		SetListString(0, iMVCount, 2, szAltTmp);

		if (PT_MV_LONG == PROP_TYPE(m_ulPropTag) || PT_MV_BINARY == PROP_TYPE(m_ulPropTag))
		{
			auto szSmartView = smartview::parsePropertySmartView(lpProp, m_lpMAPIProp, nullptr, nullptr, m_bIsAB, true);

			if (!szSmartView.empty()) SetListString(0, iMVCount, 3, szSmartView);
		}
	}

	std::vector<LONG> CMultiValuePropertyEditor::GetLongArray() const
	{
		if (PROP_TYPE(m_ulPropTag) != PT_MV_LONG) return {};
		const auto ulNumVals = GetListCount(0);
		auto ret = std::vector<LONG>{};

		for (auto i = ULONG{}; i < ulNumVals; i++)
		{
			const auto lpData = GetListRowData(0, i);
			const auto mvprop = lpData->cast<sortlistdata::mvPropData>();
			ret.push_back(mvprop ? mvprop->getVal().l : LONG{});
		}

		return ret;
	}

	std::vector<std::vector<BYTE>> CMultiValuePropertyEditor::GetBinaryArray() const
	{
		if (PROP_TYPE(m_ulPropTag) != PT_MV_BINARY) return {};
		const auto ulNumVals = GetListCount(0);
		auto ret = std::vector<std::vector<BYTE>>{};

		for (auto i = ULONG{}; i < ulNumVals; i++)
		{
			const auto lpData = GetListRowData(0, i);
			const auto mvprop = lpData->cast<sortlistdata::mvPropData>();
			const auto bin = mvprop ? mvprop->getVal().bin : SBinary{};
			ret.push_back(std::vector<BYTE>(bin.lpb, bin.lpb + bin.cb));
		}

		return ret;
	}

	void CMultiValuePropertyEditor::UpdateSmartView() const
	{
		switch (PROP_TYPE(m_ulPropTag))
		{
		case PT_MV_LONG:
		{
			auto smartViewPaneText = std::dynamic_pointer_cast<viewpane::TextPane>(GetPane(1));
			if (smartViewPaneText)
			{
				const auto rows = GetLongArray();
				auto npi = smartview::GetNamedPropInfo(m_ulPropTag, m_lpMAPIProp, nullptr, nullptr, m_bIsAB);

				smartViewPaneText->SetStringW(
					smartview::InterpretMVLongAsString(rows, m_ulPropTag, npi.first, &npi.second));
			}
		}

		break;
		case PT_MV_BINARY:
		{
			auto smartViewPane = std::dynamic_pointer_cast<viewpane::SmartViewPane>(GetPane(1));
			if (smartViewPane)
			{
				smartViewPane->Parse(GetBinaryArray());
			}
		}

		break;
		}
	}

	_Check_return_ ULONG CMultiValuePropertyEditor::HandleChange(UINT nID)
	{
		auto paneID = static_cast<ULONG>(-1);

		// We check against the list pane first so we can track if it handled the change,
		// because if it did, we're going to recalculate smart view.
		auto lpPane = std::dynamic_pointer_cast<viewpane::ListPane>(GetPane(0));
		if (lpPane)
		{
			paneID = lpPane->HandleChange(nID);
		}

		if (paneID == static_cast<ULONG>(-1))
		{
			paneID = CEditor::HandleChange(nID);
		}

		if (paneID == static_cast<ULONG>(-1)) return static_cast<ULONG>(-1);

		UpdateSmartView();
		OnRecalcLayout();

		return paneID;
	}
} // namespace dialog::editor