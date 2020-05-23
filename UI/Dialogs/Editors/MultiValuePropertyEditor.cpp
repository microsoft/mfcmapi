#include <StdAfx.h>
#include <UI/Dialogs/Editors/MultiValuePropertyEditor.h>
#include <core/smartview/SmartView.h>
#include <core/sortlistdata/mvPropData.h>
#include <UI/Dialogs/Editors/PropertyEditor.h>
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
		UINT uidTitle,
		bool bIsAB,
		_In_opt_ LPVOID lpAllocParent,
		_In_opt_ LPMAPIPROP lpMAPIProp,
		ULONG ulPropTag,
		_In_opt_ const _SPropValue* lpsPropValue)
		: CEditor(pParentWnd, uidTitle, NULL, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL), m_bIsAB(bIsAB),
		  m_lpAllocParent(lpAllocParent), m_lpMAPIProp(lpMAPIProp), m_ulPropTag(ulPropTag),
		  m_lpsInputValue(lpsPropValue)
	{
		TRACE_CONSTRUCTOR(CLASS);

		if (m_lpMAPIProp) m_lpMAPIProp->AddRef();

		SetPromptPostFix(proptags::TagToString(m_ulPropTag, m_lpMAPIProp, m_bIsAB, false));

		InitPropertyControls();
	}

	CMultiValuePropertyEditor::~CMultiValuePropertyEditor()
	{
		TRACE_DESTRUCTOR(CLASS);
		if (m_lpMAPIProp) m_lpMAPIProp->Release();
	}

	BOOL CMultiValuePropertyEditor::OnInitDialog()
	{
		const auto bRet = CEditor::OnInitDialog();

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
		WriteSPropValueToObject();
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
					sProp.Value = mvprop->m_val;
				}

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

		// Take care of allocations first
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
			const auto ulNumVals = GetListCount(0);

			m_lpsOutputValue->ulPropTag = m_ulPropTag;
			m_lpsOutputValue->dwAlignPad = NULL;

			switch (PROP_TYPE(m_lpsOutputValue->ulPropTag))
			{
			case PT_MV_I2:
				m_lpsOutputValue->Value.MVi.lpi =
					mapi::allocate<short int*>(sizeof(short int) * ulNumVals, m_lpAllocParent);
				m_lpsOutputValue->Value.MVi.cValues = ulNumVals;
				break;
			case PT_MV_LONG:
				m_lpsOutputValue->Value.MVl.lpl = mapi::allocate<LONG*>(sizeof(LONG) * ulNumVals, m_lpAllocParent);
				m_lpsOutputValue->Value.MVl.cValues = ulNumVals;
				break;
			case PT_MV_DOUBLE:
				m_lpsOutputValue->Value.MVdbl.lpdbl =
					mapi::allocate<double*>(sizeof(double) * ulNumVals, m_lpAllocParent);
				m_lpsOutputValue->Value.MVdbl.cValues = ulNumVals;
				break;
			case PT_MV_CURRENCY:
				m_lpsOutputValue->Value.MVcur.lpcur =
					mapi::allocate<CURRENCY*>(sizeof(CURRENCY) * ulNumVals, m_lpAllocParent);
				m_lpsOutputValue->Value.MVcur.cValues = ulNumVals;
				break;
			case PT_MV_APPTIME:
				m_lpsOutputValue->Value.MVat.lpat =
					mapi::allocate<double*>(sizeof(double) * ulNumVals, m_lpAllocParent);
				m_lpsOutputValue->Value.MVat.cValues = ulNumVals;
				break;
			case PT_MV_SYSTIME:
				m_lpsOutputValue->Value.MVft.lpft =
					mapi::allocate<FILETIME*>(sizeof(FILETIME) * ulNumVals, m_lpAllocParent);
				m_lpsOutputValue->Value.MVft.cValues = ulNumVals;
				break;
			case PT_MV_I8:
				m_lpsOutputValue->Value.MVli.lpli =
					mapi::allocate<LARGE_INTEGER*>(sizeof(LARGE_INTEGER) * ulNumVals, m_lpAllocParent);
				m_lpsOutputValue->Value.MVli.cValues = ulNumVals;
				break;
			case PT_MV_R4:
				m_lpsOutputValue->Value.MVflt.lpflt =
					mapi::allocate<float*>(sizeof(float) * ulNumVals, m_lpAllocParent);
				m_lpsOutputValue->Value.MVflt.cValues = ulNumVals;
				break;
			case PT_MV_STRING8:
				m_lpsOutputValue->Value.MVszA.lppszA =
					mapi::allocate<LPSTR*>(sizeof(LPSTR) * ulNumVals, m_lpAllocParent);
				m_lpsOutputValue->Value.MVszA.cValues = ulNumVals;
				break;
			case PT_MV_UNICODE:
				m_lpsOutputValue->Value.MVszW.lppszW =
					mapi::allocate<LPWSTR*>(sizeof(LPWSTR) * ulNumVals, m_lpAllocParent);
				m_lpsOutputValue->Value.MVszW.cValues = ulNumVals;
				break;
			case PT_MV_BINARY:
				m_lpsOutputValue->Value.MVbin.lpbin =
					mapi::allocate<SBinary*>(sizeof(SBinary) * ulNumVals, m_lpAllocParent);
				m_lpsOutputValue->Value.MVbin.cValues = ulNumVals;
				break;
			case PT_MV_CLSID:
				m_lpsOutputValue->Value.MVguid.lpguid =
					mapi::allocate<GUID*>(sizeof(GUID) * ulNumVals, m_lpAllocParent);
				m_lpsOutputValue->Value.MVguid.cValues = ulNumVals;
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
						switch (PROP_TYPE(m_lpsOutputValue->ulPropTag))
						{
						case PT_MV_I2:
							m_lpsOutputValue->Value.MVi.lpi[iMVCount] = mvprop->m_val.i;
							break;
						case PT_MV_LONG:
							m_lpsOutputValue->Value.MVl.lpl[iMVCount] = mvprop->m_val.l;
							break;
						case PT_MV_DOUBLE:
							m_lpsOutputValue->Value.MVdbl.lpdbl[iMVCount] = mvprop->m_val.dbl;
							break;
						case PT_MV_CURRENCY:
							m_lpsOutputValue->Value.MVcur.lpcur[iMVCount] = mvprop->m_val.cur;
							break;
						case PT_MV_APPTIME:
							m_lpsOutputValue->Value.MVat.lpat[iMVCount] = mvprop->m_val.at;
							break;
						case PT_MV_SYSTIME:
							m_lpsOutputValue->Value.MVft.lpft[iMVCount] = mvprop->m_val.ft;
							break;
						case PT_MV_I8:
							m_lpsOutputValue->Value.MVli.lpli[iMVCount] = mvprop->m_val.li;
							break;
						case PT_MV_R4:
							m_lpsOutputValue->Value.MVflt.lpflt[iMVCount] = mvprop->m_val.flt;
							break;
						case PT_MV_STRING8:
							m_lpsOutputValue->Value.MVszA.lppszA[iMVCount] =
								mapi::CopyStringA(mvprop->m_val.lpszA, m_lpAllocParent);
							break;
						case PT_MV_UNICODE:
							m_lpsOutputValue->Value.MVszW.lppszW[iMVCount] =
								mapi::CopyStringW(mvprop->m_val.lpszW, m_lpAllocParent);
							break;
						case PT_MV_BINARY:
							m_lpsOutputValue->Value.MVbin.lpbin[iMVCount] =
								mapi::CopySBinary(mvprop->m_val.bin, m_lpAllocParent);
							break;
						case PT_MV_CLSID:
							if (mvprop->m_val.lpguid)
							{
								m_lpsOutputValue->Value.MVguid.lpguid[iMVCount] = *mvprop->m_val.lpguid;
							}

							break;
						default:
							break;
						}
					}
				}
			}
		}
	}

	void CMultiValuePropertyEditor::WriteSPropValueToObject() const
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

	// Callers beware: Detatches and returns the modified prop value - this must be MAPIFreeBuffered!
	_Check_return_ LPSPropValue CMultiValuePropertyEditor::DetachModifiedSPropValue() noexcept
	{
		const auto m_lpRet = m_lpsOutputValue;
		m_lpsOutputValue = nullptr;
		return m_lpRet;
	}

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
		tmpPropVal.Value = mvprop->m_val;

		LPSPropValue lpNewValue = nullptr;
		const auto hRes = WC_H(DisplayPropertyEditor(
			this,
			IDS_EDITROW,
			m_bIsAB,
			NULL, // not passing an allocation parent because we know we're gonna free the result
			m_lpMAPIProp,
			NULL,
			true, // This is a row from a multivalued property. Only case we pass true here.
			&tmpPropVal,
			&lpNewValue));

		if (hRes == S_OK && lpNewValue)
		{
			sortlistdata::mvPropData::init(lpData, lpNewValue);

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
			ret.push_back(mvprop ? mvprop->m_val.l : LONG{});
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
			const auto bin = mvprop ? mvprop->m_val.bin : SBinary{};
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