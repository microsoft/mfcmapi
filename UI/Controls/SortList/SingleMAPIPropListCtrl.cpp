#include <StdAfx.h>
#include <UI/Controls/SortList/SingleMAPIPropListCtrl.h>
#include <core/mapi/columnTags.h>
#include <UI/Dialogs/MFCUtilityFunctions.h>
#include <UI/UIFunctions.h>
#include <UI/MySecInfo.h>
#include <core/interpret/guid.h>
#include <UI/file/FileDialogEx.h>
#include <core/utility/import.h>
#include <core/mapi/mapiProgress.h>
#include <core/mapi/cache/namedProps.h>
#include <core/smartview/SmartView.h>
#include <core/PropertyBag/PropertyBag.h>
#include <core/PropertyBag/MAPIPropPropertyBag.h>
#include <core/PropertyBag/RowPropertyBag.h>
#include <core/utility/strings.h>
#include <core/mapi/cache/globalCache.h>
#include <UI/Dialogs/Editors/Editor.h>
#include <UI/Dialogs/Editors/restriction/RestrictEditor.h>
#include <UI/Dialogs/Editors/StreamEditor.h>
#include <UI/Dialogs/Editors/TagArrayEditor.h>
#include <UI/Dialogs/Editors/propeditor/ipropeditor.h>
#include <UI/Dialogs/Editors/PropertyTagEditor.h>
#include <core/mapi/cache/mapiObjects.h>
#include <core/mapi/mapiMemory.h>
#include <UI/addinui.h>
#include <core/mapi/mapiOutput.h>
#include <core/utility/registry.h>
#include <core/interpret/proptags.h>
#include <core/interpret/proptype.h>
#include <core/utility/output.h>
#include <core/mapi/mapiFunctions.h>
#include <core/property/parseProperty.h>

namespace controls::sortlistctrl
{
	static std::wstring CLASS = L"CSingleMAPIPropListCtrl";

	// 26 columns should be enough for anybody
#define MAX_SORT_COLS 26

	CSingleMAPIPropListCtrl::CSingleMAPIPropListCtrl(
		_In_ CWnd* pCreateParent,
		_In_ dialog::CBaseDialog* lpHostDlg,
		_In_ std::shared_ptr<cache::CMapiObjects> lpMapiObjects)
		: m_lpMapiObjects(lpMapiObjects), m_lpHostDlg(lpHostDlg)
	{
		TRACE_CONSTRUCTOR(CLASS);

		Create(pCreateParent, LVS_SINGLESEL, IDC_LIST_CTRL, true);

		if (m_lpHostDlg) m_lpHostDlg->AddRef();

		for (ULONG i = 0; i < columns::PropColumns.size(); i++)
		{
			const auto szHeaderName = strings::loadstring(columns::PropColumns[i].uidName);
			InsertColumnW(i, szHeaderName);
		}

		const auto lpMyHeader = GetHeaderCtrl();

		// Column orders are stored as lowercase letters
		// bacdefghi would mean the first two columns are swapped
		if (lpMyHeader && !registry::propertyColumnOrder.empty())
		{
			auto bSetCols = false;
			const auto nColumnCount = lpMyHeader->GetItemCount();
			const auto cchOrder = registry::propertyColumnOrder.length() - 1;
			if (nColumnCount == static_cast<int>(cchOrder))
			{
				auto order = std::vector<int>(nColumnCount);

				for (auto i = 0; i < nColumnCount; i++)
				{
					order[i] = registry::propertyColumnOrder[i] - L'a';
				}

				if (SetColumnOrderArray(static_cast<int>(order.size()), order.data()))
				{
					bSetCols = true;
				}
			}

			// If we didn't like the reg key, clear it so we don't see it again
			if (!bSetCols) registry::propertyColumnOrder.clear();
		}

		AutoSizeColumns(false);
	}

	CSingleMAPIPropListCtrl::~CSingleMAPIPropListCtrl()
	{
		TRACE_DESTRUCTOR(CLASS);
		if (m_sptExtraProps) MAPIFreeBuffer(m_sptExtraProps);
		if (m_lpHostDlg) m_lpHostDlg->Release();
	}

	BEGIN_MESSAGE_MAP(CSingleMAPIPropListCtrl, CSortListCtrl)
#pragma warning(push)
#pragma warning( \
	disable : 26454) // Warning C26454 Arithmetic overflow: 'operator' operation produces a negative unsigned result at compile time
	ON_NOTIFY_REFLECT(NM_DBLCLK, OnDblclk)
#pragma warning(pop)
	ON_WM_KEYDOWN()
	ON_WM_CONTEXTMENU()
	ON_MESSAGE(WM_MFCMAPI_SAVECOLUMNORDERLIST, msgOnSaveColumnOrder)
	END_MESSAGE_MAP()

	LRESULT CSingleMAPIPropListCtrl::WindowProc(UINT message, WPARAM wParam, LPARAM lParam)
	{
		switch (message)
		{
		case WM_ERASEBKGND:
			if (!m_lpPropBag)
			{
				return true;
			}

			break;
		case WM_PAINT:
			if (!m_lpPropBag)
			{
				ui::DrawHelpText(m_hWnd, IDS_HELPTEXTNOPROPS);
				return true;
			}

			break;
		}

		return CSortListCtrl::WindowProc(message, wParam, lParam);
	}

	// WM_MFCMAPI_SAVECOLUMNORDERLIST
	_Check_return_ LRESULT CSingleMAPIPropListCtrl::msgOnSaveColumnOrder(WPARAM /*wParam*/, LPARAM /*lParam*/)
	{
		const auto lpMyHeader = GetHeaderCtrl();

		if (lpMyHeader)
		{
			const auto columnCount = lpMyHeader->GetItemCount();
			if (columnCount && columnCount <= MAX_SORT_COLS)
			{
				auto columns = std::vector<int>(columnCount);

				registry::propertyColumnOrder.clear();
				EC_B_S(GetColumnOrderArray(columns.data(), columnCount));
				for (const auto column : columns)
				{
					registry::propertyColumnOrder.push_back(static_cast<wchar_t>(L'a' + column));
				}
			}
		}

		return S_OK;
	} // namespace sortlistctrl

	void CSingleMAPIPropListCtrl::InitMenu(_In_ CMenu* pMenu) const
	{
		if (pMenu)
		{
			const auto bHasSource = m_lpPropBag != nullptr;
			const auto bPropSelected = GetSelectedPropModelData() != nullptr;

			const auto ulStatus = cache::CGlobalCache::getInstance().GetBufferStatus();
			const auto lpEIDsToCopy = cache::CGlobalCache::getInstance().GetMessagesToCopy();
			pMenu->EnableMenuItem(
				ID_PASTE_PROPERTY, DIM(bHasSource && (ulStatus & BUFFER_PROPTAG) && (ulStatus & BUFFER_SOURCEPROPOBJ)));
			pMenu->EnableMenuItem(ID_COPYTO, DIM(bHasSource && (ulStatus & BUFFER_SOURCEPROPOBJ)));
			pMenu->EnableMenuItem(
				ID_PASTE_NAMEDPROPS,
				DIM(bHasSource && (ulStatus & BUFFER_MESSAGES) && lpEIDsToCopy && 1 == lpEIDsToCopy->cValues));

			pMenu->EnableMenuItem(ID_COPY_PROPERTY, DIM(bHasSource));

			pMenu->EnableMenuItem(
				ID_DISPLAYPROPERTYASSECURITYDESCRIPTORPROPSHEET,
				DIM(bHasSource && bPropSelected && import::pfnEditSecurity));
			pMenu->EnableMenuItem(ID_EDITPROPASBINARYSTREAM, DIM(bHasSource && bPropSelected));
			pMenu->EnableMenuItem(ID_EDITPROPERTY, DIM(bPropSelected));
			pMenu->EnableMenuItem(ID_EDITPROPERTYASASCIISTREAM, DIM(bHasSource && bPropSelected));
			pMenu->EnableMenuItem(ID_EDITPROPERTYASUNICODESTREAM, DIM(bHasSource && bPropSelected));
			pMenu->EnableMenuItem(ID_EDITPROPERTYASPRRTFCOMPRESSEDSTREAM, DIM(bHasSource && bPropSelected));
			pMenu->EnableMenuItem(ID_OPEN_PROPERTY, DIM(bPropSelected));

			pMenu->EnableMenuItem(ID_SAVEPROPERTIES, DIM(bHasSource));
			pMenu->EnableMenuItem(ID_EDITGIVENPROPERTY, DIM(bHasSource));
			pMenu->EnableMenuItem(ID_OPENPROPERTYASTABLE, DIM(bHasSource));
			pMenu->EnableMenuItem(ID_FINDALLNAMEDPROPS, DIM(bHasSource));
			pMenu->EnableMenuItem(ID_COUNTNAMEDPROPS, DIM(bHasSource));

			if (m_lpHostDlg)
			{
				for (ULONG ulMenu = ID_ADDINPROPERTYMENU;; ulMenu++)
				{
					const auto lpAddInMenu = ui::addinui::GetAddinMenuItem(m_lpHostDlg->m_hWnd, ulMenu);
					if (!lpAddInMenu) break;

					pMenu->EnableMenuItem(ulMenu, DIM(bPropSelected));
				}
			}

			const bool bCanDelete = bPropSelected && m_lpPropBag && m_lpPropBag->CanDelete();
			pMenu->EnableMenuItem(ID_DELETEPROPERTY, DIM(bCanDelete));
		}
	}

	_Check_return_ bool CSingleMAPIPropListCtrl::HandleMenu(WORD wMenuSelect)
	{
		output::DebugPrint(
			output::dbgLevel::Menu,
			L"CSingleMAPIPropListCtrl::HandleMenu wMenuSelect = 0x%X = %u\n",
			wMenuSelect,
			wMenuSelect);
		switch (wMenuSelect)
		{
		case ID_COPY_PROPERTY:
			OnCopyProperty();
			return true;
		case ID_COPYTO:
			OnCopyTo();
			return true;
		case ID_DELETEPROPERTY:
			OnDeleteProperty();
			return true;
		case ID_DISPLAYPROPERTYASSECURITYDESCRIPTORPROPSHEET:
			OnDisplayPropertyAsSecurityDescriptorPropSheet();
			return true;
		case ID_EDITGIVENPROPERTY:
			OnEditGivenProperty();
			return true;
		case ID_EDITPROPERTY:
			OnEditProp();
			return true;
		case ID_EDITPROPASBINARYSTREAM:
			OnEditPropAsStream(PT_BINARY, false);
			return true;
		case ID_EDITPROPERTYASASCIISTREAM:
			OnEditPropAsStream(PT_STRING8, false);
			return true;
		case ID_EDITPROPERTYASUNICODESTREAM:
			OnEditPropAsStream(PT_UNICODE, false);
			return true;
		case ID_EDITPROPERTYASPRRTFCOMPRESSEDSTREAM:
			OnEditPropAsStream(PT_BINARY, true);
			return true;
		case ID_FINDALLNAMEDPROPS:
			FindAllNamedProps();
			return true;
		case ID_COUNTNAMEDPROPS:
			CountNamedProps();
			return true;
		case ID_MODIFYEXTRAPROPS:
			OnModifyExtraProps();
			return true;
		case ID_OPEN_PROPERTY:
			OnOpenProperty();
			return true;
		case ID_OPENPROPERTYASTABLE:
			OnOpenPropertyAsTable();
			return true;
		case ID_PASTE_NAMEDPROPS:
			OnPasteNamedProps();
			return true;
		case ID_PASTE_PROPERTY:
			OnPasteProperty();
			return true;
		case ID_SAVEPROPERTIES:
			SavePropsToXML();
			return true;
		}

		return HandleAddInMenu(wMenuSelect);
	}

	_Check_return_ std::shared_ptr<sortlistdata::propModelData>
	CSingleMAPIPropListCtrl::GetSelectedPropModelData() const
	{
		const auto iItem = GetNextItem(-1, LVNI_FOCUSED | LVNI_SELECTED);

		if (-1 != iItem)
		{
			const auto lpListData = reinterpret_cast<sortlistdata::sortListData*>(GetItemData(iItem));
			if (lpListData)
			{
				return lpListData->cast<sortlistdata::propModelData>();
			}
		}

		return {};
	}

	// Call GetProps with NULL to get a list of (almost) all properties.
	// Parse this list and render them in the control.
	// Add any extra props we've asked for through the UI
	void CSingleMAPIPropListCtrl::LoadMAPIPropList()
	{
		if (!m_lpPropBag) return;

		CWaitCursor Wait; // Change the mouse to an hourglass while we work.
		auto models = std::vector<std::shared_ptr<model::mapiRowModel>>{};
		if (!registry::onlyAdditionalProperties)
		{
			models = m_lpPropBag->GetAllModels();
		}

		// Now check if the user has given us any other properties to add and get them one at a time
		if (m_sptExtraProps)
		{
			for (ULONG i = 0; i < m_sptExtraProps->cValues; i++)
			{
				models.emplace_back(m_lpPropBag->GetOneModel(mapi::getTag(m_sptExtraProps, i)));
			}
		}

		// Add our props to the view
		if (!models.empty())
		{
			// Set the item count to speed up the addition of items
			SetItemCount(static_cast<int>(models.size()));

			// get each model in turn and add it to the list
			ULONG ulCurListBoxRow = 0;
			for (const auto model : models)
			{
				AddPropToListBox(ulCurListBoxRow++, model);
			}
		}

		if (m_lpHostDlg)
		{
			// This flag may be set by a GetProps call, so we make this check AFTER we get our props
			if (m_lpPropBag->IsBackedByGetProps())
			{
				m_lpHostDlg->UpdateStatusBarText(statusPane::infoText, IDS_PROPSFROMGETPROPS);
			}
			else
			{
				m_lpHostDlg->UpdateStatusBarText(statusPane::infoText, IDS_PROPSFROMROW);
			}
		}

		output::DebugPrintEx(
			output::dbgLevel::Generic, CLASS, L"LoadMAPIPropList", L"added %u properties\n", models.size());

		SortClickedColumn();
	}

	void CSingleMAPIPropListCtrl::RefreshMAPIPropList()
	{
		output::DebugPrintEx(output::dbgLevel::Generic, CLASS, L"RefreshMAPIPropList", L"\n");

		// Turn off redraw while we work on the window
		MySetRedraw(false);
		auto MyPos = GetFirstSelectedItemPosition();

		const auto iSelectedItem = GetNextSelectedItem(MyPos);

		EC_B_S(DeleteAllItems());

		if (m_lpPropBag)
		{
			LoadMAPIPropList();
		}

		SetItemState(iSelectedItem, LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED);

		EnsureVisible(iSelectedItem, false);

		// Turn redraw back on to update our view
		MySetRedraw(true);

		if (m_lpHostDlg) m_lpHostDlg->UpdateStatusBarText(statusPane::data2, IDS_STATUSTEXTNUMPROPS, GetItemCount());
	}

	void CSingleMAPIPropListCtrl::AddPropToExtraProps(ULONG ulPropTag, bool bRefresh)
	{
		output::DebugPrintEx(
			output::dbgLevel::Generic, CLASS, L"AddPropToExtraProps", L"adding proptag 0x%X\n", ulPropTag);

		// Cache this proptag so we continue to request it in this view
		// We've got code to refresh any props cached in m_sptExtraProps...let's add to that.

		SPropTagArray sptSingleProp = {1, ulPropTag};

		AddPropsToExtraProps(&sptSingleProp, bRefresh);
	}

	void CSingleMAPIPropListCtrl::AddPropsToExtraProps(_In_ LPSPropTagArray lpPropsToAdd, bool bRefresh)
	{
		output::DebugPrintEx(
			output::dbgLevel::Generic, CLASS, L"AddPropsToExtraProps", L"adding prop array %p\n", lpPropsToAdd);

		const auto lpNewExtraProps = mapi::ConcatSPropTagArrays(m_sptExtraProps, lpPropsToAdd);

		MAPIFreeBuffer(m_sptExtraProps);
		m_sptExtraProps = lpNewExtraProps;

		if (bRefresh)
		{
			RefreshMAPIPropList();
		}
	}

	static TypeIcon _PropTypeIcons[] = {
		{PT_UNSPECIFIED, sortIcon::ptUnspecified},
		{PT_NULL, sortIcon::ptNull},
		{PT_I2, sortIcon::ptI2},
		{PT_LONG, sortIcon::ptLong},
		{PT_R4, sortIcon::ptR4},
		{PT_DOUBLE, sortIcon::ptDouble},
		{PT_CURRENCY, sortIcon::ptCurrency},
		{PT_APPTIME, sortIcon::ptAppTime},
		{PT_ERROR, sortIcon::ptError},
		{PT_BOOLEAN, sortIcon::ptBoolean},
		{PT_OBJECT, sortIcon::ptObject},
		{PT_I8, sortIcon::ptI8},
		{PT_STRING8, sortIcon::ptString8},
		{PT_UNICODE, sortIcon::ptUnicode},
		{PT_SYSTIME, sortIcon::ptSysTime},
		{PT_CLSID, sortIcon::ptClsid},
		{PT_BINARY, sortIcon::ptBinary},
		{PT_MV_I2, sortIcon::mvI2},
		{PT_MV_LONG, sortIcon::mvLong},
		{PT_MV_R4, sortIcon::mvR4},
		{PT_MV_DOUBLE, sortIcon::mvDouble},
		{PT_MV_CURRENCY, sortIcon::mvCurrency},
		{PT_MV_APPTIME, sortIcon::mvAppTime},
		{PT_MV_SYSTIME, sortIcon::mvSysTime},
		{PT_MV_STRING8, sortIcon::mvString8},
		{PT_MV_BINARY, sortIcon::mvBinary},
		{PT_MV_UNICODE, sortIcon::mvUnicode},
		{PT_MV_CLSID, sortIcon::mvClsid},
		{PT_MV_I8, sortIcon::mvI8},
		{PT_SRESTRICTION, sortIcon::sRestriction},
		{PT_ACTIONS, sortIcon::actions},
	};

	// Render the row model in the list.
	void CSingleMAPIPropListCtrl::AddPropToListBox(int iRow, const std::shared_ptr<model::mapiRowModel>& model)
	{
		if (!model) return;
		auto ulPropTag = model->ulPropTag();
		auto image = sortIcon::siDefault;
		for (const auto& _PropTypeIcon : _PropTypeIcons)
		{
			if (_PropTypeIcon.objType == PROP_TYPE(ulPropTag))
			{
				image = _PropTypeIcon.image;
				break;
			}
		}

		auto lpData = InsertRow(iRow, L"", 0, image);
		// Data used to refer to specific property tags. See GetSelectedPropTag.
		if (lpData)
		{
			sortlistdata::propModelData::init(lpData, model);
		}

		SetItemText(iRow, columns::pcPROPBESTGUESS, model->name());
		SetItemText(iRow, columns::pcPROPOTHERNAMES, model->otherName());
		SetItemText(iRow, columns::pcPROPTAG, model->tag());
		SetItemText(iRow, columns::pcPROPTYPE, model->propType());
		SetItemText(iRow, columns::pcPROPVAL, model->value());
		SetItemText(iRow, columns::pcPROPVALALT, model->altValue());
		SetItemText(iRow, columns::pcPROPSMARTVIEW, model->smartView());
		SetItemText(iRow, columns::pcPROPNAMEDNAME, model->namedPropName());
		SetItemText(iRow, columns::pcPROPNAMEDGUID, model->namedPropGuid());
	}

	_Check_return_ bool CSingleMAPIPropListCtrl::IsModifiedPropVals() const
	{
		return m_lpPropBag && m_lpPropBag->IsModified();
	}

	void CSingleMAPIPropListCtrl::SetDataSource(
		_In_opt_ LPMAPIPROP lpMAPIProp,
		_In_opt_ sortlistdata::sortListData* lpListData,
		bool bIsAB)
	{
		output::DebugPrintEx(output::dbgLevel::Generic, CLASS, L"SetDataSource", L"setting new data source\n");

		if (lpMAPIProp)
		{
			return SetDataSource(std::make_shared<propertybag::mapiPropPropertyBag>(lpMAPIProp, lpListData, bIsAB));
		}
		else if (lpListData)
		{
			return SetDataSource(std::make_shared<propertybag::rowPropertyBag>(lpListData, bIsAB));
		}

		return SetDataSource(nullptr);
	}

	// Clear the current property list from the control.
	// Load a new list from the IMAPIProp or lpSourceProps object passed in
	// Most calls to this will come through CBaseDialog::OnUpdateSingleMAPIPropListCtrl, which will preserve the current bIsAB
	// Exceptions will be where we need to set a specific bIsAB
	void CSingleMAPIPropListCtrl::SetDataSource(const std::shared_ptr<propertybag::IMAPIPropertyBag> lpPropBag)
	{
		output::DebugPrintEx(output::dbgLevel::Generic, CLASS, L"SetDataSource", L"setting new data source\n");

		// if nothing to do...do nothing
		if (lpPropBag && lpPropBag->IsEqual(m_lpPropBag))
		{
			return;
		}

		m_lpPropBag = lpPropBag;

		// Turn off redraw while we work on the window
		MySetRedraw(false);

		RefreshMAPIPropList();

		// Reset our header widths if weren't showing anything before and are now
		if (!m_bHaveEverDisplayedSomething && m_lpPropBag && GetItemCount())
		{
			m_bHaveEverDisplayedSomething = true;

			auto lpMyHeader = GetHeaderCtrl();

			if (lpMyHeader)
			{
				// This fixes a ton of flashing problems
				lpMyHeader->SetRedraw(true);
				for (auto iCurCol = 0; iCurCol < static_cast<int>(columns::PropColumns.size()); iCurCol++)
				{
					SetColumnWidth(iCurCol, LVSCW_AUTOSIZE_USEHEADER);
					if (GetColumnWidth(iCurCol) > 200) SetColumnWidth(iCurCol, 200);
				}

				lpMyHeader->SetRedraw(false);
			}
		}

		// Turn redraw back on to update our view
		MySetRedraw(true);
	}

	const std::shared_ptr<propertybag::IMAPIPropertyBag> CSingleMAPIPropListCtrl::GetDataSource()
	{
		return m_lpPropBag;
	}

	std::wstring binPropToXML(UINT uidTag, const std::wstring str, int iIndent)
	{
		auto toks = strings::tokenize(str);
		if (toks.count(L"lpb"))
		{
			auto attr = property::Attributes();
			if (toks.count(L"cb"))
			{
				attr.AddAttribute(L"cb", toks[L"cb"]);
			}

			auto parsing = property::Parsing(toks[L"lpb"], true, attr);
			return parsing.toXML(uidTag, iIndent);
		}

		return strings::emptystring;
	}

	// Parse this for XML:
	// Err: 0x00040380=MAPI_W_ERRORS_RETURNED
	std::wstring errPropToXML(UINT uidTag, const std::wstring str, int iIndent)
	{
		auto toks = strings::tokenize(str);
		if (toks.count(L"Err"))
		{
			auto err = strings::split(toks[L"Err"], L'=');
			if (err.size() == 2)
			{
				auto attr = property::Attributes();
				attr.AddAttribute(L"err", err[0]);

				auto parsing = property::Parsing(err[1], true, attr);
				return parsing.toXML(uidTag, iIndent);
			}
		}

		return strings::emptystring;
	}

	void CSingleMAPIPropListCtrl::SavePropsToXML()
	{
		auto szFileName = file::CFileDialogExW::SaveAs(
			L"xml", // STRING_OK
			L"props.xml", // STRING_OK
			OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT,
			strings::loadstring(IDS_XMLFILES),
			this);
		if (!szFileName.empty())
		{
			const auto fProps = output::MyOpenFile(szFileName, true);
			if (fProps)
			{
				output::DebugPrintEx(
					output::dbgLevel::Generic, CLASS, L"SavePropsToXML", L"saving to %ws\n", szFileName.c_str());

				// Force a sort on the tag column to make output consistent
				FakeClickColumn(columns::pcPROPTAG, false);

				const auto iItemCount = GetItemCount();

				output::OutputToFile(fProps, output::g_szXMLHeader);
				output::OutputToFile(fProps, L"<properties listtype=\"propertypane\">\n");
				for (auto iRow = 0; iRow < iItemCount; iRow++)
				{
					const auto szTag = GetItemText(iRow, columns::pcPROPTAG);
					const auto szType = GetItemText(iRow, columns::pcPROPTYPE);
					const auto szBestGuess = GetItemText(iRow, columns::pcPROPBESTGUESS);
					const auto szOther = GetItemText(iRow, columns::pcPROPOTHERNAMES);
					const auto szNameGuid = GetItemText(iRow, columns::pcPROPNAMEDGUID);
					const auto szNameName = GetItemText(iRow, columns::pcPROPNAMEDNAME);

					output::OutputToFilef(
						fProps, L"\t<property tag = \"%ws\" type = \"%ws\" >\n", szTag.c_str(), szType.c_str());

					if (szNameName.empty() && !strings::beginsWith(szBestGuess, L"0x"))
					{
						output::OutputXMLValueToFile(
							fProps, columns::PropXMLNames[columns::pcPROPBESTGUESS].uidName, szBestGuess, false, 2);
					}

					output::OutputXMLValueToFile(
						fProps, columns::PropXMLNames[columns::pcPROPOTHERNAMES].uidName, szOther, false, 2);

					output::OutputXMLValueToFile(
						fProps, columns::PropXMLNames[columns::pcPROPNAMEDGUID].uidName, szNameGuid, false, 2);

					output::OutputXMLValueToFile(
						fProps, columns::PropXMLNames[columns::pcPROPNAMEDNAME].uidName, szNameName, false, 2);

					const auto lpListData = reinterpret_cast<sortlistdata::sortListData*>(GetItemData(iRow));
					auto ulPropType = PT_NULL;

					if (lpListData)
					{
						const auto prop = lpListData->cast<sortlistdata::propModelData>();
						if (prop)
						{
							ulPropType = PROP_TYPE(prop->getPropTag());
						}
					}

					const auto szVal = GetItemText(iRow, columns::pcPROPVAL);
					const auto szAltVal = GetItemText(iRow, columns::pcPROPVALALT);
					switch (ulPropType)
					{
					case PT_STRING8:
					case PT_UNICODE:
					{
						output::OutputXMLValueToFile(
							fProps, columns::PropXMLNames[columns::pcPROPVAL].uidName, szVal, true, 2);
						const auto binXML =
							binPropToXML(columns::PropXMLNames[columns::pcPROPVALALT].uidName, szAltVal, 2);
						output::Output(output::dbgLevel::NoDebug, fProps, false, binXML);
					}
					break;
					case PT_BINARY:
					{
						const auto binXML = binPropToXML(columns::PropXMLNames[columns::pcPROPVAL].uidName, szVal, 2);
						output::Output(output::dbgLevel::NoDebug, fProps, false, binXML);
						output::OutputXMLValueToFile(
							fProps, columns::PropXMLNames[columns::pcPROPVALALT].uidName, szAltVal, true, 2);
					}
					break;
					case PT_ERROR:
					{
						const auto errXML = errPropToXML(columns::PropXMLNames[columns::pcPROPVAL].uidName, szVal, 2);
						output::Output(output::dbgLevel::NoDebug, fProps, false, errXML);
					}
					break;
					default:
						output::OutputXMLValueToFile(
							fProps, columns::PropXMLNames[columns::pcPROPVAL].uidName, szVal, false, 2);
						output::OutputXMLValueToFile(
							fProps, columns::PropXMLNames[columns::pcPROPVALALT].uidName, szAltVal, false, 2);
						break;
					}

					const auto szSmartView = GetItemText(iRow, columns::pcPROPSMARTVIEW);
					output::OutputXMLValueToFile(
						fProps, columns::PropXMLNames[columns::pcPROPSMARTVIEW].uidName, szSmartView, true, 2);

					output::OutputToFile(fProps, L"\t</property>\n");
				}

				output::OutputToFile(fProps, L"</properties>");
				output::CloseFile(fProps);
			}
		}
	}

	void CSingleMAPIPropListCtrl::OnDblclk(_In_ NMHDR* /*pNMHDR*/, _In_ LRESULT* pResult)
	{
		output::DebugPrintEx(output::dbgLevel::Generic, CLASS, L"OnDblclk", L"calling OnEditProp\n");
		OnEditProp();
		*pResult = 0;
	}

	void CSingleMAPIPropListCtrl::OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags)
	{
		output::DebugPrintEx(output::dbgLevel::Menu, CLASS, L"OnKeyDown", L"0x%X\n", nChar);

		const auto bCtrlPressed = GetKeyState(VK_CONTROL) < 0;
		const auto bShiftPressed = GetKeyState(VK_SHIFT) < 0;
		const auto bMenuPressed = GetKeyState(VK_MENU) < 0;

		if (!bMenuPressed)
		{
			if ('X' == nChar && bCtrlPressed && !bShiftPressed)
			{
				OnDeleteProperty();
			}
			else if (VK_DELETE == nChar)
			{
				OnDeleteProperty();
			}
			else if ('S' == nChar && bCtrlPressed)
			{
				SavePropsToXML();
			}
			else if ('E' == nChar && bCtrlPressed)
			{
				OnEditProp();
			}
			else if ('C' == nChar && bCtrlPressed && !bShiftPressed)
			{
				OnCopyProperty();
			}
			else if ('V' == nChar && bCtrlPressed && !bShiftPressed)
			{
				OnPasteProperty();
			}
			else if (VK_F5 == nChar)
			{
				RefreshMAPIPropList();
			}
			else if (VK_RETURN == nChar)
			{
				if (!bCtrlPressed)
				{
					output::DebugPrintEx(output::dbgLevel::Menu, CLASS, L"OnKeyDown", L"calling OnEditProp\n");
					OnEditProp();
				}
				else
				{
					output::DebugPrintEx(output::dbgLevel::Menu, CLASS, L"OnKeyDown", L"calling OnOpenProperty\n");
					OnOpenProperty();
				}
			}
			else if (!m_lpHostDlg || !m_lpHostDlg->HandleKeyDown(nChar, bShiftPressed, bCtrlPressed, bMenuPressed))
			{
				CSortListCtrl::OnKeyDown(nChar, nRepCnt, nFlags);
			}
		}
	}

	void CSingleMAPIPropListCtrl::OnContextMenu(_In_ CWnd* pWnd, CPoint pos)
	{
		if (pWnd && -1 == pos.x && -1 == pos.y)
		{
			POINT point = {0};
			const auto iItem = GetNextItem(-1, LVNI_SELECTED);
			if (GetItemPosition(iItem, &point) && ::ClientToScreen(pWnd->m_hWnd, &point))
			{
				pos = point;
			}
		}

		ui::DisplayContextMenu(IDR_MENU_PROPERTY_POPUP, IDR_MENU_MESSAGE_POPUP, m_lpHostDlg->m_hWnd, pos.x, pos.y);
	}

	void CSingleMAPIPropListCtrl::FindAllNamedProps()
	{
		if (!m_lpPropBag) return;

		output::DebugPrintEx(
			output::dbgLevel::NamedProp, CLASS, L"FindAllNamedProps", L"Calling GetNamesFromIDs with a NULL\n");
		const auto names = cache::GetNamesFromIDs(m_lpPropBag->GetMAPIProp(), nullptr, 0);
		if (names.size() > 0)
		{
			for (const auto& name : names)
			{
				AddPropToExtraProps(PROP_TAG(NULL, name->getPropID()), false);
			}

			// Refresh the display
			RefreshMAPIPropList();
		}
	}

	void CSingleMAPIPropListCtrl::CountNamedProps()
	{
		if (!m_lpPropBag) return;

		output::DebugPrintEx(
			output::dbgLevel::NamedProp, CLASS, L"CountNamedProps", L"Searching for the highest named prop mapping\n");

		ULONG ulHighestKnown = cache::FindHighestNamedProp(m_lpPropBag->GetMAPIProp());

		dialog::editor::CEditor MyResult(this, IDS_COUNTNAMEDPROPS, IDS_COUNTNAMEDPROPSPROMPT, CEDITOR_BUTTON_OK);
		if (ulHighestKnown)
		{
			const auto ulPropTag = PROP_TAG(NULL, ulHighestKnown);
			const auto name = cache::GetNameFromID(m_lpPropBag->GetMAPIProp(), ulPropTag, NULL);
			if (cache::namedPropCacheEntry::valid(name))
			{
				output::DebugPrintEx(
					output::dbgLevel::NamedProp,
					CLASS,
					L"CountNamedProps",
					L"Found a named property at 0x%04X.\n",
					ulHighestKnown);
			}

			MyResult.AddPane(viewpane::TextPane::CreateSingleLinePane(0, IDS_HIGHESTNAMEDPROPTOTAL, true));
			MyResult.SetDecimal(0, ulHighestKnown - 0x8000);

			MyResult.AddPane(viewpane::TextPane::CreateMultiLinePane(1, IDS_HIGHESTNAMEDPROPNUM, true));

			if (cache::namedPropCacheEntry::valid(name))
			{
				const auto namePropNames =
					cache::NameIDToStrings(ulPropTag, nullptr, name->getMapiNameId(), nullptr, false);
				MyResult.SetStringW(
					1,
					strings::formatmessage(
						IDS_HIGHESTNAMEDPROPNAME,
						ulHighestKnown,
						namePropNames.name.c_str(),
						namePropNames.guid.c_str()));
			}
		}
		else
		{
			MyResult.AddPane(viewpane::TextPane::CreateSingleLinePane(0, IDS_HIGHESTNAMEDPROPTOTAL, true));
			MyResult.LoadString(0, IDS_HIGHESTNAMEDPROPNOTFOUND);
		}

		static_cast<void>(MyResult.DisplayDialog());
	}

	// Delete the selected property
	void CSingleMAPIPropListCtrl::OnDeleteProperty()
	{
		if (!m_lpPropBag || m_lpPropBag->GetType() == propertybag::propBagType::Row) return;
		auto lpPropBag = m_lpPropBag; // Hold the prop bag so it doesn't get deleted under us

		const auto propModelData = GetSelectedPropModelData();
		if (!propModelData) return;

		dialog::editor::CEditor Query(
			this, IDS_DELETEPROPERTY, IDS_DELETEPROPERTYPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
		if (Query.DisplayDialog())
		{
			output::DebugPrintEx(
				output::dbgLevel::Generic,
				CLASS,
				L"OnDeleteProperty",
				L"deleting property 0x%08X\n",
				propModelData->getPropTag());

			const auto hRes = EC_H(lpPropBag->DeleteProp(propModelData->getPropTag(), propModelData->getName()));
			if (SUCCEEDED(hRes))
			{
				// Refresh the display
				RefreshMAPIPropList();
			}
		}
	}

	// Display the selected property as a security descriptor using a property sheet
	void CSingleMAPIPropListCtrl::OnDisplayPropertyAsSecurityDescriptorPropSheet() const
	{
		if (!m_lpPropBag || !import::pfnEditSecurity) return;

		const auto propModelData = GetSelectedPropModelData();
		if (!propModelData) return;

		output::DebugPrintEx(
			output::dbgLevel::Generic,
			CLASS,
			L"OnDisplayPropertyAsSecurityDescriptorPropSheet",
			L"interpreting 0x%X as Security Descriptor\n",
			propModelData->getPropTag());

		const auto mySecInfo =
			std::make_shared<mapi::mapiui::CMySecInfo>(m_lpPropBag->GetMAPIProp(), propModelData->getPropTag());

		EC_B_S(import::pfnEditSecurity(m_hWnd, mySecInfo.get()));
	}

	void CSingleMAPIPropListCtrl::OnEditProp()
	{
		if (!m_lpPropBag) return;
		const auto propModelData = GetSelectedPropModelData();
		if (!propModelData) return;
		OnEditGivenProp(propModelData->getPropTag(), propModelData->getName());
	}

	void CSingleMAPIPropListCtrl::OnEditPropAsRestriction(ULONG ulPropTag)
	{
		if (!m_lpPropBag || !ulPropTag || PT_SRESTRICTION != PROP_TYPE(ulPropTag)) return;
		auto lpPropBag = m_lpPropBag; // Hold the prop bag so it doesn't get deleted under us

		auto lpEditProp = lpPropBag->GetOneProp(ulPropTag, {}); // TODO: Should I have a name here?

		LPSRestriction lpResIn = nullptr;
		if (lpEditProp)
		{
			lpResIn = reinterpret_cast<LPSRestriction>(lpEditProp->Value.lpszA);
		}

		output::DebugPrint(output::dbgLevel::Generic, L"Source restriction before editing:\n");
		output::outputRestriction(output::dbgLevel::Generic, nullptr, lpResIn, lpPropBag->GetMAPIProp());
		dialog::editor::RestrictEditor MyResEditor(
			this,
			m_lpMapiObjects,
			nullptr, // No alloc parent - we must MAPIFreeBuffer the result
			lpResIn);
		if (MyResEditor.DisplayDialog())
		{
			const auto lpModRes = MyResEditor.DetachModifiedSRestriction();
			if (lpModRes)
			{
				output::DebugPrint(output::dbgLevel::Generic, L"Modified restriction:\n");
				output::outputRestriction(output::dbgLevel::Generic, nullptr, lpModRes, lpPropBag->GetMAPIProp());

				// need to merge the data we got back from the RestrictEditor with our current prop set
				// so that we can free lpModRes
				SPropValue ResProp = {};
				if (lpEditProp)
				{
					ResProp.ulPropTag = lpEditProp->ulPropTag;
				}
				else
				{
					ResProp.ulPropTag = ulPropTag;
				}

				ResProp.Value.lpszA = reinterpret_cast<LPSTR>(lpModRes);

				const auto hRes = EC_H(lpPropBag->SetProp(&ResProp, ulPropTag, L""));

				// Remember, we had no alloc parent - this is safe to free
				MAPIFreeBuffer(lpModRes);

				if (SUCCEEDED(hRes))
				{
					// refresh
					RefreshMAPIPropList();
				}
			}
		}

		lpPropBag->FreeBuffer(lpEditProp);
	}

	void CSingleMAPIPropListCtrl::OnEditGivenProp(ULONG ulPropTag, const std::wstring& name)
	{
		LPSPropValue lpEditProp = nullptr;

		if (!m_lpPropBag) return;
		auto lpPropBag = m_lpPropBag; // Hold the prop bag so it doesn't get deleted under us

		// Explicit check since TagToString is expensive
		if (fIsSet(output::dbgLevel::Generic))
		{
			output::DebugPrintEx(
				output::dbgLevel::Generic,
				CLASS,
				L"OnEditGivenProp",
				L"editing property 0x%X (= %ws = %ws)\n",
				ulPropTag,
				name.c_str(),
				proptags::TagToString(ulPropTag, lpPropBag->GetMAPIProp(), lpPropBag->IsAB(), true).c_str());
		}

		ulPropTag = PT_ERROR == PROP_TYPE(ulPropTag) ? CHANGE_PROP_TYPE(ulPropTag, PT_UNSPECIFIED) : ulPropTag;

		if (PT_SRESTRICTION == PROP_TYPE(ulPropTag))
		{
			OnEditPropAsRestriction(ulPropTag);
			return;
		}

		if (PT_OBJECT == PROP_TYPE(ulPropTag))
		{
			EC_H_S(DisplayTable(lpPropBag->GetMAPIProp(), ulPropTag, dialog::objectType::otDefault, m_lpHostDlg));
			return;
		}

		const auto lpSourceObj = lpPropBag->GetMAPIProp();

		auto bUseStream = false;

		if (PROP_ID(PR_RTF_COMPRESSED) == PROP_ID(ulPropTag))
		{
			bUseStream = true;
		}
		else
		{
			lpEditProp = lpPropBag->GetOneProp(ulPropTag, name);
		}

		if (lpEditProp && PROP_TYPE(lpEditProp->ulPropTag) == PT_ERROR &&
			lpEditProp->Value.err == MAPI_E_NOT_ENOUGH_MEMORY)
		{
			bUseStream = true;
		}

		if (bUseStream)
		{
			dialog::editor::CStreamEditor MyEditor(
				this,
				IDS_PROPEDITOR,
				IDS_STREAMEDITORPROMPT,
				lpSourceObj,
				ulPropTag,
				true, // Guess the type of stream to use
				lpPropBag->IsAB(),
				false,
				false,
				NULL,
				NULL,
				NULL);

			if (MyEditor.DisplayDialog())
			{
				RefreshMAPIPropList();
			}
		}
		else
		{
			if (PROP_TYPE(ulPropTag) == PT_UNSPECIFIED && lpEditProp) ulPropTag = lpEditProp->ulPropTag;

			const auto propEditor = dialog::editor::DisplayPropertyEditor(
				this,
				m_lpMapiObjects,
				IDS_PROPEDITOR,
				name,
				lpPropBag->IsAB(),
				lpSourceObj,
				ulPropTag,
				false,
				lpEditProp);
			if (propEditor)
			{
				const auto lpModProp = propEditor->getValue();
				if (lpModProp)
				{
					// If we didn't have a source object, we need to shove our results back in to the property bag
					if (!lpSourceObj)
					{
						// SetProp does not take ownership of memory
						EC_H_S(lpPropBag->SetProp(lpModProp, ulPropTag, name));
					}

					RefreshMAPIPropList();
				}
			}
		}

		lpPropBag->FreeBuffer(lpEditProp);
	}

	// Display the selected property as a stream using CStreamEditor
	void CSingleMAPIPropListCtrl::OnEditPropAsStream(ULONG ulType, bool bEditAsRTF)
	{
		if (!m_lpPropBag) return;
		auto lpPropBag = m_lpPropBag; // Hold the prop bag so it doesn't get deleted under us

		const auto propModelData = GetSelectedPropModelData();
		if (!propModelData) return;
		auto ulPropTag = propModelData->getPropTag();

		// Explicit check since TagToString is expensive
		if (fIsSet(output::dbgLevel::Generic))
		{
			output::DebugPrintEx(
				output::dbgLevel::Generic,
				CLASS,
				L"OnEditPropAsStream",
				L"editing property 0x%X (= %ws) as stream, ulType = 0x%08X, bEditAsRTF = 0x%X\n",
				ulPropTag,
				proptags::TagToString(ulPropTag, lpPropBag->GetMAPIProp(), lpPropBag->IsAB(), true).c_str(),
				ulType,
				bEditAsRTF);
		}

		ulPropTag = CHANGE_PROP_TYPE(ulPropTag, ulType);

		auto bUseWrapEx = false;
		ULONG ulRTFFlags = NULL;
		ULONG ulInCodePage = NULL;
		ULONG ulOutCodePage = CP_ACP; // Default to ANSI - check if this is valid for UNICODE builds

		if (bEditAsRTF)
		{
			dialog::editor::CEditor MyPrompt(
				this, IDS_USEWRAPEX, IDS_USEWRAPEXPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
			MyPrompt.AddPane(viewpane::CheckPane::Create(0, IDS_USEWRAPEX, true, false));

			if (!MyPrompt.DisplayDialog()) return;

			if (MyPrompt.GetCheck(0))
			{
				bUseWrapEx = true;
				const auto lpProp = lpPropBag->GetOneProp(PR_INTERNET_CPID, {});
				if (lpProp && PT_LONG == PROP_TYPE(lpProp[0].ulPropTag))
				{
					ulInCodePage = lpProp[0].Value.l;
				}

				lpPropBag->FreeBuffer(lpProp);

				dialog::editor::CEditor MyPrompt2(
					this, IDS_WRAPEXFLAGS, IDS_WRAPEXFLAGSPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
				MyPrompt2.AddPane(viewpane::TextPane::CreateSingleLinePane(0, IDS_WRAPEXFLAGS, false));
				MyPrompt2.SetHex(0, MAPI_NATIVE_BODY);
				MyPrompt2.AddPane(viewpane::TextPane::CreateSingleLinePane(1, IDS_ULINCODEPAGE, false));
				MyPrompt2.SetDecimal(1, ulInCodePage);
				MyPrompt2.AddPane(viewpane::TextPane::CreateSingleLinePane(2, IDS_ULOUTCODEPAGE, false));
				MyPrompt2.SetDecimal(2, CP_UNICODE);

				if (!MyPrompt2.DisplayDialog()) return;

				ulRTFFlags = MyPrompt2.GetHex(0);
				ulInCodePage = MyPrompt2.GetDecimal(1);
				ulOutCodePage = MyPrompt2.GetDecimal(2);
			}
		}

		dialog::editor::CStreamEditor MyEditor(
			this,
			IDS_PROPEDITOR,
			IDS_STREAMEDITORPROMPT,
			lpPropBag->GetMAPIProp(),
			ulPropTag,
			false, // No stream guessing
			lpPropBag->IsAB(),
			bEditAsRTF,
			bUseWrapEx,
			ulRTFFlags,
			ulInCodePage,
			ulOutCodePage);

		if (MyEditor.DisplayDialog())
		{
			RefreshMAPIPropList();
		}
	}

	void CSingleMAPIPropListCtrl::OnCopyProperty() const
	{
		cache::CGlobalCache::getInstance().SetPropertyToCopy(GetSelectedPropModelData(), m_lpPropBag);
	}

	void CSingleMAPIPropListCtrl::OnPasteProperty()
	{
		if (!m_lpHostDlg || !m_lpPropBag) return;
		auto lpPropBag = m_lpPropBag; // Hold the prop bag so it doesn't get deleted under us

		const auto sourceProp = cache::CGlobalCache::getInstance().GetPropertyToCopy();
		const auto sourcePropBag = cache::CGlobalCache::getInstance().GetSourcePropBag();
		const auto ulSourcePropTag = sourceProp->getPropTag();
		auto lpSourcePropObj = sourcePropBag->GetMAPIProp();
		if (!lpSourcePropObj)
		{
			// If we don't have a source prop object, try setting the prop back to the prop bag
			const auto sourceName = sourceProp->getName();
			const auto lpSourceProp = sourcePropBag->GetOneProp(ulSourcePropTag, sourceName);
			EC_H_S(lpPropBag->SetProp(lpSourceProp, ulSourcePropTag, sourceName));
			this->RefreshMAPIPropList();
			return;
		}

		auto hRes = S_OK;
		LPSPropProblemArray lpProblems = nullptr;
		SPropTagArray TagArray = {1, ulSourcePropTag};

		dialog::editor::CEditor MyData(
			this, IDS_PASTEPROP, IDS_PASTEPROPPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);

		UINT uidDropDown[] = {IDS_DDCOPYPROPS, IDS_DDGETSETPROPS, IDS_DDCOPYSTREAM};
		MyData.AddPane(viewpane::DropDownPane::Create(0, IDS_COPYSTYLE, _countof(uidDropDown), uidDropDown, true));
		MyData.AddPane(viewpane::TextPane::CreateSingleLinePane(1, IDS_SOURCEPROP, false));
		MyData.SetHex(1, ulSourcePropTag);
		MyData.AddPane(viewpane::TextPane::CreateSingleLinePane(2, IDS_TARGETPROP, false));
		MyData.SetHex(2, ulSourcePropTag);

		if (!MyData.DisplayDialog()) return;

		const auto ulSourceTag = MyData.GetHex(1);
		auto ulTargetTag = MyData.GetHex(2);
		mapi::setTag(TagArray, 0) = ulSourceTag;

		if (PROP_TYPE(ulTargetTag) != PROP_TYPE(ulSourceTag))
			ulTargetTag = CHANGE_PROP_TYPE(ulTargetTag, PROP_TYPE(ulSourceTag));

		switch (MyData.GetDropDown(0))
		{
		case 0:
		{
			dialog::editor::CEditor MyCopyData(
				this, IDS_PASTEPROP, IDS_COPYPASTEPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);

			const auto szGuid = guid::GUIDToStringAndName(&IID_IMAPIProp);
			MyCopyData.AddPane(viewpane::TextPane::CreateSingleLinePane(0, IDS_INTERFACE, szGuid, false));
			MyCopyData.AddPane(viewpane::TextPane::CreateSingleLinePane(1, IDS_FLAGS, false));
			MyCopyData.SetHex(1, MAPI_DIALOG);

			if (!MyCopyData.DisplayDialog()) return;

			auto MyGUID = guid::StringToGUID(MyCopyData.GetStringW(0));
			auto lpProgress = mapi::mapiui::GetMAPIProgress(L"IMAPIProp::CopyProps", m_lpHostDlg->m_hWnd); // STRING_OK
			auto ulCopyFlags = MyCopyData.GetHex(1);

			if (lpProgress) ulCopyFlags |= MAPI_DIALOG;

			hRes = EC_MAPI(lpSourcePropObj->CopyProps(
				&TagArray,
				lpProgress ? reinterpret_cast<ULONG_PTR>(m_lpHostDlg->m_hWnd) : NULL, // ui param
				lpProgress, // progress
				&MyGUID,
				lpPropBag->GetMAPIProp(),
				ulCopyFlags,
				&lpProblems));

			if (lpProgress) lpProgress->Release();
		}
		break;
		case 1:
		{
			ULONG ulValues = NULL;
			LPSPropValue lpSourceProp = nullptr;
			hRes = EC_MAPI(lpSourcePropObj->GetProps(&TagArray, fMapiUnicode, &ulValues, &lpSourceProp));
			if (SUCCEEDED(hRes) && ulValues && lpSourceProp && PT_ERROR != lpSourceProp->ulPropTag)
			{
				lpSourceProp->ulPropTag = ulTargetTag;
				hRes = EC_H(lpPropBag->SetProps(ulValues, lpSourceProp));
			}
		}
		break;
		case 2:
			hRes =
				EC_H(mapi::CopyPropertyAsStream(lpSourcePropObj, lpPropBag->GetMAPIProp(), ulSourceTag, ulTargetTag));
			break;
		}

		EC_PROBLEMARRAY(lpProblems);
		MAPIFreeBuffer(lpProblems);

		if (SUCCEEDED(hRes))
		{
			hRes = EC_H(lpPropBag->Commit());

			if (SUCCEEDED(hRes))
			{
				// refresh
				RefreshMAPIPropList();
			}
		}

		lpSourcePropObj->Release();
	}

	void CSingleMAPIPropListCtrl::OnCopyTo()
	{
		// for now, we only copy from objects - copying from rows would be difficult to generalize
		if (!m_lpHostDlg || !m_lpPropBag) return;

		auto lpSourcePropObj = cache::CGlobalCache::getInstance().GetSourcePropObject();
		if (!lpSourcePropObj) return;

		const auto hRes = EC_H(mapi::CopyTo(
			m_lpHostDlg->m_hWnd,
			lpSourcePropObj,
			m_lpPropBag->GetMAPIProp(),
			&IID_IMAPIProp,
			nullptr,
			m_lpPropBag->IsAB(),
			true));
		if (SUCCEEDED(hRes))
		{
			EC_H_S(m_lpPropBag->Commit());

			// refresh
			RefreshMAPIPropList();
		}

		lpSourcePropObj->Release();
	}

	// Open a binary property as an entry ID
	void CSingleMAPIPropListCtrl::OnOpenProperty() const
	{
		auto hRes = S_OK;

		if (!m_lpHostDlg) return;

		const auto propModelData = GetSelectedPropModelData();
		if (!propModelData) return;
		const auto ulPropTag = propModelData->getPropTag();

		output::DebugPrintEx(output::dbgLevel::Generic, CLASS, L"OnOpenProperty", L"asked to open 0x%X\n", ulPropTag);
		LPSPropValue lpProp = nullptr;
		auto lpPropBag = m_lpPropBag; // Hold the prop bag so it doesn't get deleted under us
		if (lpPropBag)
		{
			lpProp = lpPropBag->GetOneProp(ulPropTag, {}); // TODO: Should I have a name here
		}

		if (SUCCEEDED(hRes) && lpProp)
		{
			if (lpPropBag && PT_OBJECT == PROP_TYPE(lpProp->ulPropTag))
			{
				EC_H_S(DisplayTable(
					lpPropBag->GetMAPIProp(), lpProp->ulPropTag, dialog::objectType::otDefault, m_lpHostDlg));
			}
			else if (PT_BINARY == PROP_TYPE(lpProp->ulPropTag) || PT_MV_BINARY == PROP_TYPE(lpProp->ulPropTag))
			{
				switch (PROP_TYPE(lpProp->ulPropTag))
				{
				case PT_BINARY:
					output::DebugPrintEx(
						output::dbgLevel::Generic, CLASS, L"OnOpenProperty", L"property is PT_BINARY\n");
					m_lpHostDlg->OnOpenEntryID(mapi::getBin(lpProp));
					break;
				case PT_MV_BINARY:
					output::DebugPrintEx(
						output::dbgLevel::Generic, CLASS, L"OnOpenProperty", L"property is PT_MV_BINARY\n");
					if (hRes == S_OK && lpProp && PT_MV_BINARY == PROP_TYPE(lpProp->ulPropTag))
					{
						output::DebugPrintEx(
							output::dbgLevel::Generic,
							CLASS,
							L"OnOpenProperty",
							L"opened MV structure. There are 0x%X binaries in it.\n",
							lpProp->Value.MVbin.cValues);
						for (ULONG i = 0; i < lpProp->Value.MVbin.cValues; i++)
						{
							m_lpHostDlg->OnOpenEntryID(lpProp->Value.MVbin.lpbin[i]);
						}
					}
					break;
				}
			}
		}

		if (lpPropBag)
		{
			lpPropBag->FreeBuffer(lpProp);
		}
	}

	void CSingleMAPIPropListCtrl::OnModifyExtraProps()
	{
		cache::CGlobalCache::getInstance().MAPIInitialize(NULL);

		dialog::editor::CTagArrayEditor MyTagArrayEditor(
			this,
			IDS_EXTRAPROPS,
			NULL,
			nullptr,
			m_sptExtraProps,
			m_lpPropBag ? m_lpPropBag->IsAB() : false,
			m_lpPropBag ? m_lpPropBag->GetMAPIProp() : nullptr);

		if (!MyTagArrayEditor.DisplayDialog()) return;

		const auto lpNewTagArray = MyTagArrayEditor.DetachModifiedTagArray();
		if (lpNewTagArray)
		{
			MAPIFreeBuffer(m_sptExtraProps);
			m_sptExtraProps = lpNewTagArray;
		}

		RefreshMAPIPropList();
	}

	void CSingleMAPIPropListCtrl::OnEditGivenProperty()
	{
		if (!m_lpPropBag) return;

		// Display a dialog to get a property number.
		dialog::editor::CPropertyTagEditor MyPropertyTag(
			IDS_EDITGIVENPROP,
			NULL, // prompt
			NULL,
			m_lpPropBag->IsAB(),
			m_lpPropBag->GetMAPIProp(),
			this);

		if (MyPropertyTag.DisplayDialog())
		{
			OnEditGivenProp(MyPropertyTag.GetPropertyTag(), {});
		}
	}

	void CSingleMAPIPropListCtrl::OnOpenPropertyAsTable()
	{
		if (!m_lpPropBag) return;
		auto lpPropBag = m_lpPropBag; // Hold the prop bag so it doesn't get deleted under us

		// Display a dialog to get a property number.
		dialog::editor::CPropertyTagEditor MyPropertyTag(
			IDS_OPENPROPASTABLE,
			NULL, // prompt
			NULL,
			lpPropBag->IsAB(),
			lpPropBag->GetMAPIProp(),
			this);
		if (!MyPropertyTag.DisplayDialog()) return;
		dialog::editor::CEditor MyData(
			this, IDS_OPENPROPASTABLE, IDS_OPENPROPASTABLEPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);

		MyData.AddPane(viewpane::CheckPane::Create(0, IDS_OPENASEXTABLE, false, false));
		if (!MyData.DisplayDialog()) return;

		if (MyData.GetCheck(0))
		{
			EC_H_S(DisplayExchangeTable(
				lpPropBag->GetMAPIProp(),
				CHANGE_PROP_TYPE(MyPropertyTag.GetPropertyTag(), PT_OBJECT),
				dialog::objectType::otDefault,
				m_lpHostDlg));
		}
		else
		{
			EC_H_S(DisplayTable(
				lpPropBag->GetMAPIProp(),
				CHANGE_PROP_TYPE(MyPropertyTag.GetPropertyTag(), PT_OBJECT),
				dialog::objectType::otDefault,
				m_lpHostDlg));
		}
	}

	void CSingleMAPIPropListCtrl::OnPasteNamedProps()
	{
		if (!m_lpPropBag) return;
		auto lpPropBag = m_lpPropBag; // Hold the prop bag so it doesn't get deleted under us

		const auto lpSourceMsgEID = cache::CGlobalCache::getInstance().GetMessagesToCopy();

		if (cache::CGlobalCache::getInstance().GetBufferStatus() & BUFFER_MESSAGES && lpSourceMsgEID &&
			1 == lpSourceMsgEID->cValues)
		{
			dialog::editor::CEditor MyData(
				this, IDS_PASTENAMEDPROPS, IDS_PASTENAMEDPROPSPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);

			const auto szGuid = guid::GUIDToStringAndName(&PS_PUBLIC_STRINGS);
			MyData.AddPane(viewpane::TextPane::CreateSingleLinePane(0, IDS_GUID, szGuid, false));
			MyData.AddPane(viewpane::CheckPane::Create(1, IDS_MAPIMOVE, false, false));
			MyData.AddPane(viewpane::CheckPane::Create(2, IDS_MAPINOREPLACE, false, false));

			if (!MyData.DisplayDialog()) return;

			ULONG ulObjType = 0;
			auto propSetGUID = guid::StringToGUID(MyData.GetStringW(0));

			auto lpSource = mapi::CallOpenEntry<LPMAPIPROP>(
				nullptr,
				nullptr,
				cache::CGlobalCache::getInstance().GetSourceParentFolder(),
				nullptr,
				lpSourceMsgEID->lpbin,
				nullptr,
				MAPI_BEST_ACCESS,
				&ulObjType);
			if (ulObjType == MAPI_MESSAGE && lpSource)
			{
				auto hRes = EC_H(mapi::CopyNamedProps(
					lpSource,
					&propSetGUID,
					MyData.GetCheck(1),
					MyData.GetCheck(2),
					lpPropBag->GetMAPIProp(),
					m_lpHostDlg->m_hWnd));
				if (SUCCEEDED(hRes))
				{
					hRes = EC_H(lpPropBag->Commit());
				}

				if (SUCCEEDED(hRes))
				{
					RefreshMAPIPropList();
				}

				lpSource->Release();
			}
		}
	}

	_Check_return_ bool CSingleMAPIPropListCtrl::HandleAddInMenu(WORD wMenuSelect) const
	{
		if (wMenuSelect < ID_ADDINPROPERTYMENU) return false;
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		const auto lpAddInMenu = ui::addinui::GetAddinMenuItem(m_lpHostDlg->m_hWnd, wMenuSelect);
		if (!lpAddInMenu) return false;

		_AddInMenuParams MyAddInMenuParams = {nullptr};
		MyAddInMenuParams.lpAddInMenu = lpAddInMenu;
		MyAddInMenuParams.ulAddInContext = MENU_CONTEXT_PROPERTY;
		MyAddInMenuParams.hWndParent = m_hWnd;
		MyAddInMenuParams.lpMAPIProp = m_lpPropBag->GetMAPIProp();
		if (m_lpMapiObjects)
		{
			MyAddInMenuParams.lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
			MyAddInMenuParams.lpMDB = m_lpMapiObjects->GetMDB(); // do not release
			MyAddInMenuParams.lpAdrBook = m_lpMapiObjects->GetAddrBook(false); // do not release
		}
		if (m_lpPropBag)
		{
			const auto lpRowPropBag = std::dynamic_pointer_cast<propertybag::rowPropertyBag>(m_lpPropBag);
			if (lpRowPropBag)
			{
				SRow MyRow = {0};
				auto hRes = WC_H(lpRowPropBag->GetAllProps(&MyRow.cValues, &MyRow.lpProps)); // No need to free
				if (SUCCEEDED(hRes))
				{
					MyAddInMenuParams.lpRow = &MyRow;
					MyAddInMenuParams.ulCurrentFlags |= MENU_FLAGS_ROW;
				}
			}
		}

		const auto propModelData = GetSelectedPropModelData();
		if (propModelData)
		{
			MyAddInMenuParams.ulPropTag = propModelData->getPropTag();
		}

		ui::addinui::InvokeAddInMenu(&MyAddInMenuParams);
		return true;
	}
} // namespace controls::sortlistctrl