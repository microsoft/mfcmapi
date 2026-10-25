#include <StdAfx.h>
#include <UI/Dialogs/ContentsTable/BinaryTableDlg.h>
#include <UI/Dialogs/Editors/Editor.h>
#include <UI/Controls/SortList/ContentsTableListCtrl.h>
#include <core/mapi/cache/mapiObjects.h>
#include <core/sortlistdata/contentsData.h>
#include <core/property/parseProperty.h>
#include <core/utility/strings.h>
#include <core/interpret/proptags.h>

namespace dialog
{
	_Check_return_ HRESULT LoadBinaryTable(
		_In_ LPMAPIPROP message,
		std::shared_ptr<cache::CMapiObjects> objects,
		bool entryIDs,
		ULONG propTag,
		std::vector<mapi::binary::BinaryTableRow>& members,
		_Deref_out_ LPMAPITABLE* table)
	{
		*table = nullptr;
		auto hRes = WC_H(mapi::binary::ReadBinaryTableRows(message, propTag, members, entryIDs));
		if (FAILED(hRes)) return hRes;

		if (entryIDs)
		{
			for (auto& member : members)
			{
				auto object = mapi::binary::OpenEntryID(objects->GetAddrBook(true), objects->GetSession(), member);
				if (!object) continue;

				SizedSPropTagArray(3, tags) = {3, {PR_DISPLAY_NAME_W, PR_ADDRTYPE_W, PR_EMAIL_ADDRESS_W}};
				ULONG count = 0;
				LPSPropValue properties = nullptr;
				hRes = WC_H_GETPROPS(
					object->GetProps(reinterpret_cast<LPSPropTagArray>(&tags), MAPI_UNICODE, &count, &properties));
				if (SUCCEEDED(hRes) && properties)
				{
					std::wstring* values[] = {&member.name, &member.addressType, &member.address};
					for (ULONG j = 0; j < count && j < _countof(values); ++j)
					{
						if (properties[j].ulPropTag == tags.aulPropTag[j] && properties[j].Value.lpszW)
						{
							*values[j] = properties[j].Value.lpszW;
						}
					}
				}

				MAPIFreeBuffer(properties);
				object->Release();
			}
		}

		return WC_H(mapi::binary::CreateBinaryTable(members, table, entryIDs));
	}

	_Check_return_ HRESULT
	DisplayBinaryTable(_In_ LPMAPIPROP message, ULONG propTag, bool entryIDs, bool isAB, _In_ CBaseDialog* parent)
	{
		if (!message || !parent) return WC_H(MAPI_E_INVALID_PARAMETER);
		if (PROP_TYPE(propTag) != PT_MV_BINARY) return WC_H(MAPI_E_INVALID_TYPE);

		std::vector<mapi::binary::BinaryTableRow> members;
		LPMAPITABLE table = nullptr;
		const auto hRes = WC_H(LoadBinaryTable(message, parent->GetMapiObjects(), entryIDs, propTag, members, &table));
		if (SUCCEEDED(hRes))
		{
			new CBinaryTableDlg(parent->GetMapiObjects(), message, propTag, entryIDs, table, std::move(members), isAB);
		}

		if (table) table->Release();

		return hRes;
	}

	_Check_return_ HRESULT DisplayBinaryTable(const SPropValue& property, bool entryIDs, _In_ CBaseDialog* parent)
	{
		if (!parent) return WC_H(MAPI_E_INVALID_PARAMETER);

		std::vector<mapi::binary::BinaryTableRow> rows;
		auto hRes = WC_H(mapi::binary::ExtractBinaryTableRows(property, property.ulPropTag, rows, entryIDs));
		if (FAILED(hRes)) return hRes;

		LPMAPITABLE table = nullptr;
		hRes = WC_H(mapi::binary::CreateBinaryTable(rows, &table, entryIDs));
		if (SUCCEEDED(hRes))
		{
			new CBinaryTableDlg(
				parent->GetMapiObjects(), nullptr, property.ulPropTag, entryIDs, table, std::move(rows));
		}

		if (table) table->Release();

		return hRes;
	}

	CBinaryTableDlg::CBinaryTableDlg(
		std::shared_ptr<cache::CMapiObjects> objects,
		_In_opt_ LPMAPIPROP message,
		ULONG propTag,
		bool entryIDs,
		_In_ LPMAPITABLE table,
		std::vector<mapi::binary::BinaryTableRow> members,
		bool isAB)
		: CContentsTableDlg(
			  objects,
			  entryIDs ? IDS_ENTRYIDTABLE : IDS_BINARYTABLE,
			  createDialogType::DO_NOT_CALL_CREATE_DIALOG,
			  nullptr,
			  table,
			  entryIDs ? &columns::sptEntryIDCols.tags : &columns::sptBinaryCols.tags,
			  entryIDs ? columns::EntryIDColumns : columns::BinaryColumns,
			  IDR_MENU_BINARYTABLE_POPUP,
			  0),
		  m_message(message), m_entryIDs(entryIDs), m_propTag(propTag), m_members(std::move(members))
	{
		if (m_message) m_message->AddRef();
		m_szTitle += L" - " + proptags::TagToString(propTag, message, isAB, true);
		CContentsTableDlg::CreateDialogAndMenu(IDR_MENU_BINARYTABLE);
	}

	CBinaryTableDlg::~CBinaryTableDlg()
	{
		if (m_message) m_message->Release();
	}

	BOOL CBinaryTableDlg::OnInitDialog()
	{
		const auto result = CContentsTableDlg::OnInitDialog();
		if (m_lpContentsTableListCtrl) m_lpContentsTableListCtrl->SetNamedColumnsOnly();

		return result;
	}

	_Check_return_ LPMAPIPROP CBinaryTableDlg::OpenItemProp(int, modifyType)
	{
		// Keep the property pane on the snapshot row, never a writable provider object.
		return nullptr;
	}

	void CBinaryTableDlg::OnDisplayItem()
	{
		if (!m_lpContentsTableListCtrl) return;

		for (const auto data : m_lpContentsTableListCtrl->GetSelectedItemData())
		{
			const auto contents = data ? data->cast<sortlistdata::contentsData>() : nullptr;
			if (!contents) continue;

			auto rowID = contents->getRowID();
			if (!rowID)
			{
				const auto key = contents->getInstanceKey();
				if (key && key->cb == sizeof(rowID) && key->lpb)
				{
					if (memcpy_s(&rowID, sizeof(rowID), key->lpb, key->cb))
					{
						EC_H_S(MAPI_E_INVALID_PARAMETER);
						continue;
					}
				}
			}

			if (!rowID || rowID > m_members.size())
			{
				EC_H_S(MAPI_E_NOT_FOUND);
				continue;
			}

			auto& member = m_members[rowID - 1];
			auto details = member.decoded;
			auto object = m_entryIDs ? mapi::binary::OpenEntryID(
										   m_lpMapiObjects->GetAddrBook(true), m_lpMapiObjects->GetSession(), member)
									 : nullptr;
			if (m_entryIDs)
			{
				details += L"\r\n" + strings::loadstring(IDS_ENTRYIDOPENSTATUS) + L": " +
						   strings::format(L"0x%08X", member.status);
			}
			if (object)
			{
				ULONG count = 0;
				LPSPropValue properties = nullptr;
				const auto hRes = WC_H_GETPROPS(object->GetProps(nullptr, MAPI_UNICODE, &count, &properties));
				details += strings::formatmessage(IDS_ENTRYIDGETPROPSSTATUS, hRes);
				if (SUCCEEDED(hRes) && properties)
				{
					for (ULONG i = 0; i < count; ++i)
					{
						std::wstring value;
						property::parseProperty(&properties[i], &value, nullptr);
						details += strings::format(L"\r\n0x%08X: ", properties[i].ulPropTag) + value;
					}
				}

				MAPIFreeBuffer(properties);
				object->Release();
			}

			editor::CEditor inspect(this, m_entryIDs ? IDS_ENTRYIDINSPECT : IDS_BINARYINSPECT, 0, CEDITOR_BUTTON_OK);
			inspect.AddPane(
				viewpane::TextPane::CreateMultiLinePane(0, m_entryIDs ? IDS_ENTRYIDENTRYID : IDS_BINARYVALUE, true));
			inspect.SetBinary(0, member.entryID.data(), member.entryID.size());
			if (m_entryIDs)
			{
				inspect.AddPane(viewpane::TextPane::CreateMultiLinePane(1, IDS_ENTRYIDDETAILS, true));
				inspect.SetStringW(1, details);
			}
			static_cast<void>(inspect.DisplayDialog());
		}
	}

	void CBinaryTableDlg::OnRefreshView()
	{
		if (!m_lpContentsTableListCtrl) return;
		if (m_lpContentsTableListCtrl->IsLoading()) m_lpContentsTableListCtrl->OnCancelTableLoad();
		if (!m_message)
		{
			m_lpContentsTableListCtrl->RefreshTable();
			return;
		}

		std::vector<mapi::binary::BinaryTableRow> members;
		LPMAPITABLE table = nullptr;
		const auto hRes = EC_H(LoadBinaryTable(m_message, m_lpMapiObjects, m_entryIDs, m_propTag, members, &table));
		if (SUCCEEDED(hRes))
		{
			m_members = std::move(members);
			m_lpContentsTableListCtrl->SetContentsTable(table, m_displayFlags, 0);
			if (m_lpContentsTable) m_lpContentsTable->Release();
			m_lpContentsTable = table;
			table = nullptr;
		}

		if (table) table->Release();
	}
} // namespace dialog
