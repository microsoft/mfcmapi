#include <StdAfx.h>
#include <UI/Controls/SortList/TableLoad.h>
#include <core/mapi/mapiFunctions.h>
#include <core/mapi/mapiOutput.h>
#include <core/utility/error.h>
#include <core/utility/output.h>

namespace controls::sortlistctrl
{
	_Check_return_ sortlistdata::sortListData*
	InsertSortListRow(HWND window, int row, const std::wstring& text, int indent, int image)
	{
		auto data = new (std::nothrow) sortlistdata::sortListData();
		if (!data)
		{
			WC_H_S(E_OUTOFMEMORY);
			return nullptr;
		}

		LVITEMW item = {};
		item.iItem = row;
		item.iSubItem = 0;
		item.mask = LVIF_TEXT | LVIF_PARAM | LVIF_INDENT | LVIF_IMAGE;
		item.pszText = const_cast<LPWSTR>(text.c_str());
		item.iIndent = indent;
		item.iImage = image;
		item.lParam = reinterpret_cast<LPARAM>(data);
		if (::SendMessage(window, LVM_INSERTITEMW, 0, reinterpret_cast<LPARAM>(&item)) == -1)
		{
			delete data;
			WC_H_S(E_FAIL);
			return nullptr;
		}

		return data;
	}

	_Check_return_ TableLoadResult LoadTableRows(
		LPMAPITABLE table,
		restrictionType restriction,
		const SRestriction* filter,
		ULONG throttle,
		ULONG rowCount,
		const std::function<bool()>& cancelled,
		const std::function<void()>& beforeBatch,
		const std::function<HRESULT(ULONG, SRow&)>& add,
		_Out_ ULONG& added)
	{
		added = 0;
		if (!table)
		{
			WC_H_S(MAPI_E_INVALID_PARAMETER);
			return TableLoadResult::failed;
		}
		for (;;)
		{
			if (cancelled()) return TableLoadResult::cancelled;
			beforeBatch();
			LPSRowSet rows = nullptr;
			auto hRes = S_OK;
			if (restriction == restrictionType::findrow && filter)
			{
				output::DebugPrintEx(
					output::dbgLevel::Generic,
					L"CContentsTableListCtrl",
					L"DoFindRows",
					L"running FindRow with restriction:\n");
				output::outputRestriction(output::dbgLevel::Generic, nullptr, filter, nullptr);
				if (cancelled()) return TableLoadResult::cancelled;
				hRes = WC_MAPI(table->FindRow(const_cast<LPSRestriction>(filter), BOOKMARK_CURRENT, 0));
				if (FAILED(hRes)) return hRes == MAPI_E_NOT_FOUND ? TableLoadResult::complete : TableLoadResult::failed;

				if (cancelled()) return TableLoadResult::cancelled;
				hRes = mapi::QueryRowsWithDiagnostics(table, 1, 0, &rows);
			}
			else
			{
				output::DebugPrintEx(
					output::dbgLevel::Generic,
					L"CContentsTableListCtrl",
					L"ThreadFuncLoadTable",
					L"Calling QueryRows. Asking for 0x%X rows.\n",
					rowCount);
				if (cancelled()) return TableLoadResult::cancelled;
				hRes = mapi::QueryRowsWithDiagnostics(table, rowCount, 0, &rows);
			}

			if (FAILED(hRes))
			{
				if (rows) FreeProws(rows);
				return TableLoadResult::failed;
			}

			if (!rows || !rows->cRows)
			{
				if (rows) FreeProws(rows);
				return TableLoadResult::complete;
			}

			output::DebugPrintEx(
				output::dbgLevel::Generic,
				L"CContentsTableListCtrl",
				L"ThreadFuncLoadTable",
				L"Got this many rows: 0x%X\n",
				rows->cRows);
			auto result = TableLoadResult::complete;
			for (ULONG i = 0; i < rows->cRows; ++i)
			{
				if (cancelled())
				{
					result = TableLoadResult::cancelled;
					break;
				}

				hRes = add(added, rows->aRow[i]);
				if (FAILED(hRes))
				{
					result = TableLoadResult::failed;
					break;
				}

				rows->aRow[i].cValues = 0;
				rows->aRow[i].lpProps = nullptr;
				++added;
			}

			FreeProws(rows);
			if (result != TableLoadResult::complete) return result;
			if (throttle && added >= throttle)
				return cancelled() ? TableLoadResult::cancelled : TableLoadResult::limited;
		}
	}
} // namespace controls::sortlistctrl
