#pragma once
#include <UI/enums.h>
#include <core/sortlistdata/sortListData.h>

namespace controls::sortlistctrl
{
	// Inserts a list row and returns its item data, or nullptr if the insert failed.
	_Check_return_ sortlistdata::sortListData*
	InsertSortListRow(HWND window, int row, const std::wstring& text, int indent, int image);

	enum class TableLoadResult
	{
		complete,
		failed,
		cancelled,
		limited
	};

	// https://learn.microsoft.com/en-us/office/client-developer/outlook/mapi/imapitable-queryrows
	// A successful add callback takes ownership of the row's property buffer.
	_Check_return_ TableLoadResult LoadTableRows(
		LPMAPITABLE table,
		restrictionType restriction,
		const SRestriction* filter,
		ULONG throttle,
		ULONG rowCount,
		const std::function<bool()>& cancelled,
		const std::function<void()>& beforeBatch,
		const std::function<HRESULT(ULONG, SRow&)>& add,
		_Out_ ULONG& added);
} // namespace controls::sortlistctrl
