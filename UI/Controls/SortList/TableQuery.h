#pragma once
#include <core/res/Resource.h>
#include <core/utility/error.h>

namespace controls::sortlistctrl
{
	_Check_return_ inline HRESULT
	QueryRowsWithDiagnostics(_In_ LPMAPITABLE table, LONG rowCount, ULONG flags, _Deref_out_ LPSRowSet* rows)
	{
		if (!table || !rows) return EC_H(MAPI_E_INVALID_PARAMETER);

		const auto hRes = table->QueryRows(rowCount, flags, rows);
		UINT message = 0;
		if (hRes == MAPI_E_NO_SUPPORT)
		{
			LPSPropTagArray columns = nullptr;
			const auto columnsResult = WC_MAPI(table->QueryColumns(0, &columns));
			if (SUCCEEDED(columnsResult) && columns)
			{
				for (ULONG i = 0; i < columns->cValues; ++i)
				{
					if (PROP_TYPE(columns->aulPropTag[i]) & MV_INSTANCE)
					{
						message = IDS_QUERYROWSMVINSTANCE;
						break;
					}
				}
			}

			MAPIFreeBuffer(columns);
		}

		error::LogFunctionCall(
			hRes, {}, true, true, false, message, "table->QueryRows(rowCount, flags, rows)", __FILE__, __LINE__);
		return hRes;
	}
} // namespace controls::sortlistctrl
