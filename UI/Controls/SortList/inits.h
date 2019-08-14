#pragma once
#include <core/sortlistdata/sortListData.h>

namespace sortlistdata
{
	void InitContents(sortListData* lpData, _In_ LPSRow lpsRowData);
	void InitNode(
		sortListData* lpData,
		ULONG cProps,
		_In_opt_ LPSPropValue lpProps,
		_In_opt_ LPSBinary lpEntryID,
		_In_opt_ LPSBinary lpInstanceKey,
		ULONG bSubfolders,
		ULONG ulContainerFlags);
	void InitNode(sortListData* lpData, _In_ LPSRow lpsRow);
	void InitPropList(sortListData* data, _In_ ULONG ulPropTag);
	void InitMV(sortListData* data, _In_ const _SPropValue* lpProp, ULONG iProp);
	void InitMV(sortListData* data, _In_opt_ const _SPropValue* lpProp);
	void InitRes(sortListData* data, _In_opt_ const _SRestriction* lpOldRes);
	void InitComment(sortListData* data, _In_opt_ const _SPropValue* lpOldProp);
	void InitBinary(sortListData* data, _In_opt_ LPSBinary lpOldBin);
} // namespace sortlistdata