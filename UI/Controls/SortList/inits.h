#pragma once
//#include <core/sortlistdata/data.h>
#include <UI/Controls/SortList/SortListData.h>

namespace controls
{
	namespace sortlistdata
	{
		void InitContents(SortListData* lpData, _In_ LPSRow lpsRowData);
		void InitNode(
			SortListData* lpData,
			ULONG cProps,
			_In_opt_ LPSPropValue lpProps,
			_In_opt_ LPSBinary lpEntryID,
			_In_opt_ LPSBinary lpInstanceKey,
			ULONG bSubfolders,
			ULONG ulContainerFlags);
		void InitNode(SortListData* lpData, _In_ LPSRow lpsRow);
		void InitPropList(SortListData* data, _In_ ULONG ulPropTag);
		void InitMV(SortListData* data, _In_ const _SPropValue* lpProp, ULONG iProp);
		void InitMV(SortListData* data, _In_opt_ const _SPropValue* lpProp);
		void InitRes(SortListData* data, _In_opt_ const _SRestriction* lpOldRes);
		void InitComment(SortListData* data, _In_opt_ const _SPropValue* lpOldProp);
		void InitBinary(SortListData* data, _In_opt_ LPSBinary lpOldBin);
	} // namespace sortlistdata
} // namespace controls