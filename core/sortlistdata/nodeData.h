#pragma once
#include <core/sortlistdata/data.h>

namespace mapi
{
	class adviseSink;
} // namespace mapi

namespace sortlistdata
{
	class sortListData;

	class nodeData : public IData
	{
	public:
		static void init(
			sortListData* lpData,
			ULONG cProps,
			_In_opt_ LPSPropValue lpProps,
			_In_opt_ LPSBinary lpEntryID,
			_In_opt_ LPSBinary lpInstanceKey,
			ULONG bSubfolders,
			ULONG ulContainerFlags);
		static void init(sortListData* lpData, _In_ LPSRow lpsRow);
		~nodeData();

		LPSBinary m_lpEntryID{}; // Allocated with MAPIAllocateBuffer
		LPSBinary m_lpInstanceKey{}; // Allocated with MAPIAllocateBuffer
		LPMAPITABLE m_lpHierarchyTable{}; // Object - free with Release
		mapi::adviseSink* m_lpAdviseSink{}; // Object - free with Release
		ULONG_PTR m_ulAdviseConnection{};
		LONG m_cSubfolders{-1}; // -1 for unknown, 0 for no subfolders, >0 for at least one subfolder

	private:
		nodeData(
			_In_opt_ LPSBinary lpEntryID,
			_In_opt_ LPSBinary lpInstanceKey,
			ULONG bSubfolders,
			ULONG ulContainerFlags);
	};
} // namespace sortlistdata