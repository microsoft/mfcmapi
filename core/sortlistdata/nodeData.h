#pragma once
#include <core/sortlistdata/data.h>
#include <core/sortlistdata/sortListData.h>

namespace mapi
{
	class adviseSink;
} // namespace mapi

namespace sortlistdata
{
	class nodeData : public IData
	{
	public:
		nodeData(
			_In_opt_ LPSBinary lpEntryID,
			_In_opt_ LPSBinary lpInstanceKey,
			ULONG bSubfolders,
			ULONG ulContainerFlags);
		~nodeData();

		LPSBinary m_lpEntryID{}; // Allocated with MAPIAllocateBuffer
		LPSBinary m_lpInstanceKey{}; // Allocated with MAPIAllocateBuffer
		LPMAPITABLE m_lpHierarchyTable{}; // Object - free with Release
		mapi::adviseSink* m_lpAdviseSink{}; // Object - free with Release
		ULONG_PTR m_ulAdviseConnection{};
		LONG m_cSubfolders{-1}; // -1 for unknown, 0 for no subfolders, >0 for at least one subfolder
	};
} // namespace sortlistdata