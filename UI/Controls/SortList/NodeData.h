#pragma once
#include <UI/Controls/SortList/Data.h>
#include <UI/Controls/SortList/SortListData.h>

namespace mapi
{
	namespace mapiui
	{
		class CAdviseSink;
	}
} // namespace mapi

namespace controls
{
	namespace sortlistdata
	{
		class NodeData : public IData
		{
		public:
			NodeData(
				_In_opt_ LPSBinary lpEntryID,
				_In_opt_ LPSBinary lpInstanceKey,
				ULONG bSubfolders,
				ULONG ulContainerFlags);
			~NodeData();

			LPSBinary m_lpEntryID; // Allocated with MAPIAllocateBuffer
			LPSBinary m_lpInstanceKey; // Allocated with MAPIAllocateBuffer
			LPMAPITABLE m_lpHierarchyTable; // Object - free with Release
			mapi::mapiui::CAdviseSink* m_lpAdviseSink; // Object - free with Release
			ULONG_PTR m_ulAdviseConnection;
			LONG m_cSubfolders; // -1 for unknown, 0 for no subfolders, >0 for at least one subfolder
		};
	} // namespace sortlistdata
} // namespace controls