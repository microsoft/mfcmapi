#include <StdAfx.h>
#include <UI/Controls/SortList/NodeData.h>
#include <UI/AdviseSink.h>
#include <core/mapi/mapiFunctions.h>

namespace controls
{
	namespace sortlistdata
	{
		NodeData::NodeData(
			_In_opt_ LPSBinary lpEntryID,
			_In_opt_ LPSBinary lpInstanceKey,
			ULONG bSubfolders,
			ULONG ulContainerFlags)
		{
			if (lpEntryID)
			{
				// Copy the data over
				m_lpEntryID = mapi::CopySBinary(lpEntryID);
			}

			if (lpInstanceKey)
			{
				m_lpInstanceKey = mapi::CopySBinary(lpInstanceKey);
			}

			if (bSubfolders != MAPI_E_NOT_FOUND)
			{
				m_cSubfolders = bSubfolders != 0;
			}
			else if (ulContainerFlags != MAPI_E_NOT_FOUND)
			{
				m_cSubfolders = ulContainerFlags & AB_SUBCONTAINERS ? 1 : 0;
			}
		}

		NodeData::~NodeData()
		{
			MAPIFreeBuffer(m_lpInstanceKey);
			MAPIFreeBuffer(m_lpEntryID);

			if (m_lpHierarchyTable) m_lpHierarchyTable->Release();
		}
	} // namespace sortlistdata
} // namespace controls