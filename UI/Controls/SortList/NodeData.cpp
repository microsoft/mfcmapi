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
			: m_lpEntryID(nullptr), m_lpInstanceKey(nullptr), m_lpHierarchyTable(nullptr), m_lpAdviseSink(nullptr),
			  m_ulAdviseConnection(0)
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

			m_cSubfolders = -1;
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

			if (m_lpAdviseSink)
			{
				// unadvise before releasing our sink
				if (m_ulAdviseConnection && m_lpHierarchyTable) m_lpHierarchyTable->Unadvise(m_ulAdviseConnection);
				m_lpAdviseSink->Release();
			}

			if (m_lpHierarchyTable) m_lpHierarchyTable->Release();
		}
	} // namespace sortlistdata
} // namespace controls