#include <StdAfx.h>
#include <UI/Controls/SortList/NodeData.h>
#include <MAPI/AdviseSink.h>
#include <MAPI/MAPIFunctions.h>
#include <MAPI/MapiMemory.h>

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
			auto hRes = S_OK;
			if (lpEntryID)
			{
				m_lpEntryID = mapi::allocate<LPSBinary>(static_cast<ULONG>(sizeof(SBinary)));
				if (m_lpEntryID)
				{
					// Copy the data over
					hRes = WC_H(mapi::CopySBinary(m_lpEntryID, lpEntryID, nullptr));
				}
			}

			if (lpInstanceKey)
			{
				m_lpInstanceKey = mapi::allocate<LPSBinary>(static_cast<ULONG>(sizeof(SBinary)));
				if (m_lpInstanceKey)
				{
					hRes = WC_H(mapi::CopySBinary(m_lpInstanceKey, lpInstanceKey, nullptr));
				}
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
	}
}