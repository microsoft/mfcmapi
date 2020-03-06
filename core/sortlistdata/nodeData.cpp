#include <core/stdafx.h>
#include <core/sortlistdata/nodeData.h>
#include <core/mapi/mapiFunctions.h>
#include <core/sortlistdata/sortListData.h>
#include <core/utility/output.h>
#include <core/mapi/adviseSink.h>
#include <core/utility/error.h>
#include <core/utility/registry.h>

namespace sortlistdata
{
	void nodeData::init(
		sortListData* data,
		ULONG cProps,
		_In_opt_ LPSPropValue lpProps,
		_In_opt_ LPSBinary lpEntryID,
		_In_opt_ LPSBinary lpInstanceKey,
		ULONG bSubfolders,
		ULONG ulContainerFlags)
	{
		if (!data) return;

		data->init(
			std::make_shared<nodeData>(lpEntryID, lpInstanceKey, bSubfolders, ulContainerFlags), cProps, lpProps);
	}

	void nodeData::init(sortListData* data, _In_ LPSRow lpsRow)
	{
		if (!data) return;

		if (!lpsRow)
		{
			data->clean();
			return;
		}

		LPSBinary lpEIDBin = nullptr; // don't free
		LPSBinary lpInstanceBin = nullptr; // don't free

		auto lpEID = PpropFindProp(lpsRow->lpProps, lpsRow->cValues, PR_ENTRYID);
		if (lpEID) lpEIDBin = &lpEID->Value.bin;

		auto lpInstance = PpropFindProp(lpsRow->lpProps, lpsRow->cValues, PR_INSTANCE_KEY);
		if (lpInstance) lpInstanceBin = &lpInstance->Value.bin;

		const auto lpSubfolders = PpropFindProp(lpsRow->lpProps, lpsRow->cValues, PR_SUBFOLDERS);

		const auto lpContainerFlags = PpropFindProp(lpsRow->lpProps, lpsRow->cValues, PR_CONTAINER_FLAGS);

		sortlistdata::nodeData::init(
			data,
			lpsRow->cValues,
			lpsRow->lpProps, // pass on props to be archived in node
			lpEIDBin,
			lpInstanceBin,
			lpSubfolders ? static_cast<ULONG>(lpSubfolders->Value.b) : MAPI_E_NOT_FOUND,
			lpContainerFlags ? lpContainerFlags->Value.ul : MAPI_E_NOT_FOUND);
	}

	nodeData::nodeData(
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

	nodeData::~nodeData()
	{
		unadvise();
		MAPIFreeBuffer(m_lpInstanceKey);
		MAPIFreeBuffer(m_lpEntryID);

		if (m_lpHierarchyTable) m_lpHierarchyTable->Release();
	}

	bool nodeData::advise(HWND m_hWnd, HTREEITEM hItem, LPMDB lpMDB)
	{
		auto lpAdviseSink = new (std::nothrow) mapi::adviseSink(m_hWnd, hItem);

		if (lpAdviseSink)
		{
			const auto hRes = WC_MAPI(m_lpHierarchyTable->Advise(
				fnevTableModified, static_cast<IMAPIAdviseSink*>(lpAdviseSink), &m_ulAdviseConnection));
			if (hRes == MAPI_E_NO_SUPPORT) // Some tables don't support this!
			{
				lpAdviseSink->Release();
				lpAdviseSink = nullptr;
				output::DebugPrint(output::dbgLevel::Notify, L"This table doesn't support notifications\n");
				return false;
			}
			else if (hRes == S_OK && lpMDB)
			{
				lpAdviseSink->SetAdviseTarget(lpMDB);
				mapi::ForceRop(lpMDB);
			}

			output::DebugPrint(
				output::dbgLevel::Notify,
				L"nodeData::advise",
				L"Advise sink %p, ulAdviseConnection = 0x%08X\n",
				lpAdviseSink,
				static_cast<int>(m_ulAdviseConnection));

			m_lpAdviseSink = lpAdviseSink;
			return true;
		}

		return false;
	}

	void nodeData::unadvise()
	{
		if (m_lpAdviseSink)
		{
			output::DebugPrint(
				output::dbgLevel::Hierarchy,
				L"nodeData::unadvise",
				L"Unadvising %p, ulAdviseConnection = 0x%08X\n",
				m_lpAdviseSink,
				static_cast<int>(m_ulAdviseConnection));

			// unadvise before releasing our sink
			if (m_lpAdviseSink && m_lpHierarchyTable)
			{
				m_lpHierarchyTable->Unadvise(m_ulAdviseConnection);
			}

			if (m_lpAdviseSink)
			{
				m_lpAdviseSink->Release();
				m_lpAdviseSink = nullptr;
			}
		}
	}
} // namespace sortlistdata