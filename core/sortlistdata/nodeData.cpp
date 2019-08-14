#include <core/stdafx.h>
#include <core/sortlistdata/nodeData.h>
#include <core/mapi/mapiFunctions.h>
#include <core/sortlistdata/sortListData.h>

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
		data->Init(
			new (std::nothrow) nodeData(lpEntryID, lpInstanceKey, bSubfolders, ulContainerFlags), cProps, lpProps);
	}

	void nodeData::init(sortListData* data, _In_ LPSRow lpsRow)
	{
		if (!data) return;

		if (!lpsRow)
		{
			data->Clean();
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
		MAPIFreeBuffer(m_lpInstanceKey);
		MAPIFreeBuffer(m_lpEntryID);

		if (m_lpHierarchyTable) m_lpHierarchyTable->Release();
	}
} // namespace sortlistdata