#include "stdafx.h"
#include "../stdafx.h"
#include "SortListData.h"
#include "ContentsData.h"
#include "../MAPIFunctions.h"
#include "../AdviseSink.h"

// TODO: Move code to copy props here and normalize calling code so it doesn't appear to leak
SortListData* BuildNodeData(
	ULONG cProps,
	_In_opt_ LPSPropValue lpProps,
	_In_opt_ LPSBinary lpEntryID,
	_In_opt_ LPSBinary lpInstanceKey,
	ULONG bSubfolders,
	ULONG ulContainerFlags)
{
	auto hRes = S_OK;
	auto lpData = new SortListData();
	if (lpData)
	{
		memset(lpData, 0, sizeof(SortListData));
		lpData->ulSortDataType = SORTLIST_TREENODE;

		if (lpEntryID)
		{
			WC_H(MAPIAllocateBuffer(
				static_cast<ULONG>(sizeof(SBinary)),
				reinterpret_cast<LPVOID*>(&lpData->data.Node.lpEntryID)));

			// Copy the data over
			WC_H(CopySBinary(
				lpData->data.Node.lpEntryID,
				lpEntryID,
				lpData->data.Node.lpEntryID));
		}

		if (lpInstanceKey)
		{
			WC_H(MAPIAllocateBuffer(
				static_cast<ULONG>(sizeof(SBinary)),
				reinterpret_cast<LPVOID*>(&lpData->data.Node.lpInstanceKey)));
			WC_H(CopySBinary(
				lpData->data.Node.lpInstanceKey,
				lpInstanceKey,
				lpData->data.Node.lpInstanceKey));
		}

		lpData->data.Node.cSubfolders = -1;
		if (bSubfolders != MAPI_E_NOT_FOUND)
		{
			lpData->data.Node.cSubfolders = (bSubfolders != 0);
		}
		else if (ulContainerFlags != MAPI_E_NOT_FOUND)
		{
			lpData->data.Node.cSubfolders = (ulContainerFlags & AB_SUBCONTAINERS) ? 1 : 0;
		}

		lpData->cSourceProps = cProps;
		lpData->lpSourceProps = lpProps;

		return lpData;
	}

	return nullptr;
}

SortListData* BuildNodeData(_In_ LPSRow lpsRow)
{
	if (!lpsRow) return nullptr;

	LPSBinary lpEIDBin = nullptr; // don't free
	LPSBinary lpInstanceBin = nullptr; // don't free

	auto lpEID = PpropFindProp(
		lpsRow->lpProps,
		lpsRow->cValues,
		PR_ENTRYID);
	if (lpEID) lpEIDBin = &lpEID->Value.bin;

	auto lpInstance = PpropFindProp(
		lpsRow->lpProps,
		lpsRow->cValues,
		PR_INSTANCE_KEY);
	if (lpInstance) lpInstanceBin = &lpInstance->Value.bin;

	auto lpSubfolders = PpropFindProp(
		lpsRow->lpProps,
		lpsRow->cValues,
		PR_SUBFOLDERS);

	auto lpContainerFlags = PpropFindProp(
		lpsRow->lpProps,
		lpsRow->cValues,
		PR_CONTAINER_FLAGS);

	return BuildNodeData(
		lpsRow->cValues,
		lpsRow->lpProps, // pass on props to be archived in node
		lpEIDBin,
		lpInstanceBin,
		lpSubfolders ? static_cast<ULONG>(lpSubfolders->Value.b) : MAPI_E_NOT_FOUND,
		lpContainerFlags ? lpContainerFlags->Value.ul : MAPI_E_NOT_FOUND);
}

// Sets data from the LPSRow into the SortListData structure
// Assumes the structure is either an existing structure or a new one which has been memset to 0
// If it's an existing structure - we need to free up some memory
// SORTLIST_CONTENTS
void BuildDataItem(_In_ LPSRow lpsRowData, _Inout_ SortListData* lpData)
{
	if (!lpData || !lpsRowData) return;

	lpData->bItemFullyLoaded = false;
	lpData->szSortText = emptystring;

	// this guy gets stolen from lpsRowData and is freed separately in FreeSortListData
	// So I do need to free it here before losing the pointer
	MAPIFreeBuffer(lpData->lpSourceProps);
	lpData->lpSourceProps = nullptr;

	lpData->ulSortValue.QuadPart = NULL;
	lpData->cSourceProps = 0;
	lpData->ulSortDataType = SORTLIST_CONTENTS;
	lpData->m_lpData = new ContentsData(lpsRowData);

	// Save off the source props
	lpData->lpSourceProps = lpsRowData->lpProps;
	lpData->cSourceProps = lpsRowData->cValues;
}

SortListData::~SortListData()
{
	switch (ulSortDataType)
	{
	case SORTLIST_PROP: // _PropListData
	case SORTLIST_MVPROP: // _MVPropData
	case SORTLIST_TAGARRAY: // _TagData
	case SORTLIST_BINARY: // _BinaryData
	case SORTLIST_RES: // _ResData
	case SORTLIST_COMMENT: // _CommentData
		break;
	case SORTLIST_TREENODE: // _NodeData
		MAPIFreeBuffer(data.Node.lpInstanceKey);
		MAPIFreeBuffer(data.Node.lpEntryID);

		if (data.Node.lpAdviseSink)
		{
			// unadvise before releasing our sink
			if (data.Node.ulAdviseConnection && data.Node.lpHierarchyTable)
				data.Node.lpHierarchyTable->Unadvise(data.Node.ulAdviseConnection);
			data.Node.lpAdviseSink->Release();
		}

		if (data.Node.lpHierarchyTable) data.Node.lpHierarchyTable->Release();
		break;
	}

	if (Contents() != nullptr) delete Contents();

	MAPIFreeBuffer(lpSourceProps);
}

ContentsData* SortListData::Contents() const
{
	if (ulSortDataType == SORTLIST_CONTENTS)
	{
		return reinterpret_cast<ContentsData*>(m_lpData);
	}

	return nullptr;
}
