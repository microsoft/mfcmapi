#include "stdafx.h"
#include "../stdafx.h"
#include "SortListData.h"
#include "ContentsData.h"
#include "NodeData.h"

SortListData* BuildNodeData(
	ULONG cProps,
	_In_opt_ LPSPropValue lpProps,
	_In_opt_ LPSBinary lpEntryID,
	_In_opt_ LPSBinary lpInstanceKey,
	ULONG bSubfolders,
	ULONG ulContainerFlags)
{
	auto lpData = new SortListData();
	if (lpData)
	{
		lpData->ulSortDataType = SORTLIST_TREENODE;

		lpData->cSourceProps = cProps;
		lpData->lpSourceProps = lpProps;

		lpData->m_lpData = new NodeData(
			lpEntryID,
			lpInstanceKey,
			bSubfolders,
			ulContainerFlags);

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
	}

	if (Contents() != nullptr) delete Contents();
	if (Node() != nullptr) delete Node();

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

NodeData* SortListData::Node() const
{
	if (ulSortDataType == SORTLIST_TREENODE)
	{
		return reinterpret_cast<NodeData*>(m_lpData);
	}

	return nullptr;
}