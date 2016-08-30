#include "StdAfx.h"
#include "SortListData.h"
#include "MAPIFunctions.h"
#include "AdviseSink.h"

void FreeSortListData(_In_ SortListData* lpData)
{
	if (!lpData) return;
	switch (lpData->ulSortDataType)
	{
	case SORTLIST_CONTENTS: // _ContentsData
	case SORTLIST_PROP: // _PropListData
	case SORTLIST_MVPROP: // _MVPropData
	case SORTLIST_TAGARRAY: // _TagData
	case SORTLIST_BINARY: // _BinaryData
	case SORTLIST_RES: // _ResData
	case SORTLIST_COMMENT: // _CommentData
		break;
	case SORTLIST_TREENODE: // _NodeData
		if (lpData->data.Node.lpAdviseSink)
		{
			// unadvise before releasing our sink
			if (lpData->data.Node.ulAdviseConnection && lpData->data.Node.lpHierarchyTable)
				lpData->data.Node.lpHierarchyTable->Unadvise(lpData->data.Node.ulAdviseConnection);
			lpData->data.Node.lpAdviseSink->Release();
		}

		if (lpData->data.Node.lpHierarchyTable) lpData->data.Node.lpHierarchyTable->Release();
		break;
	}

	MAPIFreeBuffer(lpData->lpSourceProps);
	MAPIFreeBuffer(lpData);
}

// TODO: Move code to copy props here and normalize calling code so it doesn't appear to leak
SortListData* BuildNodeData(
	ULONG cProps,
	_In_opt_ LPSPropValue lpProps,
	_In_opt_ LPSBinary lpEntryID,
	_In_opt_ LPSBinary lpInstanceKey,
	ULONG bSubfolders,
	ULONG ulContainerFlags)
{
	HRESULT hRes = S_OK;
	SortListData* lpData = NULL;
	// We're gonna set up a data item to pass off to the tree
	// Allocate some space
	WC_H(MAPIAllocateBuffer(
		(ULONG)sizeof(SortListData),
		(LPVOID *)&lpData));

	if (lpData)
	{
		memset(lpData, 0, sizeof(SortListData));
		lpData->ulSortDataType = SORTLIST_TREENODE;

		if (lpEntryID)
		{
			WC_H(MAPIAllocateMore(
				(ULONG)sizeof(SBinary),
				lpData,
				(LPVOID*)&lpData->data.Node.lpEntryID));

			// Copy the data over
			WC_H(CopySBinary(
				lpData->data.Node.lpEntryID,
				lpEntryID,
				lpData));
		}

		if (lpInstanceKey)
		{
			WC_H(MAPIAllocateMore(
				(ULONG)sizeof(SBinary),
				lpData,
				(LPVOID*)&lpData->data.Node.lpInstanceKey));
			WC_H(CopySBinary(
				lpData->data.Node.lpInstanceKey,
				lpInstanceKey,
				lpData));
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
	return NULL;
}

SortListData* BuildNodeData(_In_ LPSRow lpsRow)
{
	if (!lpsRow) return NULL;

	LPSPropValue lpEID = NULL; // don't free
	LPSPropValue lpInstance = NULL; // don't free
	LPSBinary lpEIDBin = NULL; // don't free
	LPSBinary lpInstanceBin = NULL; // don't free
	LPSPropValue lpSubfolders = NULL; // don't free
	LPSPropValue lpContainerFlags = NULL; // don't free

	lpEID = PpropFindProp(
		lpsRow->lpProps,
		lpsRow->cValues,
		PR_ENTRYID);
	if (lpEID) lpEIDBin = &lpEID->Value.bin;
	lpInstance = PpropFindProp(
		lpsRow->lpProps,
		lpsRow->cValues,
		PR_INSTANCE_KEY);
	if (lpInstance) lpInstanceBin = &lpInstance->Value.bin;

	lpSubfolders = PpropFindProp(
		lpsRow->lpProps,
		lpsRow->cValues,
		PR_SUBFOLDERS);
	lpContainerFlags = PpropFindProp(
		lpsRow->lpProps,
		lpsRow->cValues,
		PR_CONTAINER_FLAGS);

	return BuildNodeData(
		lpsRow->cValues,
		lpsRow->lpProps, // pass on props to be archived in node
		lpEIDBin,
		lpInstanceBin,
		lpSubfolders ? (ULONG)lpSubfolders->Value.b : MAPI_E_NOT_FOUND,
		lpContainerFlags ? lpContainerFlags->Value.ul : MAPI_E_NOT_FOUND);
}