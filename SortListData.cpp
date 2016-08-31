#include "StdAfx.h"
#include "SortListData.h"
#include "MAPIFunctions.h"
#include "AdviseSink.h"

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
	SortListData* lpData = new SortListData();
	if (lpData)
	{
		memset(lpData, 0, sizeof(SortListData));
		lpData->ulSortDataType = SORTLIST_TREENODE;

		if (lpEntryID)
		{
			WC_H(MAPIAllocateBuffer(
				(ULONG)sizeof(SBinary),
				(LPVOID*)&lpData->data.Node.lpEntryID));

			// Copy the data over
			WC_H(CopySBinary(
				lpData->data.Node.lpEntryID,
				lpEntryID,
				lpData->data.Node.lpEntryID));
		}

		if (lpInstanceKey)
		{
			WC_H(MAPIAllocateBuffer(
				(ULONG)sizeof(SBinary),
				(LPVOID*)&lpData->data.Node.lpInstanceKey));
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

// Sets data from the LPSRow into the SortListData structure
// Assumes the structure is either an existing structure or a new one which has been memset to 0
// If it's an existing structure - we need to free up some memory
// SORTLIST_CONTENTS
void BuildDataItem(_In_ LPSRow lpsRowData, _Inout_ SortListData* lpData)
{
	if (!lpData || !lpsRowData) return;

	HRESULT hRes = S_OK;
	LPSPropValue lpProp = NULL; // do not free this

	lpData->bItemFullyLoaded = false;
	lpData->szSortText = emptystring;

	// this guy gets stolen from lpsRowData and is freed separately in FreeSortListData
	// So I do need to free it here before losing the pointer
	MAPIFreeBuffer(lpData->lpSourceProps);
	lpData->lpSourceProps = NULL;

	lpData->ulSortValue.QuadPart = NULL;
	lpData->cSourceProps = 0;
	lpData->ulSortDataType = SORTLIST_CONTENTS;
	memset(&lpData->data, 0, sizeof(_ContentsData));

	// Save off the source props
	lpData->lpSourceProps = lpsRowData->lpProps;
	lpData->cSourceProps = lpsRowData->cValues;

	// Save the instance key into lpData
	lpProp = PpropFindProp(
		lpsRowData->lpProps,
		lpsRowData->cValues,
		PR_INSTANCE_KEY);
	if (lpProp && PR_INSTANCE_KEY == lpProp->ulPropTag)
	{
		EC_H(MAPIAllocateBuffer(
			(ULONG)sizeof(SBinary),
			(LPVOID*)&lpData->data.Contents.lpInstanceKey));
		EC_H(CopySBinary(lpData->data.Contents.lpInstanceKey, &lpProp->Value.bin, lpData->data.Contents.lpInstanceKey));
	}

	// Save the attachment number into lpData
	lpProp = PpropFindProp(
		lpsRowData->lpProps,
		lpsRowData->cValues,
		PR_ATTACH_NUM);
	if (lpProp && PR_ATTACH_NUM == lpProp->ulPropTag)
	{
		DebugPrint(DBGGeneric, L"\tPR_ATTACH_NUM = %d\n", lpProp->Value.l);
		lpData->data.Contents.ulAttachNum = lpProp->Value.l;
	}

	lpProp = PpropFindProp(
		lpsRowData->lpProps,
		lpsRowData->cValues,
		PR_ATTACH_METHOD);
	if (lpProp && PR_ATTACH_METHOD == lpProp->ulPropTag)
	{
		DebugPrint(DBGGeneric, L"\tPR_ATTACH_METHOD = %d\n", lpProp->Value.l);
		lpData->data.Contents.ulAttachMethod = lpProp->Value.l;
	}

	// Save the row ID (recipients) into lpData
	lpProp = PpropFindProp(
		lpsRowData->lpProps,
		lpsRowData->cValues,
		PR_ROWID);
	if (lpProp && PR_ROWID == lpProp->ulPropTag)
	{
		DebugPrint(DBGGeneric, L"\tPR_ROWID = %d\n", lpProp->Value.l);
		lpData->data.Contents.ulRowID = lpProp->Value.l;
	}

	// Save the row type (header/leaf) into lpData
	lpProp = PpropFindProp(
		lpsRowData->lpProps,
		lpsRowData->cValues,
		PR_ROW_TYPE);
	if (lpProp && PR_ROW_TYPE == lpProp->ulPropTag)
	{
		DebugPrint(DBGGeneric, L"\tPR_ROW_TYPE = %d\n", lpProp->Value.l);
		lpData->data.Contents.ulRowType = lpProp->Value.l;
	}

	// Save the Entry ID into lpData
	lpProp = PpropFindProp(
		lpsRowData->lpProps,
		lpsRowData->cValues,
		PR_ENTRYID);
	if (lpProp && PR_ENTRYID == lpProp->ulPropTag)
	{
		EC_H(MAPIAllocateBuffer(
			(ULONG)sizeof(SBinary),
			(LPVOID*)&lpData->data.Contents.lpEntryID));
		EC_H(CopySBinary(lpData->data.Contents.lpEntryID, &lpProp->Value.bin, lpData->data.Contents.lpEntryID));
	}

	// Save the Longterm Entry ID into lpData
	lpProp = PpropFindProp(
		lpsRowData->lpProps,
		lpsRowData->cValues,
		PR_LONGTERM_ENTRYID_FROM_TABLE);
	if (lpProp && PR_LONGTERM_ENTRYID_FROM_TABLE == lpProp->ulPropTag)
	{
		EC_H(MAPIAllocateBuffer(
			(ULONG)sizeof(SBinary),
			(LPVOID*)&lpData->data.Contents.lpLongtermID));
		EC_H(CopySBinary(lpData->data.Contents.lpLongtermID, &lpProp->Value.bin, lpData->data.Contents.lpLongtermID));
	}

	// Save the Service ID into lpData
	lpProp = PpropFindProp(
		lpsRowData->lpProps,
		lpsRowData->cValues,
		PR_SERVICE_UID);
	if (lpProp && PR_SERVICE_UID == lpProp->ulPropTag)
	{
		// Allocate some space
		EC_H(MAPIAllocateBuffer(
			(ULONG)sizeof(SBinary),
			(LPVOID*)&lpData->data.Contents.lpServiceUID));
		EC_H(CopySBinary(lpData->data.Contents.lpServiceUID, &lpProp->Value.bin, lpData->data.Contents.lpServiceUID));
	}

	// Save the Provider ID into lpData
	lpProp = PpropFindProp(
		lpsRowData->lpProps,
		lpsRowData->cValues,
		PR_PROVIDER_UID);
	if (lpProp && PR_PROVIDER_UID == lpProp->ulPropTag)
	{
		// Allocate some space
		EC_H(MAPIAllocateBuffer(
			(ULONG)sizeof(SBinary),
			(LPVOID*)&lpData->data.Contents.lpProviderUID));
		EC_H(CopySBinary(lpData->data.Contents.lpProviderUID, &lpProp->Value.bin, lpData->data.Contents.lpProviderUID));
	}

	// Save the DisplayName into lpData
	lpProp = PpropFindProp(
		lpsRowData->lpProps,
		lpsRowData->cValues,
		PR_DISPLAY_NAME_A); // We pull this properties for profiles, which do not support Unicode
	if (CheckStringProp(lpProp, PT_STRING8))
	{
		DebugPrint(DBGGeneric, L"\tPR_DISPLAY_NAME_A = %hs\n", lpProp->Value.lpszA);

		EC_H(CopyStringA(
			&lpData->data.Contents.szProfileDisplayName,
			lpProp->Value.lpszA,
			nullptr));
	}

	// Save the e-mail address (if it exists on the object) into lpData
	lpProp = PpropFindProp(
		lpsRowData->lpProps,
		lpsRowData->cValues,
		PR_EMAIL_ADDRESS);
	if (CheckStringProp(lpProp, PT_TSTRING))
	{
		DebugPrint(DBGGeneric, L"\tPR_EMAIL_ADDRESS = %ws\n", LPCTSTRToWstring(lpProp->Value.LPSZ).c_str());
		EC_H(CopyString(
			&lpData->data.Contents.szDN,
			lpProp->Value.LPSZ,
			nullptr));
	}
}

void FreeSortListData(_In_ SortListData* lpData)
{
	if (!lpData) return;
	switch (lpData->ulSortDataType)
	{
	case SORTLIST_CONTENTS: // _ContentsData
		MAPIFreeBuffer(lpData->data.Contents.lpInstanceKey);
		MAPIFreeBuffer(lpData->data.Contents.lpEntryID);
		MAPIFreeBuffer(lpData->data.Contents.lpLongtermID);
		MAPIFreeBuffer(lpData->data.Contents.lpServiceUID);
		MAPIFreeBuffer(lpData->data.Contents.lpProviderUID);
		MAPIFreeBuffer(lpData->data.Contents.szProfileDisplayName);
		MAPIFreeBuffer(lpData->data.Contents.szDN);
		break;
	case SORTLIST_PROP: // _PropListData
	case SORTLIST_MVPROP: // _MVPropData
	case SORTLIST_TAGARRAY: // _TagData
	case SORTLIST_BINARY: // _BinaryData
	case SORTLIST_RES: // _ResData
	case SORTLIST_COMMENT: // _CommentData
		break;
	case SORTLIST_TREENODE: // _NodeData
		MAPIFreeBuffer(lpData->data.Contents.lpInstanceKey);
		MAPIFreeBuffer(lpData->data.Contents.lpEntryID);

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
	delete lpData;
}