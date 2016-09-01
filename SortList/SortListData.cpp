#include "stdafx.h"
#include "../stdafx.h"
#include "SortListData.h"
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
	HRESULT hRes = S_OK;
	SortListData* lpData = new SortListData();
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

	LPSPropValue lpEID = PpropFindProp(
		lpsRow->lpProps,
		lpsRow->cValues,
		PR_ENTRYID);
	if (lpEID) lpEIDBin = &lpEID->Value.bin;

	LPSPropValue lpInstance = PpropFindProp(
		lpsRow->lpProps,
		lpsRow->cValues,
		PR_INSTANCE_KEY);
	if (lpInstance) lpInstanceBin = &lpInstance->Value.bin;

	LPSPropValue lpSubfolders = PpropFindProp(
		lpsRow->lpProps,
		lpsRow->cValues,
		PR_SUBFOLDERS);
	LPSPropValue lpContainerFlags = PpropFindProp(
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

ContentsData::ContentsData(_In_ LPSRow lpsRowData)
{
	lpEntryID = nullptr;
	lpLongtermID = nullptr;
	lpInstanceKey = nullptr;
	lpServiceUID = nullptr;
	lpProviderUID = nullptr;
	szDN = nullptr;
	szProfileDisplayName = nullptr;
	ulAttachNum = 0;
	ulAttachMethod = 0;
	ulRowID = 0;
	ulRowType = 0;

	if (!lpsRowData) return;

	HRESULT hRes = S_OK;

	// Save the instance key into lpData
	LPSPropValue lpProp = PpropFindProp(
		lpsRowData->lpProps,
		lpsRowData->cValues,
		PR_INSTANCE_KEY);
	if (lpProp && PR_INSTANCE_KEY == lpProp->ulPropTag)
	{
		EC_H(MAPIAllocateBuffer(
			static_cast<ULONG>(sizeof(SBinary)),
			reinterpret_cast<LPVOID*>(&lpInstanceKey)));
		EC_H(CopySBinary(lpInstanceKey, &lpProp->Value.bin, lpInstanceKey));
	}

	// Save the attachment number into lpData
	lpProp = PpropFindProp(
		lpsRowData->lpProps,
		lpsRowData->cValues,
		PR_ATTACH_NUM);
	if (lpProp && PR_ATTACH_NUM == lpProp->ulPropTag)
	{
		DebugPrint(DBGGeneric, L"\tPR_ATTACH_NUM = %d\n", lpProp->Value.l);
		ulAttachNum = lpProp->Value.l;
	}

	lpProp = PpropFindProp(
		lpsRowData->lpProps,
		lpsRowData->cValues,
		PR_ATTACH_METHOD);
	if (lpProp && PR_ATTACH_METHOD == lpProp->ulPropTag)
	{
		DebugPrint(DBGGeneric, L"\tPR_ATTACH_METHOD = %d\n", lpProp->Value.l);
		ulAttachMethod = lpProp->Value.l;
	}

	// Save the row ID (recipients) into lpData
	lpProp = PpropFindProp(
		lpsRowData->lpProps,
		lpsRowData->cValues,
		PR_ROWID);
	if (lpProp && PR_ROWID == lpProp->ulPropTag)
	{
		DebugPrint(DBGGeneric, L"\tPR_ROWID = %d\n", lpProp->Value.l);
		ulRowID = lpProp->Value.l;
	}

	// Save the row type (header/leaf) into lpData
	lpProp = PpropFindProp(
		lpsRowData->lpProps,
		lpsRowData->cValues,
		PR_ROW_TYPE);
	if (lpProp && PR_ROW_TYPE == lpProp->ulPropTag)
	{
		DebugPrint(DBGGeneric, L"\tPR_ROW_TYPE = %d\n", lpProp->Value.l);
		ulRowType = lpProp->Value.l;
	}

	// Save the Entry ID into lpData
	lpProp = PpropFindProp(
		lpsRowData->lpProps,
		lpsRowData->cValues,
		PR_ENTRYID);
	if (lpProp && PR_ENTRYID == lpProp->ulPropTag)
	{
		EC_H(MAPIAllocateBuffer(
			static_cast<ULONG>(sizeof(SBinary)),
			reinterpret_cast<LPVOID*>(& lpEntryID)));
		EC_H(CopySBinary(lpEntryID, &lpProp->Value.bin, lpEntryID));
	}

	// Save the Longterm Entry ID into lpData
	lpProp = PpropFindProp(
		lpsRowData->lpProps,
		lpsRowData->cValues,
		PR_LONGTERM_ENTRYID_FROM_TABLE);
	if (lpProp && PR_LONGTERM_ENTRYID_FROM_TABLE == lpProp->ulPropTag)
	{
		EC_H(MAPIAllocateBuffer(
			static_cast<ULONG>(sizeof(SBinary)),
			reinterpret_cast<LPVOID*>(& lpLongtermID)));
		EC_H(CopySBinary(lpLongtermID, &lpProp->Value.bin, lpLongtermID));
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
			static_cast<ULONG>(sizeof(SBinary)),
			reinterpret_cast<LPVOID*>(& lpServiceUID)));
		EC_H(CopySBinary(lpServiceUID, &lpProp->Value.bin, lpServiceUID));
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
			static_cast<ULONG>(sizeof(SBinary)),
			reinterpret_cast<LPVOID*>(& lpProviderUID)));
		EC_H(CopySBinary(lpProviderUID, &lpProp->Value.bin, lpProviderUID));
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
			&szProfileDisplayName,
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
			&szDN,
			lpProp->Value.LPSZ,
			nullptr));
	}
}

ContentsData::~ContentsData()
{
	MAPIFreeBuffer(lpInstanceKey);
	MAPIFreeBuffer(lpEntryID);
	MAPIFreeBuffer(lpLongtermID);
	MAPIFreeBuffer(lpServiceUID);
	MAPIFreeBuffer(lpProviderUID);
	MAPIFreeBuffer(szProfileDisplayName);
	MAPIFreeBuffer(szDN);
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
