#include "stdafx.h"
#include "NodeData.h"
#include "AdviseSink.h"
#include "MAPIFunctions.h"
#include "SortListData.h"

// TODO: Move code to copy props here and normalize calling code so it doesn't appear to leak
NodeData::NodeData(
	_In_opt_ LPSBinary lpEntryID,
	_In_opt_ LPSBinary lpInstanceKey,
	ULONG bSubfolders,
	ULONG ulContainerFlags) : lpEntryID(nullptr), lpInstanceKey(nullptr), lpHierarchyTable(nullptr), lpAdviseSink(nullptr), ulAdviseConnection(0)
{
	auto hRes = S_OK;
	if (lpEntryID)
	{
		WC_H(MAPIAllocateBuffer(
			static_cast<ULONG>(sizeof(SBinary)),
			reinterpret_cast<LPVOID*>(&lpEntryID)));

		// Copy the data over
		WC_H(CopySBinary(
			lpEntryID,
			lpEntryID,
			lpEntryID));
	}

	if (lpInstanceKey)
	{
		WC_H(MAPIAllocateBuffer(
			static_cast<ULONG>(sizeof(SBinary)),
			reinterpret_cast<LPVOID*>(&lpInstanceKey)));
		WC_H(CopySBinary(
			lpInstanceKey,
			lpInstanceKey,
			lpInstanceKey));
	}

	cSubfolders = -1;
	if (bSubfolders != MAPI_E_NOT_FOUND)
	{
		cSubfolders = (bSubfolders != 0);
	}
	else if (ulContainerFlags != MAPI_E_NOT_FOUND)
	{
		cSubfolders = (ulContainerFlags & AB_SUBCONTAINERS) ? 1 : 0;
	}
}

NodeData::~NodeData()
{
	MAPIFreeBuffer(lpInstanceKey);
	MAPIFreeBuffer(lpEntryID);

	if (lpAdviseSink)
	{
		// unadvise before releasing our sink
		if (ulAdviseConnection &&  lpHierarchyTable)
			lpHierarchyTable->Unadvise(ulAdviseConnection);
		lpAdviseSink->Release();
	}

	if (lpHierarchyTable)  lpHierarchyTable->Release();
}