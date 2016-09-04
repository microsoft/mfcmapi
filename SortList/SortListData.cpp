#include "stdafx.h"
#include "SortListData.h"
#include "ContentsData.h"
#include "NodeData.h"

SortListData::SortListData() :
	ulSortDataType(SORTLIST_UNKNOWN),
	cSourceProps(0),
	lpSourceProps(nullptr),
	bItemFullyLoaded(false),
	m_lpData(nullptr)
{
	ulSortValue.QuadPart = NULL;
}

SortListData::~SortListData()
{
	Clean();
}

void SortListData::Clean()
{
	if (Contents() != nullptr)
	{
		delete Contents();
	}
	else if (Node() != nullptr)
	{
		delete Node();
	}

	m_lpData = nullptr;

	ulSortDataType = SORTLIST_UNKNOWN;
	MAPIFreeBuffer(lpSourceProps);
	lpSourceProps = nullptr;
	cSourceProps = 0;

	bItemFullyLoaded = false;
	szSortText = emptystring;

	ulSortValue.QuadPart = NULL;
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

// Sets data from the LPSRow into the SortListData structure
// Assumes the structure is either an existing structure or a new one which has been memset to 0
// If it's an existing structure - we need to free up some memory
// SORTLIST_CONTENTS
void SortListData::InitializeContents(_In_ LPSRow lpsRowData)
{
	Clean();
	if (!lpsRowData) return;

	ulSortDataType = SORTLIST_CONTENTS;
	lpSourceProps = lpsRowData->lpProps;
	cSourceProps = lpsRowData->cValues;
	m_lpData = new ContentsData(lpsRowData);
}

void SortListData::InitializeNode(
	ULONG cProps,
	_In_opt_ LPSPropValue lpProps,
	_In_opt_ LPSBinary lpEntryID,
	_In_opt_ LPSBinary lpInstanceKey,
	ULONG bSubfolders,
	ULONG ulContainerFlags)
{
	Clean();

	ulSortDataType = SORTLIST_TREENODE;
	cSourceProps = cProps;
	lpSourceProps = lpProps;
	m_lpData = new NodeData(
		lpEntryID,
		lpInstanceKey,
		bSubfolders,
		ulContainerFlags);
}

void SortListData::InitializeNode(_In_ LPSRow lpsRow)
{
	Clean();
	if (!lpsRow) return;

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

	InitializeNode(
		lpsRow->cValues,
		lpsRow->lpProps, // pass on props to be archived in node
		lpEIDBin,
		lpInstanceBin,
		lpSubfolders ? static_cast<ULONG>(lpSubfolders->Value.b) : MAPI_E_NOT_FOUND,
		lpContainerFlags ? lpContainerFlags->Value.ul : MAPI_E_NOT_FOUND);
}