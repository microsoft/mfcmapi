#include "stdafx.h"
#include "SortListData.h"
#include "ContentsData.h"
#include "NodeData.h"
#include "PropListData.h"
#include "MVPropData.h"
#include "ResData.h"
#include "CommentData.h"
#include "BinaryData.h"

SortListData::SortListData() :
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
	if (m_lpData) delete m_lpData;
	m_lpData = nullptr;

	MAPIFreeBuffer(lpSourceProps);
	lpSourceProps = nullptr;
	cSourceProps = 0;

	bItemFullyLoaded = false;
	szSortText.clear();

	ulSortValue.QuadPart = NULL;
}

ContentsData* SortListData::Contents() const
{
	return reinterpret_cast<ContentsData*>(m_lpData);
}

NodeData* SortListData::Node() const
{
	return reinterpret_cast<NodeData*>(m_lpData);
}

PropListData* SortListData::Prop() const
{
	return reinterpret_cast<PropListData*>(m_lpData);
}

MVPropData* SortListData::MV() const
{
	return reinterpret_cast<MVPropData*>(m_lpData);
}

ResData* SortListData::Res() const
{
	return reinterpret_cast<ResData*>(m_lpData);
}

CommentData* SortListData::Comment() const
{
	return reinterpret_cast<CommentData*>(m_lpData);
}

BinaryData* SortListData::Binary() const
{
	return reinterpret_cast<BinaryData*>(m_lpData);
}

// Sets data from the LPSRow into the SortListData structure
// Assumes the structure is either an existing structure or a new one which has been memset to 0
// If it's an existing structure - we need to free up some memory
// SORTLIST_CONTENTS
void SortListData::InitializeContents(_In_ LPSRow lpsRowData)
{
	Clean();

	if (!lpsRowData) return;
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

void SortListData::InitializePropList(_In_ ULONG ulPropTag)
{
	Clean();
	bItemFullyLoaded = true;
	m_lpData = new PropListData(ulPropTag);
}

void SortListData::InitializeMV(_In_ LPSPropValue lpProp, ULONG iProp)
{
	Clean();
	m_lpData = new MVPropData(lpProp, iProp);
}

void SortListData::InitializeMV(_In_opt_ LPSPropValue lpProp)
{
	Clean();
	m_lpData = new MVPropData(lpProp);
}

void SortListData::InitializeRes(_In_ LPSRestriction lpOldRes)
{
	Clean();
	bItemFullyLoaded = true;
	m_lpData = new ResData(lpOldRes);
}

void SortListData::InitializeComment(_In_ LPSPropValue lpOldProp)
{
	Clean();
	bItemFullyLoaded = true;
	m_lpData = new CommentData(lpOldProp);
}

void SortListData::InitializeBinary(_In_ LPSBinary lpOldBin)
{
	Clean();
	m_lpData = new BinaryData(lpOldBin);
}