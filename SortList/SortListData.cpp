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
	m_Type(SORTLIST_UNKNOWN),
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
	switch (m_Type)
	{
	case SORTLIST_UNKNOWN: break;
	case SORTLIST_CONTENTS: delete Contents(); break;
	case SORTLIST_PROP: delete Prop(); break;
	case SORTLIST_MVPROP: delete MV(); break;
	case SORTLIST_RES: delete Res(); break;
	case SORTLIST_COMMENT: delete Comment(); break;
	case SORTLIST_BINARY: delete Binary(); break;
	case SORTLIST_TREENODE: delete Node(); break;
	default: break;
	}

	m_lpData = nullptr;

	m_Type = SORTLIST_UNKNOWN;
	MAPIFreeBuffer(lpSourceProps);
	lpSourceProps = nullptr;
	cSourceProps = 0;

	bItemFullyLoaded = false;
	szSortText.clear();

	ulSortValue.QuadPart = NULL;
}

ContentsData* SortListData::Contents() const
{
	if (m_Type == SORTLIST_CONTENTS)
	{
		return reinterpret_cast<ContentsData*>(m_lpData);
	}

	return nullptr;
}

NodeData* SortListData::Node() const
{
	if (m_Type == SORTLIST_TREENODE)
	{
		return reinterpret_cast<NodeData*>(m_lpData);
	}

	return nullptr;
}

PropListData* SortListData::Prop() const
{
	if (m_Type == SORTLIST_PROP)
	{
		return reinterpret_cast<PropListData*>(m_lpData);
	}

	return nullptr;
}

MVPropData* SortListData::MV() const
{
	if (m_Type == SORTLIST_MVPROP)
	{
		return reinterpret_cast<MVPropData*>(m_lpData);
	}

	return nullptr;
}

ResData* SortListData::Res() const
{
	if (m_Type == SORTLIST_RES)
	{
		return reinterpret_cast<ResData*>(m_lpData);
	}

	return nullptr;
}

CommentData* SortListData::Comment() const
{
	if (m_Type == SORTLIST_COMMENT)
	{
		return reinterpret_cast<CommentData*>(m_lpData);
	}

	return nullptr;
}

BinaryData* SortListData::Binary() const
{
	if (m_Type == SORTLIST_BINARY)
	{
		return reinterpret_cast<BinaryData*>(m_lpData);
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
	m_Type = SORTLIST_CONTENTS;

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
	m_Type = SORTLIST_TREENODE;
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
	m_Type = SORTLIST_PROP;
	bItemFullyLoaded = true;
	m_lpData = new PropListData(ulPropTag);
}

void SortListData::InitializeMV(_In_ LPSPropValue lpProp, ULONG iProp)
{
	Clean();
	m_Type = SORTLIST_MVPROP;
	m_lpData = new MVPropData(lpProp, iProp);
}

void SortListData::InitializeMV(_In_opt_ LPSPropValue lpProp)
{
	Clean();
	m_Type = SORTLIST_MVPROP;
	m_lpData = new MVPropData(lpProp);
}

void SortListData::InitializeRes(_In_ LPSRestriction lpOldRes)
{
	Clean();
	m_Type = SORTLIST_RES;
	bItemFullyLoaded = true;
	m_lpData = new ResData(lpOldRes);
}

void SortListData::InitializeComment(_In_ LPSPropValue lpOldProp)
{
	Clean();
	m_Type = SORTLIST_COMMENT;
	bItemFullyLoaded = true;
	m_lpData = new CommentData(lpOldProp);
}

void SortListData::InitializeBinary(_In_ LPSBinary lpOldBin)
{
	Clean();
	m_Type = SORTLIST_BINARY;
	m_lpData = new BinaryData(lpOldBin);
}