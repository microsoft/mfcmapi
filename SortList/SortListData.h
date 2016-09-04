#pragma once
class ContentsData;
class NodeData;

enum __SortListDataTypes
{
	SORTLIST_UNKNOWN = 0,
	SORTLIST_CONTENTS,
	SORTLIST_PROP,
	SORTLIST_MVPROP,
	SORTLIST_TAGARRAY,
	SORTLIST_RES,
	SORTLIST_COMMENT,
	SORTLIST_BINARY,
	SORTLIST_TREENODE
};

struct _PropListData
{
	ULONG ulPropTag;
};

struct _MVPropData
{
	_PV val; // Allocated with MAPIAllocateMore
};

struct _TagData
{
	ULONG ulPropTag;
};

struct _ResData
{
	LPSRestriction lpOldRes; // not allocated - just a pointer
	LPSRestriction lpNewRes; // Owned by an alloc parent - do not free
};

struct _CommentData
{
	LPSPropValue lpOldProp; // not allocated - just a pointer
	LPSPropValue lpNewProp; // Owned by an alloc parent - do not free
};

struct _BinaryData
{
	SBinary OldBin; // not allocated - just a pointer
	SBinary NewBin; // MAPIAllocateMore from m_lpNewEntryList
};

class SortListData
{
public:
	SortListData();
	~SortListData();
	void Clean();
	// SORTLIST_CONTENTS
	void InitializeContents(_In_ LPSRow lpsRowData);
	// SORTLIST_TREENODE
	void InitializeNode(
		ULONG cProps,
		_In_opt_ LPSPropValue lpProps,
		_In_opt_ LPSBinary lpEntryID,
		_In_opt_ LPSBinary lpInstanceKey,
		ULONG bSubfolders,
		ULONG ulContainerFlags);
	void InitializeNode(_In_ LPSRow lpsRow);

	wstring szSortText;
	ULARGE_INTEGER ulSortValue;
	ULONG ulSortDataType;
	union
	{
		_PropListData Prop; // SORTLIST_PROP
		_MVPropData MV; // SORTLIST_MVPROP
		_TagData Tag; // SORTLIST_TAGARRAY
		_ResData Res; // SORTLIST_RES
		_CommentData Comment; // SORTLIST_COMMENT
		_BinaryData Binary; // SORTLIST_BINARY
	} data;
	ContentsData* Contents() const; // SORTLIST_CONTENTS
	NodeData* Node() const; // SORTLIST_TREENODE

	ULONG cSourceProps;
	LPSPropValue lpSourceProps; // Stolen from lpsRowData in SortListData::InitializeContents - free with MAPIFreeBuffer
	bool bItemFullyLoaded;

	LPVOID m_lpData;
};