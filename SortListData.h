#pragma once
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

class ContentsData
{
public:
	ContentsData(_In_ LPSRow lpsRowData);
	~ContentsData();
	LPSBinary lpEntryID; // Allocated with MAPIAllocateBuffer
	LPSBinary lpLongtermID; // Allocated with MAPIAllocateBuffer
	LPSBinary lpInstanceKey; // Allocated with MAPIAllocateBuffer
	LPSBinary lpServiceUID; // Allocated with MAPIAllocateBuffer
	LPSBinary lpProviderUID; // Allocated with MAPIAllocateBuffer
	TCHAR* szDN; // Allocated with MAPIAllocateBuffer
	CHAR* szProfileDisplayName; // Allocated with MAPIAllocateBuffer
	ULONG ulAttachNum;
	ULONG ulAttachMethod;
	ULONG ulRowID; // for recipients
	ULONG ulRowType; // PR_ROW_TYPE
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

struct _NodeData
{
	LPSBinary lpEntryID; // Allocated with MAPIAllocateBuffer
	LPSBinary lpInstanceKey; // Allocated with MAPIAllocateBuffer
	LPMAPITABLE lpHierarchyTable; // Object - free with Release
	CAdviseSink* lpAdviseSink; // Object - free with Release
	ULONG_PTR ulAdviseConnection;
	LONG cSubfolders; // -1 for unknown, 0 for no subfolders, >0 for at least one subfolder
};

class SortListData
{
public:
	~SortListData();
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
		_NodeData Node; // SORTLIST_TREENODE

	} data;
	ContentsData* Contents(); // SORTLIST_CONTENTS

	ULONG cSourceProps;
	LPSPropValue lpSourceProps; // Stolen from lpsRowData in CContentsTableListCtrl::BuildDataItem - free with MAPIFreeBuffer
	bool bItemFullyLoaded;

	LPVOID m_lpData;
};

// SORTLIST_TREENODE
SortListData* BuildNodeData(
	ULONG cProps,
	_In_opt_ LPSPropValue lpProps,
	_In_opt_ LPSBinary lpEntryID,
	_In_opt_ LPSBinary lpInstanceKey,
	ULONG bSubfolders,
	ULONG ulContainerFlags);

// SORTLIST_TREENODE
SortListData* BuildNodeData(_In_ LPSRow lpsRow);

// SORTLIST_CONTENTS
void BuildDataItem(_In_ LPSRow lpsRowData, _Inout_ SortListData* lpData);