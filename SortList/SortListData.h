#pragma once
class ContentsData;
class NodeData;
class PropListData;
class MVPropData;
class ResData;
class CommentData;

enum __SortListDataTypes
{
	SORTLIST_UNKNOWN = 0,
	SORTLIST_CONTENTS,
	SORTLIST_PROP,
	SORTLIST_MVPROP,
	SORTLIST_RES,
	SORTLIST_COMMENT,
	SORTLIST_BINARY,
	SORTLIST_TREENODE
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
	// SORTLIST_PROP
	void InitializePropList(_In_ ULONG ulPropTag);
	// SORTLIST_MVPROP
	void InitializeMV(_In_ LPSPropValue lpProp, ULONG iProp);
	void InitializeMV(_In_ LPSPropValue lpProp);
	// SORTLIST_RES
	void InitializeRes(_In_ LPSRestriction lpOldRes);
	// SORTLIST_COMMENT
	void InitializeComment(_In_ LPSPropValue lpOldProp);

	wstring szSortText;
	ULARGE_INTEGER ulSortValue;
	__SortListDataTypes m_Type;
	union
	{
		_BinaryData Binary; // SORTLIST_BINARY
	} data;
	ContentsData* Contents() const; // SORTLIST_CONTENTS
	NodeData* Node() const; // SORTLIST_TREENODE
	PropListData* Prop() const; // SORTLIST_PROP
	MVPropData* MV() const; // SORTLIST_MVPROP
	ResData* Res() const; // SORTLIST_RES
	CommentData* Comment() const; // SORTLIST_COMMENT

	ULONG cSourceProps;
	LPSPropValue lpSourceProps; // Stolen from lpsRowData in SortListData::InitializeContents - free with MAPIFreeBuffer
	bool bItemFullyLoaded;

	LPVOID m_lpData;
};