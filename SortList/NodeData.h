#pragma once
class SortListData;

class NodeData
{
public:
	NodeData(
		_In_opt_ LPSBinary lpEntryID,
		_In_opt_ LPSBinary lpInstanceKey,
		ULONG bSubfolders,
		ULONG ulContainerFlags);
	~NodeData();
	LPSBinary lpEntryID; // Allocated with MAPIAllocateBuffer
	LPSBinary lpInstanceKey; // Allocated with MAPIAllocateBuffer
	LPMAPITABLE lpHierarchyTable; // Object - free with Release
	CAdviseSink* lpAdviseSink; // Object - free with Release
	ULONG_PTR ulAdviseConnection;
	LONG cSubfolders; // -1 for unknown, 0 for no subfolders, >0 for at least one subfolder
};