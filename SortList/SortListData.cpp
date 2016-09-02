#include "stdafx.h"
#include "SortListData.h"
#include "ContentsData.h"
#include "NodeData.h"

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
	}

	if (Contents() != nullptr) delete Contents();
	if (Node() != nullptr) delete Node();

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

NodeData* SortListData::Node() const
{
	if (ulSortDataType == SORTLIST_TREENODE)
	{
		return reinterpret_cast<NodeData*>(m_lpData);
	}

	return nullptr;
}