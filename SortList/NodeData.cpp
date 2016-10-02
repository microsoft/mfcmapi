#include "stdafx.h"
#include "NodeData.h"
#include "AdviseSink.h"
#include "MAPIFunctions.h"

NodeData::NodeData(
	_In_opt_ LPSBinary lpEntryID,
	_In_opt_ LPSBinary lpInstanceKey,
	ULONG bSubfolders,
	ULONG ulContainerFlags) : m_lpEntryID(nullptr), m_lpInstanceKey(nullptr), m_lpHierarchyTable(nullptr), m_lpAdviseSink(nullptr), m_ulAdviseConnection(0)
{
	auto hRes = S_OK;
	if (lpEntryID)
	{
		WC_H(MAPIAllocateBuffer(
			static_cast<ULONG>(sizeof(SBinary)),
			reinterpret_cast<LPVOID*>(&m_lpEntryID)));

		// Copy the data over
		WC_H(CopySBinary(
			m_lpEntryID,
			lpEntryID,
			nullptr));
	}

	if (lpInstanceKey)
	{
		WC_H(MAPIAllocateBuffer(
			static_cast<ULONG>(sizeof(SBinary)),
			reinterpret_cast<LPVOID*>(&m_lpInstanceKey)));
		WC_H(CopySBinary(
			m_lpInstanceKey,
			lpInstanceKey,
			nullptr));
	}

	m_cSubfolders = -1;
	if (bSubfolders != MAPI_E_NOT_FOUND)
	{
		m_cSubfolders = bSubfolders != 0;
	}
	else if (ulContainerFlags != MAPI_E_NOT_FOUND)
	{
		m_cSubfolders = ulContainerFlags & AB_SUBCONTAINERS ? 1 : 0;
	}
}

NodeData::~NodeData()
{
	MAPIFreeBuffer(m_lpInstanceKey);
	MAPIFreeBuffer(m_lpEntryID);

	if (m_lpAdviseSink)
	{
		// unadvise before releasing our sink
		if (m_ulAdviseConnection &&  m_lpHierarchyTable)
			m_lpHierarchyTable->Unadvise(m_ulAdviseConnection);
		m_lpAdviseSink->Release();
	}

	if (m_lpHierarchyTable)  m_lpHierarchyTable->Release();
}