#include "stdafx.h"
#include "ContentsData.h"
#include "MAPIFunctions.h"

ContentsData::ContentsData(_In_ LPSRow lpsRowData)
{
	lpEntryID = nullptr;
	lpLongtermID = nullptr;
	lpInstanceKey = nullptr;
	lpServiceUID = nullptr;
	lpProviderUID = nullptr;
	ulAttachNum = 0;
	ulAttachMethod = 0;
	ulRowID = 0;
	ulRowType = 0;

	if (!lpsRowData) return;

	auto hRes = S_OK;

	// Save the instance key into lpData
	auto lpProp = PpropFindProp(
		lpsRowData->lpProps,
		lpsRowData->cValues,
		PR_INSTANCE_KEY);
	if (lpProp && PR_INSTANCE_KEY == lpProp->ulPropTag)
	{
		EC_H(MAPIAllocateBuffer(
			static_cast<ULONG>(sizeof(SBinary)),
			reinterpret_cast<LPVOID*>(&lpInstanceKey)));
		EC_H(CopySBinary(lpInstanceKey, &lpProp->Value.bin, lpInstanceKey));
	}

	// Save the attachment number into lpData
	lpProp = PpropFindProp(
		lpsRowData->lpProps,
		lpsRowData->cValues,
		PR_ATTACH_NUM);
	if (lpProp && PR_ATTACH_NUM == lpProp->ulPropTag)
	{
		DebugPrint(DBGGeneric, L"\tPR_ATTACH_NUM = %d\n", lpProp->Value.l);
		ulAttachNum = lpProp->Value.l;
	}

	lpProp = PpropFindProp(
		lpsRowData->lpProps,
		lpsRowData->cValues,
		PR_ATTACH_METHOD);
	if (lpProp && PR_ATTACH_METHOD == lpProp->ulPropTag)
	{
		DebugPrint(DBGGeneric, L"\tPR_ATTACH_METHOD = %d\n", lpProp->Value.l);
		ulAttachMethod = lpProp->Value.l;
	}

	// Save the row ID (recipients) into lpData
	lpProp = PpropFindProp(
		lpsRowData->lpProps,
		lpsRowData->cValues,
		PR_ROWID);
	if (lpProp && PR_ROWID == lpProp->ulPropTag)
	{
		DebugPrint(DBGGeneric, L"\tPR_ROWID = %d\n", lpProp->Value.l);
		ulRowID = lpProp->Value.l;
	}

	// Save the row type (header/leaf) into lpData
	lpProp = PpropFindProp(
		lpsRowData->lpProps,
		lpsRowData->cValues,
		PR_ROW_TYPE);
	if (lpProp && PR_ROW_TYPE == lpProp->ulPropTag)
	{
		DebugPrint(DBGGeneric, L"\tPR_ROW_TYPE = %d\n", lpProp->Value.l);
		ulRowType = lpProp->Value.l;
	}

	// Save the Entry ID into lpData
	lpProp = PpropFindProp(
		lpsRowData->lpProps,
		lpsRowData->cValues,
		PR_ENTRYID);
	if (lpProp && PR_ENTRYID == lpProp->ulPropTag)
	{
		EC_H(MAPIAllocateBuffer(
			static_cast<ULONG>(sizeof(SBinary)),
			reinterpret_cast<LPVOID*>(&lpEntryID)));
		EC_H(CopySBinary(lpEntryID, &lpProp->Value.bin, lpEntryID));
	}

	// Save the Longterm Entry ID into lpData
	lpProp = PpropFindProp(
		lpsRowData->lpProps,
		lpsRowData->cValues,
		PR_LONGTERM_ENTRYID_FROM_TABLE);
	if (lpProp && PR_LONGTERM_ENTRYID_FROM_TABLE == lpProp->ulPropTag)
	{
		EC_H(MAPIAllocateBuffer(
			static_cast<ULONG>(sizeof(SBinary)),
			reinterpret_cast<LPVOID*>(&lpLongtermID)));
		EC_H(CopySBinary(lpLongtermID, &lpProp->Value.bin, lpLongtermID));
	}

	// Save the Service ID into lpData
	lpProp = PpropFindProp(
		lpsRowData->lpProps,
		lpsRowData->cValues,
		PR_SERVICE_UID);
	if (lpProp && PR_SERVICE_UID == lpProp->ulPropTag)
	{
		// Allocate some space
		EC_H(MAPIAllocateBuffer(
			static_cast<ULONG>(sizeof(SBinary)),
			reinterpret_cast<LPVOID*>(&lpServiceUID)));
		EC_H(CopySBinary(lpServiceUID, &lpProp->Value.bin, lpServiceUID));
	}

	// Save the Provider ID into lpData
	lpProp = PpropFindProp(
		lpsRowData->lpProps,
		lpsRowData->cValues,
		PR_PROVIDER_UID);
	if (lpProp && PR_PROVIDER_UID == lpProp->ulPropTag)
	{
		// Allocate some space
		EC_H(MAPIAllocateBuffer(
			static_cast<ULONG>(sizeof(SBinary)),
			reinterpret_cast<LPVOID*>(&lpProviderUID)));
		EC_H(CopySBinary(lpProviderUID, &lpProp->Value.bin, lpProviderUID));
	}

	// Save the DisplayName into lpData
	lpProp = PpropFindProp(
		lpsRowData->lpProps,
		lpsRowData->cValues,
		PR_DISPLAY_NAME_A); // We pull this properties for profiles, which do not support Unicode
	if (CheckStringProp(lpProp, PT_STRING8))
	{
		m_szProfileDisplayName = lpProp->Value.lpszA;
		DebugPrint(DBGGeneric, L"\tPR_DISPLAY_NAME_A = %hs\n", m_szProfileDisplayName.c_str());
	}

	// Save the e-mail address (if it exists on the object) into lpData
	lpProp = PpropFindProp(
		lpsRowData->lpProps,
		lpsRowData->cValues,
		PR_EMAIL_ADDRESS);
	if (CheckStringProp(lpProp, PT_TSTRING))
	{
		m_szDN = LPCTSTRToWstring(lpProp->Value.LPSZ);
		DebugPrint(DBGGeneric, L"\tPR_EMAIL_ADDRESS = %ws\n", m_szDN.c_str());
	}
}

ContentsData::~ContentsData()
{
	MAPIFreeBuffer(lpInstanceKey);
	MAPIFreeBuffer(lpEntryID);
	MAPIFreeBuffer(lpLongtermID);
	MAPIFreeBuffer(lpServiceUID);
	MAPIFreeBuffer(lpProviderUID);
}