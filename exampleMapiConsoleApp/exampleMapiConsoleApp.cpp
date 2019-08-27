#include <windows.h>

#include <mapix.h>
#include <mapiutil.h>
#include <mapi.h>
#include <atlbase.h>

#include <iostream>
#include <iomanip>

#pragma warning(disable : 4127)

#define CORg(_hr) \
	do \
	{ \
		hr = _hr; \
		if (FAILED(hr)) \
		{ \
			std::wcout << L"FAILED! hr = " << std::hex << hr << L".  LINE = " << std::dec << __LINE__ << std::endl; \
			std::wcout << L" >>> " << (wchar_t*) L#_hr << std::endl; \
			goto Error; \
		} \
	} while (0)

int ListStoresTable(IMAPISession* pSession)
{
	HRESULT hr = S_OK;
	CComPtr<IMAPITable> spTable;
	SRowSet* pmrows = NULL;

	enum MAPIColumns
	{
		COL_ENTRYID = 0,
		COL_DISPLAYNAME_W,
		COL_DISPLAYNAME_A,
		cCols // End marker
	};

	static const SizedSPropTagArray(cCols, mcols) = {cCols,
													 {
														 PR_ENTRYID,
														 PR_DISPLAY_NAME_W,
														 PR_DISPLAY_NAME_A,
													 }};

	CORg(pSession->GetMsgStoresTable(0, &spTable));

	CORg(spTable->SetColumns((LPSPropTagArray) &mcols, TBL_BATCH));
	CORg(spTable->SeekRow(BOOKMARK_BEGINNING, 0, 0));

	CORg(spTable->QueryRows(50, 0, &pmrows));

	std::wcout << L"Found " << pmrows->cRows << L" stores in MAPI profile:" << std::endl;
	for (UINT i = 0; i != pmrows->cRows; i++)
	{
		SRow* prow = pmrows->aRow + i;
		LPCWSTR pwz = NULL;
		LPCSTR pwzA = NULL;
		if (PR_DISPLAY_NAME_W == prow->lpProps[COL_DISPLAYNAME_W].ulPropTag)
			pwz = prow->lpProps[COL_DISPLAYNAME_W].Value.lpszW;
		if (PR_DISPLAY_NAME_A == prow->lpProps[COL_DISPLAYNAME_A].ulPropTag)
			pwzA = prow->lpProps[COL_DISPLAYNAME_A].Value.lpszA;
		if (pwz)
			std::wcout << L" " << std::setw(4) << i << ": " << (pwz ? pwz : L"<NULL>") << std::endl;
		else
			std::wcout << L" " << std::setw(4) << i << ": " << (pwzA ? pwzA : "<NULL>") << std::endl;
	}

Error:
	FreeProws(pmrows);
	return 0;
}

void TestSimpleMapi()
{
	MapiMessage msg = {0};
	MapiRecipDesc recip = {0};

	recip.ulRecipClass = MAPI_TO;
	recip.lpszName = "John Doe";
	recip.lpszAddress = "johndoe@example.com";

	msg.lpszSubject = "Hello World";
	msg.lpszNoteText = "Test Body";
	msg.lpRecips = &recip;
	msg.nRecipCount = 1;

	MAPISendMail(NULL, 0, &msg, MAPI_LOGON_UI | MAPI_DIALOG, 0);
}

void ForceOutlookMAPI(bool fForce);

int __cdecl main()
{
	HRESULT hr = S_OK;

	MAPIINIT_0 MAPIINIT = {0, MAPI_MULTITHREAD_NOTIFICATIONS};

	//ForceOutlookMAPI(true);

	//TestSimpleMapi();

	CORg(MAPIInitialize(&MAPIINIT));

	{ // IMAPISession Smart Pointer Context
		CComPtr<IMAPISession> spSession;
		CORg(MAPILogonEx(NULL, NULL, NULL, MAPI_UNICODE | MAPI_LOGON_UI, &spSession));

		ListStoresTable(spSession);
	}

	MAPIUninitialize();

Error:
	return 0;
}
