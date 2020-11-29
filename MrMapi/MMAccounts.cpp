#include <StdAfx.h>
#include <MrMapi/MMAccounts.h>
#include <MrMapi/mmcli.h>
#include <core/mapi/account/accountHelper.h>
#include <core/mapi/mapiFunctions.h>
#include <core/interpret/proptags.h>

void IterateAllProps(LPOLKACCOUNT lpAccount);
HRESULT EnumerateAccounts(LPMAPISESSION lpSession, LPCWSTR lpwszProfile, bool bIterateAllProps);
HRESULT DisplayAccountList(LPMAPISESSION lpSession, LPCWSTR lpwszProfile, ULONG ulFlags);
void PrintBinary(DWORD cb, const BYTE* lpb);

void LogProp(LPOLKACCOUNT lpAccount, ULONG ulPropTag)
{
	auto pProp = ACCT_VARIANT();
	auto hRes = lpAccount->GetProp(ulPropTag, &pProp);
	auto szName = proptags::TagToString(ulPropTag, nullptr, false, true);
	if (SUCCEEDED(hRes))
	{
		switch (PROP_TYPE(ulPropTag))
		{
		case PT_UNICODE:
			if (pProp.Val.pwsz)
			{
				wprintf(L"%ws = \"%ws\"\n", szName.c_str(), pProp.Val.pwsz);
				(void) lpAccount->FreeMemory(reinterpret_cast<LPBYTE>(pProp.Val.pwsz));
			}
			break;
		case PT_LONG:
			wprintf(L"%ws = 0x%08lX\n", szName.c_str(), pProp.Val.dw);
			break;
		case PT_BINARY:
			wprintf(L"%ws = ", szName.c_str());
			PrintBinary(pProp.Val.bin.cb, pProp.Val.bin.pb);
			wprintf(L"\n");
			(void) lpAccount->FreeMemory(pProp.Val.bin.pb);
			break;
		}
	}
}

void IterateAllProps(LPOLKACCOUNT lpAccount)
{
	if (!lpAccount) return;

	wprintf(L"Iterating all properties\n");
	auto hRes = S_OK;
	ULONG i = 0;
	ACCT_VARIANT pProp = {0};

	for (i = 0; i < 0x8000; i++)
	{
		memset(&pProp, 0, sizeof(ACCT_VARIANT));
		hRes = lpAccount->GetProp(PROP_TAG(PT_LONG, i), &pProp);
		if (SUCCEEDED(hRes))
		{
			wprintf(L"Prop = 0x%08lX, Type = PT_LONG, Value = 0x%08lX\n", PROP_TAG(PT_LONG, i), pProp.Val.dw);
		}

		hRes = lpAccount->GetProp(PROP_TAG(PT_UNICODE, i), &pProp);
		if (SUCCEEDED(hRes))
		{
			wprintf(L"Prop = 0x%08lX, Type = PT_UNICODE, Value = %ws\n", PROP_TAG(PT_UNICODE, i), pProp.Val.pwsz);
			(void) lpAccount->FreeMemory(reinterpret_cast<LPBYTE>(pProp.Val.pwsz));
		}

		hRes = lpAccount->GetProp(PROP_TAG(PT_BINARY, i), &pProp);
		if (SUCCEEDED(hRes))
		{
			wprintf(L"Prop = 0x%08lX, Type = PT_BINARY, Value = ", PROP_TAG(PT_BINARY, i));
			PrintBinary(pProp.Val.bin.cb, pProp.Val.bin.pb);
			wprintf(L"\n");
			(void) lpAccount->FreeMemory(pProp.Val.bin.pb);
		}
	}

	wprintf(L"Done iterating all properties\n");
}

HRESULT EnumerateAccounts(LPMAPISESSION lpSession, LPCWSTR lpwszProfile, bool bIterateAllProps)
{
	auto hRes = S_OK;
	LPOLKACCOUNTMANAGER lpAcctMgr = nullptr;

	hRes = EC_H(CoCreateInstance(
		CLSID_OlkAccountManager,
		nullptr,
		CLSCTX_INPROC_SERVER,
		IID_IOlkAccountManager,
		reinterpret_cast<LPVOID*>(&lpAcctMgr)));
	if (SUCCEEDED(hRes) && lpAcctMgr)
	{
		auto pMyAcctHelper = new CAccountHelper(lpwszProfile, lpSession);
		if (pMyAcctHelper)
		{
			LPOLKACCOUNTHELPER lpAcctHelper = nullptr;
			hRes = EC_H(pMyAcctHelper->QueryInterface(IID_IOlkAccountHelper, reinterpret_cast<LPVOID*>(&lpAcctHelper)));
			if (SUCCEEDED(hRes) && lpAcctHelper)
			{
				hRes = EC_H(lpAcctMgr->Init(lpAcctHelper, ACCT_INIT_NOSYNCH_MAPI_ACCTS));
				if (SUCCEEDED(hRes))
				{
					LPOLKENUM lpAcctEnum = nullptr;

					hRes =
						EC_H(lpAcctMgr->EnumerateAccounts(&CLSID_OlkMail, nullptr, OLK_ACCOUNT_NO_FLAGS, &lpAcctEnum));
					if (SUCCEEDED(hRes) && lpAcctEnum)
					{
						DWORD cAccounts = 0;

						hRes = EC_H(lpAcctEnum->GetCount(&cAccounts));
						if (SUCCEEDED(hRes) && cAccounts)
						{
							hRes = EC_H(lpAcctEnum->Reset());
							if (SUCCEEDED(hRes))
							{
								DWORD i = 0;
								for (i = 0; i < cAccounts; i++)
								{
									if (i > 0) printf("\n");
									wprintf(L"Account #%lu\n", i + 1);
									LPUNKNOWN lpUnk = nullptr;

									hRes = EC_H(lpAcctEnum->GetNext(&lpUnk));
									if (SUCCEEDED(hRes) && lpUnk)
									{
										LPOLKACCOUNT lpAccount = nullptr;

										hRes = EC_H(lpUnk->QueryInterface(
											IID_IOlkAccount, reinterpret_cast<LPVOID*>(&lpAccount)));
										if (lpAccount)
										{
											LogProp(lpAccount, PROP_ACCT_NAME);
											LogProp(lpAccount, PROP_ACCT_USER_DISPLAY_NAME);
											LogProp(lpAccount, PROP_ACCT_USER_EMAIL_ADDR);
											LogProp(lpAccount, PROP_ACCT_STAMP);
											LogProp(lpAccount, PROP_ACCT_SEND_STAMP);
											LogProp(lpAccount, PROP_ACCT_IS_EXCH);
											LogProp(lpAccount, PROP_ACCT_ID);
											LogProp(lpAccount, PROP_ACCT_DELIVERY_FOLDER);
											LogProp(lpAccount, PROP_ACCT_DELIVERY_STORE);
											LogProp(lpAccount, PROP_ACCT_SENTITEMS_EID);
											LogProp(lpAccount, PR_NEXT_SEND_ACCT);
											LogProp(lpAccount, PR_PRIMARY_SEND_ACCT);
											LogProp(lpAccount, PROP_MAPI_IDENTITY_ENTRYID);
											LogProp(lpAccount, PROP_MAPI_TRANSPORT_FLAGS);

											if (bIterateAllProps) IterateAllProps(lpAccount);
											lpAccount->Release();
										}
									}

									if (lpUnk) lpUnk->Release();
									lpUnk = nullptr;
								}
							}
						}
					}

					if (lpAcctEnum) lpAcctEnum->Release();
				}
			}

			if (lpAcctHelper) lpAcctHelper->Release();
			pMyAcctHelper->Release();
		}
	}

	if (lpAcctMgr) lpAcctMgr->Release();

	return hRes;
}

HRESULT DisplayAccountList(LPMAPISESSION lpSession, LPCWSTR lpwszProfile, ULONG ulFlags)
{
	auto hRes = S_OK;
	LPOLKACCOUNTMANAGER lpAcctMgr = nullptr;

	hRes = EC_H(CoCreateInstance(
		CLSID_OlkAccountManager,
		nullptr,
		CLSCTX_INPROC_SERVER,
		IID_IOlkAccountManager,
		reinterpret_cast<LPVOID*>(&lpAcctMgr)));
	if (SUCCEEDED(hRes) && lpAcctMgr)
	{
		auto pMyAcctHelper = new CAccountHelper(lpwszProfile, lpSession);
		if (pMyAcctHelper)
		{
			LPOLKACCOUNTHELPER lpAcctHelper = nullptr;
			hRes = EC_H(pMyAcctHelper->QueryInterface(IID_IOlkAccountHelper, reinterpret_cast<LPVOID*>(&lpAcctHelper)));
			if (SUCCEEDED(hRes) && lpAcctHelper)
			{
				hRes = EC_H(lpAcctMgr->Init(lpAcctHelper, ACCT_INIT_NOSYNCH_MAPI_ACCTS));
				if (SUCCEEDED(hRes))
				{
					hRes = EC_H(lpAcctMgr->DisplayAccountList(
						nullptr, // hwnd
						ulFlags, // dwFlags
						nullptr, // wszTitle
						NULL, // cCategories
						nullptr, // rgclsidCategories
						nullptr)); // pclsidType
				}
			}

			if (lpAcctHelper) lpAcctHelper->Release();
			pMyAcctHelper->Release();
		}
	}

	if (lpAcctMgr) lpAcctMgr->Release();

	return hRes;
}

void PrintBinary(const DWORD cb, const BYTE* lpb)
{
	if (!cb || !lpb) return;
	LPSTR lpszHex = nullptr;
	ULONG i = 0;
	ULONG iBinPos = 0;
	lpszHex = new CHAR[1 + 2 * cb];
	if (lpszHex)
	{
		for (i = 0; i < cb; i++)
		{
			const auto bLow = static_cast<BYTE>(lpb[i] & 0xf);
			const auto bHigh = static_cast<BYTE>(lpb[i] >> 4 & 0xf);
			const auto szLow = static_cast<CHAR>(bLow <= 0x9 ? '0' + bLow : 'A' + bLow - 0xa);
			const auto szHigh = static_cast<CHAR>(bHigh <= 0x9 ? '0' + bHigh : 'A' + bHigh - 0xa);

			lpszHex[iBinPos] = szHigh;
			lpszHex[iBinPos + 1] = szLow;

			iBinPos += 2;
		}

		lpszHex[iBinPos] = '\0';
		wprintf(L"%hs", lpszHex);
		delete[] lpszHex;
	}
}

void DoAccounts(_In_opt_ LPMAPISESSION lpMAPISession)
{
	const auto flags = cli::switchFlag.atULONG(0, 16);
	const auto bIterate = cli::switchIterate.isSet();
	const auto bWizard = cli::switchWizard.isSet();
	auto profileName = mapi::GetProfileName(lpMAPISession);

	wprintf(L"Enum Accounts\n");
	wprintf(L"Options specified:\n");
	wprintf(L"   Profile: %ws\n", profileName.c_str());
	wprintf(L"   Flags: 0x%03X\n", flags);
	wprintf(L"   Iterate: %ws\n", bIterate ? L"true" : L"false");
	wprintf(L"   Wizard: %ws\n", bWizard ? L"true" : L"false");

	if (!profileName.empty())
	{
		if (bWizard)
		{
			(void) DisplayAccountList(lpMAPISession, profileName.c_str(), flags);
		}
		else
		{
			(void) EnumerateAccounts(lpMAPISession, profileName.c_str(), bIterate);
		}
	}
}