#include <StdAfx.h>
#include <MrMapi/MMAccounts.h>
#include <MrMapi/mmcli.h>
#include <core/mapi/account/accountHelper.h>
#include <core/mapi/mapiFunctions.h>
#include <core/interpret/proptags.h>
#include <core/interpret/proptype.h>

void LogProp(LPOLKACCOUNT lpAccount, ULONG ulPropTag)
{
	auto pProp = ACCT_VARIANT();
	const auto hRes = WC_H(lpAccount->GetProp(ulPropTag, &pProp));
	if (SUCCEEDED(hRes))
	{
		const auto name = proptags::PropTagToPropName(ulPropTag, false).bestGuess;
		const auto tag = proptype::TypeToString(PROP_TYPE(ulPropTag));

		wprintf(
			L"Prop = %ws, Type = %ws, ",
			name.empty() ? strings::format(L"0x%08X", ulPropTag).c_str() : name.c_str(),
			tag.c_str());
		switch (pProp.dwType)
		{
		case PT_UNICODE:
			if (pProp.Val.pwsz)
			{
				wprintf(L"\"%ws\"", pProp.Val.pwsz);
				(void) lpAccount->FreeMemory(reinterpret_cast<LPBYTE>(pProp.Val.pwsz));
			}
			break;
		case PT_LONG:
			wprintf(L"0x%08lX", pProp.Val.dw);
			break;
		case PT_BINARY:
			wprintf(L"%ws", strings::BinToHexString(pProp.Val.bin.pb, pProp.Val.bin.cb, true).c_str());
			(void) lpAccount->FreeMemory(pProp.Val.bin.pb);
			break;
		}

		wprintf(L"\n");
	}
}

void IterateAllProps(LPOLKACCOUNT lpAccount)
{
	if (!lpAccount) return;

	wprintf(L"Iterating all properties\n");

	for (auto i = 0; i < 0x8000; i++)
	{
		auto pProp = ACCT_VARIANT{};
		LogProp(lpAccount, PROP_TAG(PT_LONG, i));
		LogProp(lpAccount, PROP_TAG(PT_UNICODE, i));
		LogProp(lpAccount, PROP_TAG(PT_BINARY, i));
	}

	wprintf(L"Done iterating all properties\n");
}

void EnumerateAccounts(LPMAPISESSION lpSession, LPCWSTR lpwszProfile, bool bIterateAllProps)
{
	LPOLKACCOUNTMANAGER lpAcctMgr = nullptr;

	EC_H_S(CoCreateInstance(
		CLSID_OlkAccountManager,
		nullptr,
		CLSCTX_INPROC_SERVER,
		IID_IOlkAccountManager,
		reinterpret_cast<LPVOID*>(&lpAcctMgr)));
	if (lpAcctMgr)
	{
		auto lpAcctHelper = new CAccountHelper(lpwszProfile, lpSession);
		if (lpAcctHelper)
		{
			const auto hRes = EC_H(lpAcctMgr->Init(lpAcctHelper, ACCT_INIT_NOSYNCH_MAPI_ACCTS));
			if (SUCCEEDED(hRes))
			{
				LPOLKENUM lpAcctEnum = nullptr;

				EC_H_S(lpAcctMgr->EnumerateAccounts(&CLSID_OlkMail, nullptr, OLK_ACCOUNT_NO_FLAGS, &lpAcctEnum));
				if (lpAcctEnum)
				{
					DWORD cAccounts = 0;

					EC_H_S(lpAcctEnum->GetCount(&cAccounts));
					if (cAccounts)
					{
						EC_H_S(lpAcctEnum->Reset());
						for (DWORD i = 0; i < cAccounts; i++)
						{
							if (i > 0) printf("\n");
							wprintf(L"Account #%lu\n", i + 1);
							LPUNKNOWN lpUnk = nullptr;

							EC_H(lpAcctEnum->GetNext(&lpUnk));
							if (lpUnk)
							{
								auto lpAccount = mapi::safe_cast<LPOLKACCOUNT>(lpUnk);

								if (lpAccount)
								{
									if (bIterateAllProps)
									{
										IterateAllProps(lpAccount);
									}
									else
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
									}

									lpAccount->Release();
								}
							}

							if (lpUnk) lpUnk->Release();
							lpUnk = nullptr;
						}
					}

					lpAcctEnum->Release();
				}
			}

			lpAcctHelper->Release();
		}

		lpAcctMgr->Release();
	}
}

void DisplayAccountList(LPMAPISESSION lpSession, LPCWSTR lpwszProfile, ULONG ulFlags)
{
	LPOLKACCOUNTMANAGER lpAcctMgr = nullptr;

	EC_H_S(CoCreateInstance(
		CLSID_OlkAccountManager,
		nullptr,
		CLSCTX_INPROC_SERVER,
		IID_IOlkAccountManager,
		reinterpret_cast<LPVOID*>(&lpAcctMgr)));
	if (lpAcctMgr)
	{
		auto lpAcctHelper = new CAccountHelper(lpwszProfile, lpSession);
		if (lpAcctHelper)
		{
			const auto hRes = EC_H(lpAcctMgr->Init(lpAcctHelper, ACCT_INIT_NOSYNCH_MAPI_ACCTS));
			if (SUCCEEDED(hRes))
			{
				EC_H_S(lpAcctMgr->DisplayAccountList(
					nullptr, // hwnd
					ulFlags, // dwFlags
					nullptr, // wszTitle
					NULL, // cCategories
					nullptr, // rgclsidCategories
					nullptr)); // pclsidType
			}

			lpAcctHelper->Release();
		}

		lpAcctMgr->Release();
	}
}

void DoAccounts(_In_opt_ LPMAPISESSION lpMAPISession)
{
	const auto flags = cli::switchFlag.atULONG(0, 16);
	const auto bIterate = cli::switchIterate.isSet();
	const auto bWizard = cli::switchWizard.isSet();
	const auto profileName = mapi::GetProfileName(lpMAPISession);

	wprintf(L"Enum Accounts\n");

	if (!profileName.empty())
	{
		if (bWizard)
		{
			DisplayAccountList(lpMAPISession, profileName.c_str(), flags);
		}
		else
		{
			EnumerateAccounts(lpMAPISession, profileName.c_str(), bIterate);
		}
	}
}