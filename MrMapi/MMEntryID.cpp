#include <StdAfx.h>
#include <MrMapi/MMStore.h>
#include <MrMapi/mmcli.h>
#include <core/interpret/proptags.h>
#include <core/mapi/mapiOutput.h>
#include <core/utility/output.h>
#include <core/utility/registry.h>

HRESULT DoEntryID(_In_opt_ LPMDB lpMDB)
{
	if (!lpMDB) return MAPI_E_NOT_INITIALIZED;

	auto entryID = strings::HexStringToBin(cli::switchEntryID[0]);
	ULONG objectType{};
	LPMESSAGE message{};
	auto hRes = WC_MAPI(lpMDB->OpenEntry(
		static_cast<ULONG>(entryID.size()),
		reinterpret_cast<LPENTRYID>(entryID.data()),
		&IID_IMessage,
		registry::forceMapiNoCache ? MAPI_NO_CACHE : 0,
		&objectType,
		reinterpret_cast<LPUNKNOWN*>(&message)));
	if (SUCCEEDED(hRes) && (!message || objectType != MAPI_MESSAGE)) hRes = MAPI_E_INVALID_OBJECT;
	if (SUCCEEDED(hRes))
	{
		SPropTagArray tags = {1, proptags::PropNameToPropTag(cli::switchUnswitched[0])};
		ULONG count{};
		LPSPropValue props{};
		hRes = WC_H_GETPROPS(message->GetProps(&tags, fMapiUnicode, &count, &props));
		if (SUCCEEDED(hRes) && props && count == 1)
		{
			wprintf(L"%ws<messageprops>\n", output::g_szXMLHeader.c_str());
			output::outputProperties(output::dbgLevel::NoDebug, stdout, count, props, message, true);
			wprintf(L"</messageprops>\n");
			if (PROP_TYPE(props[0].ulPropTag) == PT_ERROR) hRes = props[0].Value.err;
		}
		else if (SUCCEEDED(hRes))
		{
			hRes = MAPI_E_CALL_FAILED;
		}

		MAPIFreeBuffer(props);
	}

	if (message) message->Release();
	return hRes;
}
