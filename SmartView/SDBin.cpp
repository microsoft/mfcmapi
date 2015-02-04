#include "stdafx.h"
#include "..\stdafx.h"
#include "SDBin.h"
#include "..\String.h"
#include "..\MySecInfo.h"
#include "..\MAPIFunctions.h"
#include "..\ExtraPropTags.h"
#include "..\InterpretProp2.h"

SDBin::SDBin(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin, _In_opt_ LPMAPIPROP lpMAPIProp, bool bFB) : SmartViewParser(cbBin, lpBin)
{
	m_lpMAPIProp = lpMAPIProp;
	if (m_lpMAPIProp) m_lpMAPIProp->AddRef();
	m_bFB = bFB;
}

SDBin::~SDBin()
{
	if (m_lpMAPIProp) m_lpMAPIProp->Release();
}

void SDBin::Parse()
{
}

_Check_return_ wstring SDBin::ToStringInternal()
{
	HRESULT hRes = S_OK;
	LPBYTE lpSDToParse = m_Parser.GetCurrentAddress();
	ULONG ulSDToParse = m_Parser.RemainingBytes();
	m_Parser.Advance(ulSDToParse);

	eAceType acetype = acetypeMessage;
	switch (GetMAPIObjectType(m_lpMAPIProp))
	{
	case MAPI_STORE:
	case MAPI_ADDRBOOK:
	case MAPI_FOLDER:
	case MAPI_ABCONT:
		acetype = acetypeContainer;
		break;
	}

	if (m_bFB) acetype = acetypeFreeBusy;

	wstring szDACL;
	wstring szInfo;

	WC_H(SDToString(lpSDToParse, ulSDToParse, acetype, szDACL, szInfo));

	wstring szFlags = InterpretFlags(flagSecurityVersion, SECURITY_DESCRIPTOR_VERSION(lpSDToParse));
	wstring szResult = formatmessage(IDS_SECURITYDESCRIPTORHEADER);
	szResult += szInfo;
	szResult += formatmessage(IDS_SECURITYDESCRIPTORVERSION, SECURITY_DESCRIPTOR_VERSION(lpSDToParse), szFlags.c_str());
	szResult += szDACL;

	return szResult;
}