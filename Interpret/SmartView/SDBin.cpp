#include <StdAfx.h>
#include <Interpret/SmartView/SDBin.h>
#include <Interpret/String.h>
#include <UI/MySecInfo.h>
#include <MAPI/MAPIFunctions.h>
#include <Interpret/ExtraPropTags.h>
#include <Interpret/InterpretProp.h>
#include <Interpret/Sid.h>

namespace smartview
{
	SDBin::SDBin()
	{
		m_lpMAPIProp = nullptr;
		m_bFB = false;
	}

	SDBin::~SDBin()
	{
		if (m_lpMAPIProp) m_lpMAPIProp->Release();
	}

	void SDBin::Init(_In_opt_ LPMAPIPROP lpMAPIProp, bool bFB)
	{
		if (m_lpMAPIProp) m_lpMAPIProp->Release();
		m_lpMAPIProp = lpMAPIProp;
		if (m_lpMAPIProp) m_lpMAPIProp->AddRef();
		m_bFB = bFB;
	}

	void SDBin::Parse() {}

	_Check_return_ std::wstring SDBin::ToStringInternal()
	{
		const auto lpSDToParse = m_Parser.GetCurrentAddress();
		const auto ulSDToParse = m_Parser.RemainingBytes();
		m_Parser.Advance(ulSDToParse);

		auto acetype = sid::acetypeMessage;
		switch (mapi::GetMAPIObjectType(m_lpMAPIProp))
		{
		case MAPI_STORE:
		case MAPI_ADDRBOOK:
		case MAPI_FOLDER:
		case MAPI_ABCONT:
			acetype = sid::acetypeContainer;
			break;
		}

		if (m_bFB) acetype = sid::acetypeFreeBusy;

		auto securityDescriptor = SDToString(lpSDToParse, ulSDToParse, acetype);
		auto szDACL = securityDescriptor.dacl;
		auto szInfo = securityDescriptor.info;
		auto szFlags = interpretprop::InterpretFlags(flagSecurityVersion, SECURITY_DESCRIPTOR_VERSION(lpSDToParse));

		std::vector<std::wstring> result;
		result.push_back(strings::formatmessage(IDS_SECURITYDESCRIPTORHEADER) + szInfo);
		result.push_back(strings::formatmessage(
			IDS_SECURITYDESCRIPTORVERSION, SECURITY_DESCRIPTOR_VERSION(lpSDToParse), szFlags.c_str()));
		result.push_back(szDACL);

		return strings::join(result, L"\r\n");
	}
}