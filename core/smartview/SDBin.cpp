#include <core/stdafx.h>
#include <core/smartview/SDBin.h>
#include <core/mapi/extraPropTags.h>
#include <core/interpret/flags.h>
#include <core/interpret/sid.h>
#include <core/mapi/mapiFunctions.h>

namespace smartview
{
	SDBin::~SDBin()
	{
		if (m_lpMAPIProp) m_lpMAPIProp->Release();
	}

	void SDBin::Init(_In_opt_ LPMAPIPROP lpMAPIProp, bool bFB) noexcept
	{
		if (m_lpMAPIProp) m_lpMAPIProp->Release();
		m_lpMAPIProp = lpMAPIProp;
		if (m_lpMAPIProp) m_lpMAPIProp->AddRef();
		m_bFB = bFB;
	}

	void SDBin::parse() { m_SDbin = blockBytes::parse(m_Parser, m_Parser->getSize()); }

	void SDBin::parseBlocks()
	{
		auto acetype = sid::aceType::Message;
		switch (mapi::GetMAPIObjectType(m_lpMAPIProp))
		{
		case MAPI_STORE:
		case MAPI_ADDRBOOK:
		case MAPI_FOLDER:
		case MAPI_ABCONT:
			acetype = sid::aceType::Container;
			break;
		}

		if (m_bFB) acetype = sid::aceType::FreeBusy;

		if (m_SDbin)
		{
			// TODO: more accurately break this parsing into blocks with proper offsets
			const auto sd = SDToString(*m_SDbin, acetype);
			setText(L"Security Descriptor:\r\n");
			addHeader(L"Security Info: ");
			addChild(m_SDbin, sd.info);

			terminateBlock();
			const auto sdVersion = SECURITY_DESCRIPTOR_VERSION(m_SDbin->data());
			auto szFlags = flags::InterpretFlags(flagSecurityVersion, sdVersion);
			addHeader(L"Security Version: 0x%1!04X! = %2!ws!\r\n", sdVersion, szFlags.c_str());
			addHeader(L"Descriptor:\r\n");
			addHeader(sd.dacl);
		}
	}
} // namespace smartview