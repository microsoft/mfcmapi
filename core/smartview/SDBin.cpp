#include <core/stdafx.h>
#include <core/smartview/SDBin.h>
#include <core/mapi/extraPropTags.h>
#include <core/interpret/flags.h>
#include <core/interpret/sid.h>
#include <core/mapi/mapiFunctions.h>

namespace smartview
{
	SDBin::SDBin(_In_opt_ LPMAPIPROP lpMAPIProp, bool bFB)
	{
		switch (mapi::GetMAPIObjectType(lpMAPIProp))
		{
		case MAPI_STORE:
		case MAPI_ADDRBOOK:
		case MAPI_FOLDER:
		case MAPI_ABCONT:
			acetype = sid::aceType::Container;
			break;
		}

		if (bFB) acetype = sid::aceType::FreeBusy;
	}

	void SDBin::parse() { m_SDbin = blockBytes::parse(parser, parser->getSize()); }

	void SDBin::parseBlocks()
	{
		if (m_SDbin)
		{
			setText(L"Security Descriptor");

			// TODO: more accurately break this parsing into blocks with proper offsets
			const auto sd = SDToString(*m_SDbin, acetype);
			auto si = create(L"Security Info");
			addChild(si);
			if (!sd.info.empty())
			{
				si->addChild(m_SDbin, sd.info);
			}

			if (m_SDbin->size() >= 2 * sizeof(WORD))
			{
				const auto sdVersion = SECURITY_DESCRIPTOR_VERSION(m_SDbin->data());
				auto szFlags = flags::InterpretFlags(flagSecurityVersion, sdVersion);
				addHeader(L"Security Version: 0x%1!04X! = %2!ws!", sdVersion, szFlags.c_str());
			}

			addHeader(L"Descriptor");
			addHeader(sd.dacl);
		}
	}
} // namespace smartview