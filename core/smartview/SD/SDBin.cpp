#include <core/stdafx.h>
#include <core/smartview/SD/SDBin.h>
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

			addHeader(L"Descriptor");
			addHeader(sd);
		}
	}
} // namespace smartview