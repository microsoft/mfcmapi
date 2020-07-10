#include <core/stdafx.h>
#include <core/smartview/addinParser.h>
#include <core/addin/addin.h>

namespace smartview
{
	void addinParser::parse() { bin = blockBytes::parse(parser, parser->getSize()); }

	void addinParser::parseBlocks()
	{
		auto sv = addin::AddInSmartView(type, static_cast<ULONG>(bin->size()), bin->data());
		if (!sv.empty())
		{
			setText(addin::AddInStructTypeToString(type));
			addChild(bin, sv);
		}
		else
		{
			setText(L"Unknown Parser %1!d!", type);
			addChild(bin);
		}
	}
} // namespace smartview