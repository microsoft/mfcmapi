#include <StdAfx.h>
#include <UI/ViewPane/SectionPane.h>

namespace viewpane
{
	SectionPane::SectionPane(const int paneID, const UINT uidLabel)
	{
		m_paneID = paneID;
		SetLabel(uidLabel);
		m_Header.SetSectionStyle();
	}

	int SectionPane::GetFixedHeight() { return GetHeaderHeight(); }
} // namespace viewpane
