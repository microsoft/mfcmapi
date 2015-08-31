#include "StdAfx.h"
#include "Property.h"
#include "ColumnTags.h"
#include "String.h"

wstring ScrubStringForXML(_In_ wstring szString)
{
	size_t i = 0;

	for (i = 0; i < szString.length(); i++)
	{
		switch (szString[i])
		{
		case L'\t':
		case L'\r':
		case L'\n':
			break;
		default:
			if (szString[i] < 0x20)
			{
				szString[i] = L'.';
			}

			break;
		}
	}

	return szString;
}

wstring indent(int iIndent)
{
	wstring szIndent;
	int i = 0;
	for (i = 0; i < iIndent; i++)
	{
		szIndent += L"\t";
	}

	return szIndent;
}

wstring tagopen(_In_ wstring szTag, int iIndent)
{
	return indent(iIndent) + L"<" + szTag + L">";
}

wstring tagclose(_In_ wstring szTag, int iIndent)
{
	return indent(iIndent) + L"</" + szTag + L">\n";
}

wstring cdataopen = L"<![CDATA[";
wstring cdataclose = L"]]>";

Parsing::Parsing(wstring const& szParsing, bool bXMLSafe, Attributes const& attributes)
{
	m_szParsing = szParsing;
	m_bXMLSafe = bXMLSafe;
	m_attributes = attributes;
}

Parsing::Parsing(Parsing const& other)
{
	m_szParsing = other.m_szParsing;
	m_bXMLSafe = other.m_bXMLSafe;
	m_attributes = other.m_attributes;
}

wstring Parsing::toXML(UINT uidTag, int iIndent)
{
	if (m_szParsing.empty()) return L"";

	wstring szTag = loadstring(uidTag);
	wstring szXML = tagopen(szTag + m_attributes.toXML(), iIndent);

	if (!m_bXMLSafe)
	{
		szXML += cdataopen;
	}

	szXML += ScrubStringForXML(m_szParsing);

	if (!m_bXMLSafe)
	{
		szXML += cdataclose;
	}

	szXML += tagclose(szTag, 0);

	return szXML;
}

wstring Parsing::toString()
{
	wstring cb = m_attributes.GetAttribute(L"cb");
	if (!cb.empty())
	{
		return L"cb: " + cb + L" lpb: " + m_szParsing;
	}

	wstring err = m_attributes.GetAttribute(L"err");
	if (!err.empty())
	{
		return L"Err: " + err + L"=" + m_szParsing;
	}

	return m_szParsing;
}

void Property::AddParsing(Parsing const& mainParsing, Parsing const& altParsing)
{
	m_MainParsing.push_back(mainParsing);
	m_AltParsing.push_back(altParsing);
}

// Only used to add a MV row parsing to the main set, so we only
// need to copy the first entry and no attributes
void Property::AddMVParsing(Property const& Property)
{
	m_MainParsing.push_back(Property.m_MainParsing[0]);
	m_AltParsing.push_back(Property.m_AltParsing[0]);
}

void Property::AddAttribute(wstring const& key, wstring const& value)
{
	m_attributes.AddAttribute(key, value);
}

wstring Property::toXML(int iIndent)
{
	wstring mv = m_attributes.GetAttribute(L"mv");

	if (mv == L"true")
	{
		wstring szValue = loadstring(PropXMLNames[pcPROPVAL].uidName);
		wstring szRow = loadstring(IDS_ROW);

		wstring szXML = tagopen(szValue + m_attributes.toXML(), iIndent) + L"\n";

		size_t iRow = 0;
		for (iRow = 0; iRow < m_MainParsing.size(); iRow++)
		{
			szXML += tagopen(szRow, iIndent + 1) + L"\n";
			szXML += m_MainParsing[iRow].toXML(PropXMLNames[pcPROPVAL].uidName, iIndent + 2);
			szXML += m_AltParsing[iRow].toXML(PropXMLNames[pcPROPVALALT].uidName, iIndent + 2);
			szXML += tagclose(szRow, iIndent + 1);
		}

		return szXML + tagclose(szValue, iIndent);
	}

	return m_MainParsing[0].toXML(PropXMLNames[pcPROPVAL].uidName, iIndent) + 
		m_AltParsing[0].toXML(PropXMLNames[pcPROPVALALT].uidName, iIndent);
}

wstring Property::toString()
{
	return toString(m_MainParsing);
}

wstring Property::toAltString()
{
	return toString(m_AltParsing);
}

wstring Property::toString(vector<Parsing>& parsing)
{
	wstring mv = m_attributes.GetAttribute(L"mv");
	if (mv == L"true")
	{
		size_t iRow = 0;
		wstring szString = m_attributes.GetAttribute(L"count") + L": ";
		for (iRow = 0; iRow < parsing.size(); iRow++)
		{
			if (iRow != 0)
			{
				szString += L"; ";
			}

			szString += parsing[iRow].toString();
		}

		return szString;
	}
	else if (parsing.size() == 1)
	{
		return parsing[0].toString();
	}

	return L"";
}