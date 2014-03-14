#pragma once

#include "Attributes.h"
#pragma warning(disable : 4995)
#include <vector>
#include <string>
using namespace std;

class Parsing
{
public:
	Parsing(_In_ wstring szParsing, bool bXMLSafe, Attributes attributes);
	Parsing(const Parsing& other);

	wstring toXML(UINT uidTag, int iIndent);
	wstring toString();

private:
	wstring m_szParsing;
	bool m_bXMLSafe;
	Attributes m_attributes;
};

class Property
{
public:
	void AddParsing(Parsing mainParsing, Parsing altParsing);
	void AddMVParsing(Property Property);

	void AddAttribute(_In_ wstring key, _In_ wstring value);

	wstring toXML(int iIndent);
	wstring toString();
	wstring toAltString();

private:
	wstring toString(vector<Parsing> parsing);

	vector<Parsing> m_MainParsing;
	vector<Parsing> m_AltParsing;
	Attributes m_attributes;
};