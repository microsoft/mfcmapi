#pragma once

#include "Attributes.h"

class Parsing
{
public:
	Parsing(wstring const& szParsing, bool bXMLSafe, Attributes const& attributes);
	Parsing(Parsing const& other);

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
	void AddParsing(Parsing const& mainParsing, Parsing const& altParsing);
	void AddMVParsing(Property const& Property);

	void AddAttribute(wstring const& key, wstring const& value);

	wstring toXML(int iIndent);
	wstring toString();
	wstring toAltString();

private:
	wstring toString(vector<Parsing>& parsing);

	vector<Parsing> m_MainParsing;
	vector<Parsing> m_AltParsing;
	Attributes m_attributes;
};