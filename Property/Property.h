#pragma once
#include "Attributes.h"

class Parsing
{
public:
	Parsing(const wstring& szParsing, bool bXMLSafe, Attributes const& attributes);
	Parsing(Parsing const& other);

	wstring toXML(UINT uidTag, int iIndent) const;
	wstring toString() const;

private:
	wstring m_szParsing;
	bool m_bXMLSafe;
	Attributes m_attributes;
};

class Property
{
public:
	void AddParsing(const Parsing& mainParsing, const Parsing& altParsing);
	void AddMVParsing(const Property& Property);

	void AddAttribute(const wstring& key, const wstring& value);

	wstring toXML(int iIndent) const;
	wstring toString() const;
	wstring toAltString() const;

private:
	wstring toString(const vector<Parsing>& parsing) const;

	vector<Parsing> m_MainParsing;
	vector<Parsing> m_AltParsing;
	Attributes m_attributes;
};