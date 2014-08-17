#pragma once

#pragma warning(disable : 4995)
#include <vector>
#include <string>
using namespace std;

class Attribute
{
public:
	Attribute();
	Attribute(_In_ wstring key, _In_ wstring value);
	Attribute(const Attribute& other);

	bool empty();
	wstring Key();
	wstring Value();

	wstring toXML();

private:
	wstring m_key;
	wstring m_value;
};

class Attributes
{
public:
	void AddAttribute(_In_ wstring key, _In_ wstring value);
	wstring GetAttribute(_In_ wstring key);

	wstring toXML();

private:
	vector<Attribute> m_attributes;
};
