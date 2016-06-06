#pragma once

class Attribute
{
public:
	Attribute();
	Attribute(wstring const& key, wstring const& value);
	Attribute(Attribute const& other);

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
	void AddAttribute(wstring const& key, wstring const& value);
	wstring GetAttribute(wstring const& key);

	wstring toXML();

private:
	vector<Attribute> m_attributes;
};
