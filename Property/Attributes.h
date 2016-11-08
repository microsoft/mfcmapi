#pragma once

class Attribute
{
public:
	Attribute(const wstring& key, const wstring& value);
	Attribute(Attribute const& other);

	bool empty() const;
	wstring Key() const;
	wstring Value() const;

	wstring toXML() const;

private:
	wstring m_key;
	wstring m_value;
};

class Attributes
{
public:
	void AddAttribute(const wstring& key, const wstring& value);
	wstring GetAttribute(const wstring& key);

	wstring toXML();

private:
	vector<Attribute> m_attributes;
};
