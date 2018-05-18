#pragma once

class Attribute
{
public:
	Attribute(const std::wstring& key, const std::wstring& value);
	Attribute(Attribute const& other);

	bool empty() const;
	std::wstring Key() const;
	std::wstring Value() const;

	std::wstring toXML() const;

private:
	std::wstring m_key;
	std::wstring m_value;
};

class Attributes
{
public:
	void AddAttribute(const std::wstring& key, const std::wstring& value);
	std::wstring GetAttribute(const std::wstring& key) const;

	std::wstring toXML() const;

private:
	std::vector<Attribute> m_attributes;
};
