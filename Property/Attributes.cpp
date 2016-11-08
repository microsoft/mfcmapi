#include "StdAfx.h"
#include "Attributes.h"

Attribute::Attribute(const wstring& key, const wstring& value)
{
	m_key = key;
	m_value = value;
}

Attribute::Attribute(Attribute const& other)
{
	m_key = other.m_key;
	m_value = other.m_value;
}

bool Attribute::empty() const
{
	return m_key.empty() && m_value.empty();
}

wstring Attribute::Key() const
{
	return m_key;
}

wstring Attribute::Value() const
{
	return m_value;
}

wstring Attribute::toXML() const
{
	return m_key + L"=\"" + m_value + L"\" ";
}

void Attributes::AddAttribute(const wstring& key, const wstring& value)
{
	m_attributes.push_back(Attribute(key, value));
}

wstring Attributes::GetAttribute(const wstring& key)
{
	for (auto attribute : m_attributes)
	{
		if (attribute.Key() == key)
		{
			return attribute.Value();
		}
	}

	return L"";
}

wstring Attributes::toXML()
{
	if (!m_attributes.empty())
	{
		wstring szXML = L" ";

		for (auto attribute : m_attributes)
		{
			szXML += attribute.toXML();
		}

		return szXML;
	}

	return L"";
}