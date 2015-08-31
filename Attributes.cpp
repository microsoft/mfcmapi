#include "StdAfx.h"
#include "Attributes.h"

Attribute::Attribute()
{
}

Attribute::Attribute(wstring const& key, wstring const& value)
{
	m_key = key;
	m_value = value;
}

Attribute::Attribute(Attribute const& other)
{
	m_key = other.m_key;
	m_value = other.m_value;
}

bool Attribute::empty()
{
	return m_key.empty() && m_value.empty();
}

wstring Attribute::Key()
{
	return m_key;
}

wstring Attribute::Value()
{
	return m_value;
}

wstring Attribute::toXML()
{
	return m_key + L"=\"" + m_value + L"\" ";
}

void Attributes::AddAttribute(wstring const& key, wstring const& value)
{
	m_attributes.push_back(Attribute(key, value));
}

wstring Attributes::GetAttribute(wstring const& key)
{
	size_t iAttribute = 0;
	for (iAttribute = 0; iAttribute < m_attributes.size(); iAttribute++)
	{
		if (m_attributes[iAttribute].Key() == key)
		{
			return m_attributes[iAttribute].Value();
		}
	}

	return L"";
}

wstring Attributes::toXML()
{
	size_t iCount = m_attributes.size();

	if (iCount > 0)
	{
		wstring szXML = L" ";

		size_t iAttribute = 0;
		for (iAttribute = 0; iAttribute < iCount; iAttribute++)
		{
			szXML += m_attributes[iAttribute].toXML();
		}

		return szXML;
	}

	return L"";
}