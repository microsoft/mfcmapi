#include "StdAfx.h"
#include "Attributes.h"

Attribute::Attribute()
{
}

Attribute::Attribute(_In_ wstring key, _In_ wstring value)
{
	m_key = key;
	m_value = value;
}

Attribute::Attribute(const Attribute& other)
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

void Attributes::AddAttribute(_In_z_ wstring key, _In_z_ wstring value)
{
	m_attributes.push_back(Attribute(key, value));
}

wstring Attributes::GetAttribute(_In_z_ wstring key)
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
	wstring szXML;
	size_t iCount = m_attributes.size();

	if (iCount > 0)
	{
		szXML += L" ";

		size_t iAttribute = 0;
		for (iAttribute = 0; iAttribute < iCount; iAttribute++)
		{
			szXML += m_attributes[iAttribute].toXML();
		}
	}

	return szXML;
}