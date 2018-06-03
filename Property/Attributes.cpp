#include <StdAfx.h>
#include <Property/Attributes.h>
#include <sstream>

namespace property
{
	Attribute::Attribute(const std::wstring& key, const std::wstring& value)
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

	std::wstring Attribute::Key() const
	{
		return m_key;
	}

	std::wstring Attribute::Value() const
	{
		return m_value;
	}

	std::wstring Attribute::toXML() const
	{
		return m_key + L"=\"" + m_value + L"\" ";
	}

	void Attributes::AddAttribute(const std::wstring& key, const std::wstring& value)
	{
		m_attributes.emplace_back(key, value);
	}

	std::wstring Attributes::GetAttribute(const std::wstring& key) const
	{
		for (const auto& attribute : m_attributes)
		{
			if (attribute.Key() == key)
			{
				return attribute.Value();
			}
		}

		return L"";
	}

	std::wstring Attributes::toXML() const
	{
		if (!m_attributes.empty())
		{
			std::wstringstream szXML;
			szXML << L" ";

			for (const auto& attribute : m_attributes)
			{
				szXML << attribute.toXML();
			}

			return szXML.str();
		}

		return L"";
	}
}