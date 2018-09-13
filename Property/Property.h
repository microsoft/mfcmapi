#pragma once
#include <Property/Attributes.h>

namespace property
{
	class Parsing
	{
	public:
		Parsing(const std::wstring& szParsing, bool bXMLSafe, Attributes const& attributes);
		Parsing(Parsing const& other);

		std::wstring toXML(UINT uidTag, int iIndent) const;
		std::wstring toString() const;

	private:
		std::wstring m_szParsing;
		bool m_bXMLSafe;
		Attributes m_attributes;
	};

	class Property
	{
	public:
		void AddParsing(const Parsing& mainParsing, const Parsing& altParsing);
		void AddMVParsing(const Property& Property);

		void AddAttribute(const std::wstring& key, const std::wstring& value);

		std::wstring toXML(int iIndent) const;
		std::wstring toString() const;
		std::wstring toAltString() const;

	private:
		std::wstring toString(const std::vector<Parsing>& parsings) const;

		std::vector<Parsing> m_MainParsing;
		std::vector<Parsing> m_AltParsing;
		Attributes m_attributes;
	};
} // namespace property