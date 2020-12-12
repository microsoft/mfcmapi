#pragma once
#include <core/utility/strings.h>

namespace model
{
	/*
	Name: Single prop tag, named prop name, reg key name - string
	Other Names: multiple prop names - string
	Tag: 8 digit hex prop tag / ulPropTag - number
	Type: PT_* type of the property - number as string
	Value: string representation of value - string
	Value (Alternate): alternate string representation - string
	Smart view: smart view as string - string
	Named prop name: named prop name - either a string or id: * - string
	Named prop GUID: guid string
	Object should expose these strings/numbers (maybe all strings?)
		Only data stored in a row is ulPropTag
		maybe a prop type for icon selection?
	InsertRow returns a sortListData object, which we store ulPropTag in
	*/

	/*
		pcPROPBESTGUESS,
		pcPROPOTHERNAMES,
		pcPROPTAG,
		pcPROPTYPE,
		pcPROPVAL,
		pcPROPVALALT,
		pcPROPSMARTVIEW,
		pcPROPNAMEDNAME,
		pcPROPNAMEDGUID,
	*/

	// TODO: Find right balance between this being a dumb "just data" class
	// and having smarts, like named prop mapping
	// Some obvious ideas:
	// Setting ulPropTag could set name/otherName
	// should we hold address book status for prop lookup?
	class mapiRowModel
	{
	public:
		// Name logic:
		// _name - either a proptag or named prop name
		// ulPropTag as string
		const std::wstring name() { return _name.empty() ? tag() : _name; }
		void name(const std::wstring& v) { _name = v; }

		const std::wstring otherName() { return _otherName; }
		void otherName(const std::wstring& v) { _otherName = v; }

		const std::wstring tag() { return strings::format(L"0x%08X", _ulPropTag); }

		const std::wstring value() { return _value; }
		void value(const std::wstring& v) { _value = v; }

		const std::wstring altValue() { return _altValue; }
		void altValue(const std::wstring& v) { _altValue = v; }

		const std::wstring smartView() { return _smartView; }
		void smartView(const std::wstring& v) { _smartView = v; }

		const std::wstring namedPropName() { return _namedPropName; }
		void namedPropName(const std::wstring& v) { _namedPropName = v; }

		const std::wstring namedPropGuid() { return _namedPropGuid; }
		void namedPropGuid(const std::wstring& v) { _namedPropGuid = v; }

		const ULONG ulPropTag() { return _ulPropTag; }
		void ulPropTag(const ULONG v) { _ulPropTag = v; }

		const ULONG propType() { return PROP_TYPE(_ulPropTag); }

	private:
		std::wstring _name;
		std::wstring _otherName;
		std::wstring _tag;
		std::wstring _value;
		std::wstring _altValue;
		std::wstring _smartView;
		std::wstring _namedPropName;
		std::wstring _namedPropGuid;
		ULONG _ulPropTag;
	};
} // namespace model