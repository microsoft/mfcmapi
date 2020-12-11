#pragma once

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

	// TODO: Find right balance between this being a dumb "just data" class
	// and having smarts, like named prop mapping
	// Some obvious ideas:
	// Setting ulPropTag could set name/otherName
	// should we hold address book status for prop lookup?
	class mapiRowModel
	{
	public:
		const std::wstring name() { return _name; }
		void name(std::wstring& v) { _name = v; }

		const std::wstring otherName() { return _otherName; }
		void otherName(std::wstring& v) { _otherName = v; }

		const std::wstring tag() { return _tag; }
		void tag(std::wstring& v) { _tag = v; }

		const std::wstring value() { return _value; }
		void value(std::wstring& v) { _value = v; }

		const std::wstring altValue() { return _altValue; }
		void altValue(std::wstring& v) { _altValue = v; }

		const std::wstring smartView() { return _smartView; }
		void smartView(std::wstring& v) { _smartView = v; }

		const std::wstring namedPropName() { return _namedPropName; }
		void namedPropName(std::wstring& v) { _namedPropName = v; }

		const std::wstring namedPropGuid() { return _namedPropGuid; }
		void namedPropGuid(std::wstring& v) { _namedPropGuid = v; }

		const ULONG ulPropTag() { return _ulPropTag; }
		void ulPropTag(ULONG v) { _ulPropTag = v; }

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