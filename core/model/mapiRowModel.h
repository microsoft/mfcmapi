#pragma once
#include <core/utility/strings.h>
#include <core/interpret/proptype.h>

namespace model
{
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

		const std::wstring propType() { return proptype::TypeToString(PROP_TYPE(_ulPropTag)); }

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

	_Check_return_ std::vector<std::shared_ptr<model::mapiRowModel>>
	propsToModels(ULONG cValues, const SPropValue* lpPropVals, const LPMAPIPROP lpProp, const bool bIsAB);
	_Check_return_ std::shared_ptr<model::mapiRowModel>
	propToModel(const SPropValue* lpPropVal, const ULONG ulPropTag, const LPMAPIPROP lpProp, const bool bIsAB);
} // namespace model