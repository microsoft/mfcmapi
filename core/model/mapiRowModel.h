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
		void name(_In_ const std::wstring& v) { _name = v; }

		const std::wstring otherName() { return _otherName; }
		void otherName(_In_ const std::wstring& v) { _otherName = v; }

		const std::wstring tag() { return strings::format(L"0x%08X", _ulPropTag); }

		const std::wstring value() { return _value; }
		void value(_In_ const std::wstring& v) { _value = v; }

		const std::wstring altValue() { return _altValue; }
		void altValue(_In_ const std::wstring& v) { _altValue = v; }

		const std::wstring smartView() { return _smartView; }
		void smartView(_In_ const std::wstring& v) { _smartView = v; }

		const std::wstring namedPropName() { return _namedPropName; }
		void namedPropName(_In_ const std::wstring& v) { _namedPropName = v; }

		const std::wstring namedPropGuid() { return _namedPropGuid; }
		void namedPropGuid(_In_ const std::wstring& v) { _namedPropGuid = v; }

		const ULONG ulPropTag() noexcept { return _ulPropTag; }
		void ulPropTag(_In_ const ULONG v) noexcept { _ulPropTag = v; }

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
		ULONG _ulPropTag{};
	};

	_Check_return_ std::vector<std::shared_ptr<model::mapiRowModel>> propsToModels(
		_In_ ULONG cValues,
		_In_ const SPropValue* lpPropVals,
		_In_ const LPMAPIPROP lpProp,
		_In_ const bool bIsAB);
	_Check_return_ std::shared_ptr<model::mapiRowModel> propToModel(
		_In_ const SPropValue* lpPropVal,
		_In_ const ULONG ulPropTag,
		_In_ const LPMAPIPROP lpProp,
		_In_ const bool bIsAB);
} // namespace model