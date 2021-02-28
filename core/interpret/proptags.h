#pragma once

struct NAME_ARRAY_ENTRY_V2;

namespace proptags
{
#define PROP_TAG_MASK 0xffff0000

	struct PropTagNames
	{
		std::wstring bestGuess;
		std::wstring otherMatches;
	};

	std::wstring TagToString(ULONG ulPropTag, _In_opt_ LPMAPIPROP lpObj, bool bIsAB, bool bSingleLine);
	std::wstring TagToString(const std::wstring& name, ULONG ulPropTag);

	// Function to convert property tags to their names
	PropTagNames PropTagToPropName(ULONG ulPropTag, bool bIsAB);

	void FindTagArrayMatches(
		_In_ ULONG ulTarget,
		bool bIsAB,
		const std::vector<NAME_ARRAY_ENTRY_V2>& MyArray,
		std::vector<ULONG>& ulExacts,
		std::vector<ULONG>& ulPartials);

	// Strictly does a lookup in the array. Does not convert otherwise
	_Check_return_ ULONG LookupPropName(_In_ const std::wstring& lpszPropName);
	_Check_return_ ULONG PropNameToPropTag(_In_ const std::wstring& lpszPropName);
} // namespace proptags