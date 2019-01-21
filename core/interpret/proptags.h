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

	// Function to convert property tags to their names
	PropTagNames PropTagToPropName(ULONG ulPropTag, bool bIsAB);

	void FindTagArrayMatches(
		_In_ ULONG ulTarget,
		bool bIsAB,
		const std::vector<NAME_ARRAY_ENTRY_V2>& MyArray,
		std::vector<ULONG>& ulExacts,
		std::vector<ULONG>& ulPartials);
} // namespace propTags