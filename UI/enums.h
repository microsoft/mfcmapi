#pragma once
// Overload && to provide semantic sugar for testing bits
// Instead of writing something like this:
// bool result = (bool)(flags & flag);
// We can now write:
// bool result = flags && flag;
#define DEFINE_ENUM_FLAG_CONTAINS_OPERATOR(ENUMTYPE) \
	extern "C++" { \
	inline _ENUM_FLAG_CONSTEXPR bool operator&&(ENUMTYPE a, ENUMTYPE b) throw() \
	{ \
		return bool(((_ENUM_FLAG_SIZED_INTEGER<ENUMTYPE>::type) a) & ((_ENUM_FLAG_SIZED_INTEGER<ENUMTYPE>::type) b)); \
	} \
	}

enum class modifyType
{
	REQUEST_MODIFY,
	DO_NOT_REQUEST_MODIFY
};

enum createDialogType
{
	CALL_CREATE_DIALOG,
	DO_NOT_CALL_CREATE_DIALOG
};

enum statusPane
{
	data1,
	data2,
	infoText,
	numPanes
};

// Flags to indicate which contents/hierarchy tables to render
enum class tableDisplayFlags
{
	dfNormal = 0x0000,
	dfAssoc = 0x0001,
	dfDeleted = 0x0002,
};
DEFINE_ENUM_FLAG_OPERATORS(tableDisplayFlags)
DEFINE_ENUM_FLAG_CONTAINS_OPERATOR(tableDisplayFlags)

enum class restrictionType
{
	none,
	normal,
	findrow
};
