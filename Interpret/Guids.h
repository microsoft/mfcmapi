#pragma once
// Guid definitions and helper functions

namespace guid
{
	std::wstring GUIDToString(_In_opt_ LPCGUID lpGUID);
	std::wstring GUIDToString(_In_ GUID guid);
	std::wstring GUIDToStringAndName(_In_opt_ LPCGUID lpGUID);
	std::wstring GUIDToStringAndName(_In_ GUID guid);
	LPCGUID GUIDNameToGUID(_In_ const std::wstring& szGUID, bool bByteSwapped);
	_Check_return_ GUID StringToGUID(_In_ const std::wstring& szGUID);
	_Check_return_ GUID StringToGUID(_In_ const std::wstring& szGUID, bool bByteSwapped);
} // namespace guid