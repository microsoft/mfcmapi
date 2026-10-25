#pragma once

namespace mapi
{
	_Check_return_ HRESULT GetAutoDiscoverXML(const std::wstring& address, std::vector<BYTE>& xml);
	_Check_return_ HRESULT AutoDiscoverXMLToText(const std::vector<BYTE>& xml, std::wstring& text);
} // namespace mapi
