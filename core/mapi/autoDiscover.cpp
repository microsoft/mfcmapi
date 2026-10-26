#include <core/stdafx.h>
#include <core/mapi/autoDiscover.h>
#include <core/utility/error.h>

namespace mapi
{
	_Check_return_ HRESULT GetAutoDiscoverXML(const std::wstring& address, std::vector<BYTE>& xml)
	{
		xml.clear();
		if (address.empty()) return WC_H(MAPI_E_INVALID_PARAMETER);

		LPSTREAM stream = nullptr;
		auto hRes = WC_MAPI(HrGetAutoDiscoverXML(address.c_str(), nullptr, nullptr, 0, &stream));
		if (SUCCEEDED(hRes) && !stream) hRes = WC_H(MAPI_E_CALL_FAILED);
		if (SUCCEEDED(hRes))
		{
			LARGE_INTEGER beginning{};
			hRes = WC_MAPI(stream->Seek(beginning, STREAM_SEEK_SET, nullptr));
			while (SUCCEEDED(hRes))
			{
				BYTE buffer[4096]{};
				ULONG read{};
				hRes = WC_H_IGNORE_RET(std::list<HRESULT>{S_FALSE}, stream->Read(buffer, sizeof(buffer), &read));
				if (FAILED(hRes)) break;

				xml.insert(xml.end(), buffer, buffer + read);
				if (hRes == S_FALSE || !read)
				{
					hRes = S_OK;
					break;
				}
			}
		}

		if (stream) stream->Release();
		if (FAILED(hRes)) xml.clear();

		return hRes;
	}

	_Check_return_ HRESULT AutoDiscoverXMLToText(const std::vector<BYTE>& xml, std::wstring& text)
	{
		text.clear();
		if (xml.empty()) return WC_H(MAPI_E_NOT_FOUND);

		const size_t offset = xml.size() >= 3 && xml[0] == 0xef && xml[1] == 0xbb && xml[2] == 0xbf ? 3 : 0;
		if (xml.size() == offset) return WC_H(MAPI_E_NOT_FOUND);
		if (xml.size() - offset > INT_MAX) return WC_H(MAPI_E_TOO_BIG);
		if (std::find(xml.begin(), xml.end(), BYTE{}) != xml.end())
		{
			return WC_H(HRESULT_FROM_WIN32(ERROR_NO_UNICODE_TRANSLATION));
		}

		const auto data = reinterpret_cast<LPCSTR>(xml.data() + offset);
		const auto size = static_cast<int>(xml.size() - offset);
		const auto count = WC_D(int, MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, data, size, nullptr, 0));
		if (!count) return HRESULT_FROM_WIN32(ERROR_NO_UNICODE_TRANSLATION);

		text.resize(count);
		const auto written =
			WC_D(int, MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, data, size, text.data(), count));
		if (!written)
		{
			text.clear();
			return HRESULT_FROM_WIN32(ERROR_NO_UNICODE_TRANSLATION);
		}

		return S_OK;
	}
} // namespace mapi
