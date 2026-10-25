#pragma once

namespace mapi
{
	_Check_return_ HRESULT WrapStreamForRTF(
		_In_ LPSTREAM lpCompressedRTFStream,
		bool bUseWrapEx,
		ULONG ulFlags,
		ULONG ulInCodePage,
		ULONG ulOutCodePage,
		_Deref_out_ LPSTREAM* lpUncompressedRTFStream,
		_Out_opt_ ULONG* pulStreamFlags);

	// Conversion callers must initialize MAPI. Conversion never opens or modifies a property stream.
	// E_INVALIDARG: invalid hex; MAPI_E_CORRUPT_DATA: invalid RTFCP; MAPI_E_TOO_BIG: API size overflow.
	_Check_return_ HRESULT DecompressRtfHex(const std::wstring& hex, std::vector<BYTE>& rtf);
	_Check_return_ HRESULT DecompressRtf(
		const std::vector<BYTE>& compressed,
		std::vector<BYTE>& rtf,
		bool useWrapEx = false,
		ULONG flags = 0,
		ULONG inCodePage = 0,
		ULONG outCodePage = 0,
		_Out_opt_ ULONG* streamFlags = nullptr);
	_Check_return_ HRESULT CompressRtf(const std::vector<BYTE>& rtf, std::vector<BYTE>& compressed);
	_Check_return_ HRESULT ReadCompressedRtf(_In_ LPMAPIPROP prop, std::vector<BYTE>& compressed);
	_Check_return_ HRESULT SaveCompressedRtf(_In_ LPMAPIPROP prop, const std::vector<BYTE>& compressed);
} // namespace mapi
