#include <core/stdafx.h>
#include <core/mapi/compressedRtf.h>
#include <core/mapi/extraPropTags.h>
#include <core/utility/error.h>
#include <core/utility/strings.h>

namespace mapi
{
	_Check_return_ HRESULT WrapStreamForRTF(
		_In_ LPSTREAM lpCompressedRTFStream,
		bool bUseWrapEx,
		ULONG ulFlags,
		ULONG ulInCodePage,
		ULONG ulOutCodePage,
		_Deref_out_ LPSTREAM* lpUncompressedRTFStream,
		_Out_opt_ ULONG* pulStreamFlags)
	{
		if (pulStreamFlags) *pulStreamFlags = {};
		if (!lpCompressedRTFStream || !lpUncompressedRTFStream) return MAPI_E_INVALID_PARAMETER;
		auto hRes = S_OK;

		if (!bUseWrapEx)
		{
			hRes = WC_MAPI(WrapCompressedRTFStream(lpCompressedRTFStream, ulFlags, lpUncompressedRTFStream));
		}
		else
		{
			RTF_WCSINFO wcsinfo = {};
			RTF_WCSRETINFO retinfo = {};

			retinfo.size = sizeof(RTF_WCSRETINFO);

			wcsinfo.size = sizeof(RTF_WCSINFO);
			wcsinfo.ulFlags = ulFlags;
			wcsinfo.ulInCodePage = ulInCodePage; // Get ulCodePage from PR_INTERNET_CPID on the IMessage
			wcsinfo.ulOutCodePage = ulOutCodePage; // Desired code page for return

			hRes =
				WC_MAPI(WrapCompressedRTFStreamEx(lpCompressedRTFStream, &wcsinfo, lpUncompressedRTFStream, &retinfo));
			if (pulStreamFlags) *pulStreamFlags = retinfo.ulStreamFlags;
		}

		return hRes;
	}

	_Check_return_ HRESULT ValidateCompressedRtf(const std::vector<BYTE>& compressed, ULONG& rawSize)
	{
		rawSize = 0;
		if (compressed.size() > ULONG_MAX) return WC_H(MAPI_E_TOO_BIG);

		// MS-OXRTFCP 2.1.3.1 / 2.1.3.1.1: four little-endian DWORDs:
		// COMPSIZE (total bytes minus four), RAWSIZE, COMPTYPE, and CRC.
		// https://learn.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxrtfcp/a1f226ed-31e3-4565-9d26-af3c6714c93f
		ULONG header[4]{};
		if (compressed.size() < sizeof(header)) return WC_H(MAPI_E_CORRUPT_DATA);
		if (memcpy_s(header, sizeof(header), compressed.data(), sizeof(header)) != 0) return WC_H(E_INVALIDARG);
		if (header[0] != compressed.size() - sizeof(ULONG)) return WC_H(MAPI_E_CORRUPT_DATA);
		constexpr ULONG lzfu = 0x75465A4C;
		constexpr ULONG mela = 0x414C454D;
		if (header[2] != lzfu && header[2] != mela) return WC_H(MAPI_E_CORRUPT_DATA);
		if (header[2] == mela)
		{
			// MELA stores the raw bytes directly and requires a zero CRC.
			if (header[1] != compressed.size() - sizeof(header)) return WC_H(MAPI_E_CORRUPT_DATA);
			if (header[3]) return WC_H(MAPI_E_CORRUPT_DATA);
		}
		else
		{
			// MS-OXRTFCP 2.1.2.2 / 2.1.3.2: CRC covers CONTENTS only,
			// starts at zero, and has no final XOR. This is the bitwise table equivalent.
			// https://learn.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxrtfcp/a395d090-82f6-4ac0-85c4-057678b2c8e8
			ULONG crc{};
			for (size_t i = sizeof(header); i < compressed.size(); ++i)
			{
				crc ^= compressed[i];
				for (auto bit = 0; bit < 8; ++bit)
					crc = (crc >> 1) ^ ((crc & 1) ? 0xEDB88320UL : 0);
			}

			if (crc != header[3]) return WC_H(MAPI_E_CORRUPT_DATA);
		}

		rawSize = header[1];
		return S_OK;
	}

	_Check_return_ HRESULT DecompressRtfHex(const std::wstring& hex, std::vector<BYTE>& rtf)
	{
		rtf.clear();
		const auto compressed = strings::HexStringToBin(hex);
		if (compressed.empty()) return WC_H(E_INVALIDARG);
		return DecompressRtf(compressed, rtf);
	}

	_Check_return_ HRESULT DecompressRtf(
		const std::vector<BYTE>& compressed,
		std::vector<BYTE>& rtf,
		const bool useWrapEx,
		const ULONG flags,
		const ULONG inCodePage,
		const ULONG outCodePage,
		_Out_opt_ ULONG* streamFlags)
	{
		rtf.clear();
		if (streamFlags) *streamFlags = 0;
		ULONG rawSize{};
		auto hRes = ValidateCompressedRtf(compressed, rawSize);
		if (FAILED(hRes)) return hRes;

		std::vector<BYTE> bytes;

		LPSTREAM source{};
		hRes = WC_H(CreateStreamOnHGlobal(nullptr, true, &source));
		if (SUCCEEDED(hRes))
		{
			ULONG written{};
			hRes = WC_MAPI(source->Write(compressed.data(), static_cast<ULONG>(compressed.size()), &written));
			if (SUCCEEDED(hRes) && written != compressed.size()) hRes = WC_H(STG_E_WRITEFAULT);
			if (SUCCEEDED(hRes)) hRes = WC_MAPI(source->Seek({}, STREAM_SEEK_SET, nullptr));
		}

		LPSTREAM decoded{};
		if (SUCCEEDED(hRes))
			hRes = WrapStreamForRTF(source, useWrapEx, flags, inCodePage, outCodePage, &decoded, streamFlags);
		if (SUCCEEDED(hRes) && !decoded) hRes = WC_H(E_UNEXPECTED);
		if (SUCCEEDED(hRes))
		{
			BYTE buffer[16384]{};
			for (;;)
			{
				ULONG read{};
				hRes = WC_H_IGNORE_RET(std::list<HRESULT>{S_FALSE}, decoded->Read(buffer, sizeof(buffer), &read));
				if (FAILED(hRes)) break;
				if (read > sizeof(buffer) || (!useWrapEx && (bytes.size() > rawSize || read > rawSize - bytes.size())))
				{
					hRes = WC_H(MAPI_E_CORRUPT_DATA);
					break;
				}

				if (read > bytes.max_size() - bytes.size())
				{
					hRes = WC_H(MAPI_E_TOO_BIG);
					break;
				}

				bytes.insert(bytes.end(), buffer, buffer + read);
				if (!read) break;
			}

			// WrapEx can return a different native body or code page, so RAWSIZE does not describe its output.
			if (SUCCEEDED(hRes) && !useWrapEx && bytes.size() != rawSize) hRes = WC_H(MAPI_E_CORRUPT_DATA);
			if (SUCCEEDED(hRes))
			{
				rtf = std::move(bytes);
				hRes = S_OK;
			}
		}

		if (decoded) decoded->Release();
		if (source) source->Release();
		return hRes;
	}

	_Check_return_ HRESULT ReadCompressedRtfStream(_In_ LPSTREAM stream, std::vector<BYTE>& compressed)
	{
		compressed.clear();
		STATSTG stat{};
		auto hRes = WC_MAPI(stream->Stat(&stat, STATFLAG_NONAME));
		if (FAILED(hRes)) return hRes;
		if (stat.cbSize.QuadPart > ULONG_MAX) return WC_H(MAPI_E_TOO_BIG);

		std::vector<BYTE> bytes(static_cast<size_t>(stat.cbSize.QuadPart));
		ULONG read{};
		hRes = WC_H_IGNORE_RET(
			std::list<HRESULT>{S_FALSE}, stream->Read(bytes.data(), static_cast<ULONG>(bytes.size()), &read));
		if (SUCCEEDED(hRes) && read != bytes.size()) hRes = WC_H(STG_E_READFAULT);
		if (SUCCEEDED(hRes)) compressed = std::move(bytes);
		return hRes;
	}

	_Check_return_ HRESULT CompressRtf(const std::vector<BYTE>& rtf, std::vector<BYTE>& compressed)
	{
		compressed.clear();
		if (rtf.size() > ULONG_MAX) return WC_H(MAPI_E_TOO_BIG);
		LPSTREAM target{};
		LPSTREAM decoded{};
		auto hRes = WC_H(CreateStreamOnHGlobal(nullptr, true, &target));
		// MAPI_MODIFY wraps a read/write compressed stream. Commit compresses the entire replacement body.
		// https://learn.microsoft.com/en-us/office/client-developer/outlook/mapi/wrapcompressedrtfstream
		// MAPI's empty LZFu output does not round-trip reliably; request the documented MELA form for empty input.
		const auto flags = MAPI_MODIFY | (rtf.empty() ? STORE_UNCOMPRESSED_RTF : 0);
		if (SUCCEEDED(hRes)) hRes = WrapStreamForRTF(target, false, flags, 0, 0, &decoded, nullptr);
		if (SUCCEEDED(hRes) && !decoded) hRes = WC_H(E_UNEXPECTED);
		if (SUCCEEDED(hRes))
		{
			ULONG written{};
			if (!rtf.empty()) hRes = WC_MAPI(decoded->Write(rtf.data(), static_cast<ULONG>(rtf.size()), &written));
			if (SUCCEEDED(hRes) && written != rtf.size()) hRes = WC_H(STG_E_WRITEFAULT);
			if (SUCCEEDED(hRes)) hRes = WC_MAPI(decoded->Commit(STGC_DEFAULT));
		}

		if (SUCCEEDED(hRes)) hRes = WC_MAPI(target->Seek({}, STREAM_SEEK_SET, nullptr));
		std::vector<BYTE> bytes;
		if (SUCCEEDED(hRes)) hRes = ReadCompressedRtfStream(target, bytes);
		if (SUCCEEDED(hRes))
		{
			ULONG rawSize{};
			hRes = ValidateCompressedRtf(bytes, rawSize);
			if (SUCCEEDED(hRes) && rawSize != rtf.size()) hRes = WC_H(MAPI_E_CORRUPT_DATA);
			if (SUCCEEDED(hRes)) compressed = std::move(bytes);
		}

		if (decoded) decoded->Release();
		if (target) target->Release();
		return hRes;
	}

	_Check_return_ HRESULT ReadCompressedRtf(_In_ LPMAPIPROP prop, std::vector<BYTE>& compressed)
	{
		compressed.clear();
		if (!prop) return WC_H(MAPI_E_INVALID_PARAMETER);
		LPSTREAM stream{};
		auto hRes = WC_MAPI(
			prop->OpenProperty(PR_RTF_COMPRESSED, &IID_IStream, STGM_READ, 0, reinterpret_cast<LPUNKNOWN*>(&stream)));
		if (SUCCEEDED(hRes) && !stream) hRes = WC_H(E_UNEXPECTED);
		if (SUCCEEDED(hRes)) hRes = ReadCompressedRtfStream(stream, compressed);
		if (stream) stream->Release();
		return hRes;
	}

	_Check_return_ HRESULT SaveCompressedRtf(_In_ LPMAPIPROP prop, const std::vector<BYTE>& compressed)
	{
		if (!prop) return WC_H(MAPI_E_INVALID_PARAMETER);
		ULONG rawSize{};
		auto hRes = ValidateCompressedRtf(compressed, rawSize);
		if (FAILED(hRes)) return hRes;

		LPSTREAM stream{};
		hRes = WC_MAPI(prop->OpenProperty(
			PR_RTF_COMPRESSED,
			&IID_IStream,
			STGM_READWRITE,
			MAPI_CREATE | MAPI_MODIFY,
			reinterpret_cast<LPUNKNOWN*>(&stream)));
		if (SUCCEEDED(hRes) && !stream) hRes = WC_H(E_UNEXPECTED);
		if (SUCCEEDED(hRes))
		{
			ULONG written{};
			hRes = WC_MAPI(stream->Write(compressed.data(), static_cast<ULONG>(compressed.size()), &written));
			if (SUCCEEDED(hRes) && written != compressed.size()) hRes = WC_H(STG_E_WRITEFAULT);
			ULARGE_INTEGER size{};
			size.QuadPart = compressed.size();
			if (SUCCEEDED(hRes)) hRes = WC_MAPI(stream->SetSize(size));
			if (SUCCEEDED(hRes)) hRes = WC_MAPI(stream->Commit(STGC_DEFAULT));
		}

		if (stream) stream->Release();
		if (SUCCEEDED(hRes)) hRes = WC_MAPI(prop->SaveChanges(KEEP_OPEN_READWRITE));
		return hRes;
	}
} // namespace mapi
