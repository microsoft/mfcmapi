#include <core/stdafx.h>
#include <core/addin/addinfile.h>
#include <wincrypt.h>
#include <cstring>
#include <limits>

#pragma comment(lib, "Advapi32.lib")

namespace addin
{
	namespace
	{
		_Check_return_ bool IsRangeValid(size_t imageSize, size_t offset, size_t cb)
		{
			return offset <= imageSize && cb <= imageSize - offset;
		}

		_Check_return_ size_t MinSize(size_t first, size_t second) { return first < second ? first : second; }

		template <typename T>
		_Check_return_ bool
		ReadImageValue(_In_reads_bytes_(cbImage) const BYTE* image, size_t cbImage, size_t offset, T& value)
		{
			if (!image || !IsRangeValid(cbImage, offset, sizeof(T))) return false;
			memcpy(&value, image + offset, sizeof(T));
			return true;
		}

		_Check_return_ bool MapRva(
			_In_reads_bytes_(cbImage) const BYTE* image,
			size_t cbImage,
			DWORD rva,
			size_t cbRequired,
			DWORD sizeOfHeaders,
			size_t sectionTable,
			WORD sectionCount,
			_Out_ size_t& fileOffset,
			_Out_ size_t& cbAvailable)
		{
			if (rva < sizeOfHeaders && IsRangeValid(cbImage, rva, cbRequired))
			{
				fileOffset = rva;
				cbAvailable = MinSize(static_cast<size_t>(sizeOfHeaders - rva), cbImage - rva);
				return cbRequired <= cbAvailable;
			}

			for (WORD i = 0; i < sectionCount; ++i)
			{
				IMAGE_SECTION_HEADER section = {};
				const auto offset = sectionTable + static_cast<size_t>(i) * sizeof(section);
				if (!ReadImageValue(image, cbImage, offset, section)) return false;
				if (rva < section.VirtualAddress) continue;

				const auto delta = static_cast<uint64_t>(rva) - section.VirtualAddress;
				if (delta >= section.SizeOfRawData) continue;
				if (cbRequired > section.SizeOfRawData - delta) return false;

				const auto rawOffset = static_cast<uint64_t>(section.PointerToRawData) + delta;
				if (rawOffset > (std::numeric_limits<size_t>::max)()) return false;
				fileOffset = static_cast<size_t>(rawOffset);
				cbAvailable = MinSize(
					static_cast<size_t>(section.SizeOfRawData - delta),
					fileOffset <= cbImage ? cbImage - fileOffset : 0);
				return IsRangeValid(cbImage, fileOffset, cbRequired) && cbRequired <= cbAvailable;
			}

			return false;
		}

		_Check_return_ bool HasExport(
			_In_reads_bytes_(cbImage) const BYTE* image,
			size_t cbImage,
			size_t optionalHeader,
			WORD optionalHeaderSize,
			WORD sectionCount,
			size_t sectionTable,
			WORD magic)
		{
			IMAGE_DATA_DIRECTORY exportDirectory = {};
			DWORD sizeOfHeaders = 0;

			if (magic == IMAGE_NT_OPTIONAL_HDR32_MAGIC)
			{
				IMAGE_OPTIONAL_HEADER32 header = {};
				if (optionalHeaderSize < sizeof(header) || !ReadImageValue(image, cbImage, optionalHeader, header))
					return false;
				if (header.NumberOfRvaAndSizes <= IMAGE_DIRECTORY_ENTRY_EXPORT) return false;
				exportDirectory = header.DataDirectory[IMAGE_DIRECTORY_ENTRY_EXPORT];
				sizeOfHeaders = header.SizeOfHeaders;
			}
			else if (magic == IMAGE_NT_OPTIONAL_HDR64_MAGIC)
			{
				IMAGE_OPTIONAL_HEADER64 header = {};
				if (optionalHeaderSize < sizeof(header) || !ReadImageValue(image, cbImage, optionalHeader, header))
					return false;
				if (header.NumberOfRvaAndSizes <= IMAGE_DIRECTORY_ENTRY_EXPORT) return false;
				exportDirectory = header.DataDirectory[IMAGE_DIRECTORY_ENTRY_EXPORT];
				sizeOfHeaders = header.SizeOfHeaders;
			}
			else
			{
				return false;
			}

			if (!exportDirectory.VirtualAddress || exportDirectory.Size < sizeof(IMAGE_EXPORT_DIRECTORY)) return false;
			size_t exportOffset = 0;
			size_t cbAvailable = 0;
			if (!MapRva(
					image,
					cbImage,
					exportDirectory.VirtualAddress,
					sizeof(IMAGE_EXPORT_DIRECTORY),
					sizeOfHeaders,
					sectionTable,
					sectionCount,
					exportOffset,
					cbAvailable) ||
				exportDirectory.Size > cbAvailable)
				return false;

			IMAGE_EXPORT_DIRECTORY exports = {};
			if (!ReadImageValue(image, cbImage, exportOffset, exports)) return false;
			if (!exports.NumberOfNames) return false;

			const auto nameTableBytes = static_cast<uint64_t>(exports.NumberOfNames) * sizeof(DWORD);
			if (nameTableBytes > (std::numeric_limits<size_t>::max)()) return false;

			size_t namesOffset = 0;
			if (!MapRva(
					image,
					cbImage,
					exports.AddressOfNames,
					static_cast<size_t>(nameTableBytes),
					sizeOfHeaders,
					sectionTable,
					sectionCount,
					namesOffset,
					cbAvailable))
				return false;

			for (DWORD i = 0; i < exports.NumberOfNames; ++i)
			{
				DWORD nameRva = 0;
				if (!ReadImageValue(image, cbImage, namesOffset + static_cast<size_t>(i) * sizeof(DWORD), nameRva))
					return false;

				size_t nameOffset = 0;
				if (!MapRva(
						image, cbImage, nameRva, 1, sizeOfHeaders, sectionTable, sectionCount, nameOffset, cbAvailable))
					return false;

				const auto nameEnd = static_cast<const BYTE*>(memchr(image + nameOffset, 0, cbAvailable));
				if (!nameEnd) return false;
				const auto nameLength = static_cast<size_t>(nameEnd - (image + nameOffset));
				if (nameLength == sizeof("LoadAddIn") - 1 && !memcmp(image + nameOffset, "LoadAddIn", nameLength))
					return true;
			}

			return false;
		}
	} // namespace

	_Check_return_ bool
	IsCompatibleBitness(_In_reads_bytes_(cbImage) const BYTE* image, size_t cbImage, bool processIs64Bit)
	{
		IMAGE_DOS_HEADER dosHeader = {};
		if (!ReadImageValue(image, cbImage, 0, dosHeader) || dosHeader.e_magic != IMAGE_DOS_SIGNATURE ||
			dosHeader.e_lfanew < 0)
			return false;

		const auto ntOffset = static_cast<size_t>(dosHeader.e_lfanew);
		DWORD signature = 0;
		IMAGE_FILE_HEADER fileHeader = {};
		if (!ReadImageValue(image, cbImage, ntOffset, signature) || signature != IMAGE_NT_SIGNATURE ||
			!ReadImageValue(image, cbImage, ntOffset + sizeof(signature), fileHeader) ||
			fileHeader.SizeOfOptionalHeader < sizeof(WORD))
			return false;

		const auto optionalHeader = ntOffset + sizeof(signature) + sizeof(fileHeader);
		WORD magic = 0;
		if (!ReadImageValue(image, cbImage, optionalHeader, magic)) return false;
		if (magic == IMAGE_NT_OPTIONAL_HDR32_MAGIC) return !processIs64Bit;
		if (magic == IMAGE_NT_OPTIONAL_HDR64_MAGIC) return processIs64Bit;
		return false;
	}

	_Check_return_ bool HasLoadAddInExport(_In_reads_bytes_(cbImage) const BYTE* image, size_t cbImage)
	{
		IMAGE_DOS_HEADER dosHeader = {};
		if (!ReadImageValue(image, cbImage, 0, dosHeader) || dosHeader.e_magic != IMAGE_DOS_SIGNATURE ||
			dosHeader.e_lfanew < 0)
			return false;

		const auto ntOffset = static_cast<size_t>(dosHeader.e_lfanew);
		DWORD signature = 0;
		IMAGE_FILE_HEADER fileHeader = {};
		if (!ReadImageValue(image, cbImage, ntOffset, signature) || signature != IMAGE_NT_SIGNATURE ||
			!ReadImageValue(image, cbImage, ntOffset + sizeof(signature), fileHeader) || !fileHeader.NumberOfSections)
			return false;

		const auto optionalHeader = ntOffset + sizeof(signature) + sizeof(fileHeader);
		WORD magic = 0;
		if (!ReadImageValue(image, cbImage, optionalHeader, magic)) return false;
		const auto sectionTable = optionalHeader + fileHeader.SizeOfOptionalHeader;
		if (!IsRangeValid(cbImage, optionalHeader, fileHeader.SizeOfOptionalHeader) ||
			!IsRangeValid(
				cbImage, sectionTable, static_cast<size_t>(fileHeader.NumberOfSections) * sizeof(IMAGE_SECTION_HEADER)))
			return false;

		return HasExport(
			image,
			cbImage,
			optionalHeader,
			fileHeader.SizeOfOptionalHeader,
			fileHeader.NumberOfSections,
			sectionTable,
			magic);
	}

	CAddInFile::~CAddInFile()
	{
		CloseView();
		if (m_file != INVALID_HANDLE_VALUE) CloseHandle(m_file);
	}

	bool CAddInFile::Open(_In_ const std::wstring& path)
	{
		if (m_file != INVALID_HANDLE_VALUE) return false;

		m_file = CreateFileW(
			path.c_str(),
			GENERIC_READ,
			FILE_SHARE_READ,
			nullptr,
			OPEN_EXISTING,
			FILE_ATTRIBUTE_NORMAL | FILE_FLAG_OPEN_REPARSE_POINT,
			nullptr);
		if (m_file == INVALID_HANDLE_VALUE) return false;

		BY_HANDLE_FILE_INFORMATION fileInfo = {};
		if (!GetFileInformationByHandle(m_file, &fileInfo)) return false;
		if (fileInfo.dwFileAttributes & (FILE_ATTRIBUTE_DIRECTORY | FILE_ATTRIBUTE_REPARSE_POINT))
		{
			SetLastError(ERROR_INVALID_DATA);
			return false;
		}

		LARGE_INTEGER fileSize = {};
		if (!GetFileSizeEx(m_file, &fileSize) || fileSize.QuadPart <= 0 ||
			static_cast<uint64_t>(fileSize.QuadPart) > (std::numeric_limits<size_t>::max)())
			return false;

		m_size = static_cast<size_t>(fileSize.QuadPart);
		m_mapping = CreateFileMappingW(m_file, nullptr, PAGE_READONLY, 0, 0, nullptr);
		if (!m_mapping) return false;

		m_image = static_cast<const BYTE*>(MapViewOfFile(m_mapping, FILE_MAP_READ, 0, 0, 0));
		return m_image != nullptr;
	}

	void CAddInFile::CloseView()
	{
		if (m_image) UnmapViewOfFile(m_image);
		if (m_mapping) CloseHandle(m_mapping);
		m_image = nullptr;
		m_mapping = nullptr;
	}

	bool CAddInFile::HasLoadAddInExport() const { return addin::HasLoadAddInExport(m_image, m_size); }

	bool CAddInFile::HasCompatibleBitness() const
	{
		return IsCompatibleBitness(m_image, m_size, sizeof(void*) == sizeof(uint64_t));
	}

	bool CAddInFile::GetSHA256(_Out_ std::wstring& hashText) const
	{
		if (!m_image || !m_size) return false;
		hashText.clear();

		HCRYPTPROV provider = 0;
		if (!CryptAcquireContextW(&provider, nullptr, nullptr, PROV_RSA_AES, CRYPT_VERIFYCONTEXT)) return false;

		HCRYPTHASH hash = 0;
		auto success = CryptCreateHash(provider, CALG_SHA_256, 0, 0, &hash) != FALSE;
		for (size_t offset = 0; success && offset < m_size;)
		{
			const auto cb = static_cast<DWORD>(MinSize(m_size - offset, static_cast<size_t>(MAXDWORD)));
			success = CryptHashData(hash, m_image + offset, cb, 0) != FALSE;
			offset += cb;
		}

		BYTE digest[32] = {};
		DWORD cbDigest = sizeof(digest);
		success = success && CryptGetHashParam(hash, HP_HASHVAL, digest, &cbDigest, 0) && cbDigest == sizeof(digest);
		if (hash) CryptDestroyHash(hash);
		CryptReleaseContext(provider, 0);
		if (!success) return false;

		static constexpr WCHAR digits[] = L"0123456789abcdef";
		hashText.reserve(sizeof(digest) * 2);
		for (const auto byte : digest)
		{
			hashText += digits[byte >> 4];
			hashText += digits[byte & 0x0f];
		}

		return true;
	}
} // namespace addin
