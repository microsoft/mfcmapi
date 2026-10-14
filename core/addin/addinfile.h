#pragma once

#include <string>

namespace addin
{
	class CAddInFile final
	{
	public:
		~CAddInFile();
		CAddInFile() = default;
		CAddInFile(const CAddInFile&) = delete;
		CAddInFile& operator=(const CAddInFile&) = delete;
		CAddInFile(CAddInFile&&) = delete;
		CAddInFile& operator=(CAddInFile&&) = delete;

		_Check_return_ bool Open(_In_ const std::wstring& path);
		void CloseView();
		_Check_return_ bool HasCompatibleBitness() const;
		_Check_return_ bool HasLoadAddInExport() const;
		_Check_return_ bool GetSHA256(_Out_ std::wstring& hashText) const;

	private:
		HANDLE m_file{INVALID_HANDLE_VALUE};
		HANDLE m_mapping{};
		const BYTE* m_image{};
		size_t m_size{};
	};

	_Check_return_ bool
	IsCompatibleBitness(_In_reads_bytes_(cbImage) const BYTE* image, size_t cbImage, bool processIs64Bit);
	_Check_return_ bool HasLoadAddInExport(_In_reads_bytes_(cbImage) const BYTE* image, size_t cbImage);
} // namespace addin
