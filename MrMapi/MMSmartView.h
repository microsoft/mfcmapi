#pragma once
// Smart View parsing for MrMAPI

#include <cstdio>
#include <functional>
#include <string>
#include <vector>

enum class parserType;

namespace smartviewfile
{
	struct InputOperations
	{
		std::function<__int64(FILE*)> size;
		std::function<size_t(BYTE*, size_t, FILE*)> read;
		std::function<bool(FILE*)> error;
		std::function<std::wstring(SBinary, parserType)> parse;
	};

	const InputOperations& DefaultInputOperations();
	_Check_return_ HRESULT ValidateInputSize(__int64 length, size_t bufferMaxSize, _Out_ size_t& size);
	_Check_return_ HRESULT ParseFile(
		const std::wstring& input,
		const std::wstring& output,
		parserType parser,
		bool binary,
		const InputOperations& operations = DefaultInputOperations());
} // namespace smartviewfile

_Check_return_ HRESULT DoSmartView();